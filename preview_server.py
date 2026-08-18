"""Local product-preview server with a live UPrinting price endpoint."""

from __future__ import annotations

import json
import hashlib
import os
import re
import socket
import threading
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

import urllib.error
import urllib.request

from uprinting_scraper import DEFAULT_URL, ScraperError, UPrintingScraper, build_export, save_json, save_xlsx

ROOT = Path(__file__).resolve().parent
SCRAPER = UPrintingScraper(DEFAULT_URL)
VARIANT_SCRAPERS: dict[str, UPrintingScraper] = {}
LIVE_SCRAPER_CACHE: dict[str, tuple[UPrintingScraper, dict[str, UPrintingScraper]]] = {}
LIVE_SCRAPER_LOCK = threading.Lock()
OFFLINE_EXPORT: dict | None = None
STATE_LOCK = threading.RLock()
STATE_FILE = ROOT / "active_product.json"
CONFIG_CACHE_FILE = ROOT / "active_product_cache.json"
EXPORT_DIR = ROOT / "exports"
EXPORT_LOCK = threading.Lock()

# Keep integration credentials outside source control. The URL and email have
# harmless local-development defaults; the password is deliberately required.
PRINTOE_API_URL = os.environ.get("PRINTOE_API_URL", "http://localhost:4000/api").rstrip("/")
PRINTOE_ADMIN_EMAIL = os.environ.get("PRINTOE_ADMIN_EMAIL", "demouser@gmail.com")
PRINTOE_ADMIN_PASSWORD = os.environ.get("PRINTOE_ADMIN_PASSWORD", "")
_printoe_token: str | None = None
EXPORT_SCHEMA_VERSION = 4


def _dynamic_rules(
    variant: UPrintingScraper,
    rules: list[dict],
    visible_attr_ids: set[str],
) -> list[dict[str, str]]:
    """Resolve fixed hidden conditions and keep only storefront-visible dependencies."""
    normalized: list[dict[str, str]] = []
    for rule in rules:
        dynamic: dict[str, str] = {}
        impossible = False
        for attr_id, value in rule.items():
            attr_id = str(attr_id)
            value = str(value)
            if attr_id in visible_attr_ids:
                dynamic[attr_id] = value
            elif variant.defaults.get(f"attr{attr_id}") != value:
                impossible = True
                break
        if not impossible:
            normalized.append(dynamic)
    return normalized


def calculator_export_attributes(
    variants: dict[str, UPrintingScraper],
) -> list[dict]:
    """Merge linked calculators while retaining each calculator's dependency rules."""
    merged: list[dict] = []
    for product_id, variant in variants.items():
        attributes = variant.attributes()
        visible_ids = {item["attribute_id"] for item in attributes}
        for attr in attributes:
            attr_id = attr["attribute_id"]
            raw = variant.catalog.get("prod_attrs", {}).get(attr_id, {})
            existing = next((item for item in merged if item["attribute_id"] == attr_id), None)
            if existing is None:
                existing = {
                    **{key: value for key, value in attr.items() if key != "options"},
                    "options": [],
                    "defaults_by_product": {},
                    "hide_rules_by_product": {},
                }
                merged.append(existing)
            option_ids = {option["option_id"] for option in attr["options"]}
            preferred_default = variant.defaults.get(f"attr{attr_id}")
            if preferred_default not in option_ids:
                preferred_default = attr["default_option_id"]
            if preferred_default not in option_ids and attr["options"]:
                preferred_default = attr["options"][0]["option_id"]
            existing["defaults_by_product"][product_id] = preferred_default
            existing["hide_rules_by_product"][product_id] = _dynamic_rules(
                variant,
                raw.get("exceptions", {}).get("-1", []),
                visible_ids,
            )
            for option in attr["options"]:
                option_id = option["option_id"]
                rules = _dynamic_rules(
                    variant,
                    raw.get("exceptions", {}).get(option_id, []),
                    visible_ids,
                )
                # An empty rule has no remaining dynamic conditions and is
                # therefore permanently excluded for this linked calculator.
                if any(not rule for rule in rules):
                    continue
                merged_option = next(
                    (item for item in existing["options"] if item["option_id"] == option_id),
                    None,
                )
                if merged_option is None:
                    merged_option = {
                        **option,
                        "available_product_ids": [],
                        "exclusion_rules_by_product": {},
                    }
                    existing["options"].append(merged_option)
                merged_option["available_product_ids"].append(product_id)
                merged_option["exclusion_rules_by_product"][product_id] = rules
    for attribute in merged:
        attribute["options"].sort(
            key=lambda option: (
                option.get("sort_order") is None,
                option.get("sort_order") or 0,
                option.get("label", ""),
            )
        )
        global_option_ids = {
            option["option_id"] for option in attribute["options"]
        }
        if attribute.get("default_option_id") not in global_option_ids:
            attribute["default_option_id"] = (
                attribute["options"][0]["option_id"]
                if attribute["options"]
                else ""
            )
        for product_id, current in list(
            attribute["defaults_by_product"].items()
        ):
            available = [
                option["option_id"]
                for option in attribute["options"]
                if product_id in option["available_product_ids"]
            ]
            if current not in available:
                attribute["defaults_by_product"][product_id] = (
                    available[0] if available else ""
                )
    merged.sort(
        key=lambda attribute: (
            attribute.get("sort_order") is None,
            attribute.get("sort_order") or 0,
        )
    )
    return merged


def export_fingerprint(scraper: UPrintingScraper, variants: dict[str, UPrintingScraper]) -> str:
    """Stable identity for the currently loaded calculator configuration."""
    payload = {
        "schema_version": EXPORT_SCHEMA_VERSION,
        "source_url": scraper.url,
        "product_id": scraper.product_id,
        "defaults": scraper.defaults,
        "attributes": scraper.attributes(),
        "variants": {
            product_id: {
                "defaults": variant.defaults,
                "attributes": variant.attributes(),
                "exceptions": {
                    str(attr_id): attr.get("exceptions", {})
                    for attr_id, attr in variant.catalog.get("prod_attrs", {}).items()
                },
            }
            for product_id, variant in sorted(variants.items())
        },
    }
    encoded = json.dumps(payload, sort_keys=True, separators=(",", ":"), ensure_ascii=False)
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def export_is_current(path: Path, fingerprint: str) -> bool:
    if not path.exists():
        return False
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data.get("metadata", {}).get("export_fingerprint") == fingerprint
    except (OSError, json.JSONDecodeError, AttributeError):
        return False


def deduplicate_prices(prices: list[dict]) -> list[dict]:
    """Collapse API substitutions that resolve multiple requests to one selection."""
    unique: dict[str, dict] = {}
    for row in prices:
        selection = row.get("selection")
        if not isinstance(selection, dict):
            continue
        key = "&".join(f"{name}={selection[name]}" for name in sorted(selection))
        unique.setdefault(key, row)
    return list(unique.values())


def prune_prices_to_exported_options(data: dict) -> None:
    """Keep only visible selections whose values exist in the exported UI."""
    for attribute in data.get("attributes", []):
        if isinstance(attribute, dict):
            attribute["options"] = [
                option
                for option in attribute.get("options", [])
                if str(option.get("option_id", "")) not in ("", "custom")
            ]
    allowed = {
        f"attr{attribute['attribute_id']}": {
            str(option["option_id"]) for option in attribute.get("options", [])
        }
        for attribute in data.get("attributes", [])
        if isinstance(attribute, dict) and attribute.get("attribute_id") is not None
    }
    retained = []
    for row in data.get("prices", []):
        selection = row.get("selection")
        if not isinstance(selection, dict):
            continue
        visible_selection = {
            key: str(value)
            for key, value in selection.items()
            if key in allowed
        }
        if not visible_selection or any(
            value not in allowed[key]
            for key, value in visible_selection.items()
        ):
            continue
        row["selection"] = visible_selection
        retained.append(row)
    data["prices"] = retained


def validate_printoe_export(data: dict) -> None:
    attributes = data.get("attributes")
    prices = data.get("prices")
    metadata = data.get("metadata")
    if not isinstance(metadata, dict) or not metadata.get("source_url") or not metadata.get("product_id"):
        raise ValueError("Export is missing source product identity.")
    if not isinstance(attributes, list) or not any(item.get("options") for item in attributes if isinstance(item, dict)):
        raise ValueError("Export has no usable product options.")
    if not isinstance(prices, list) or not prices:
        raise ValueError("Export has no valid pricing rows.")
    allowed = {
        f"attr{attribute['attribute_id']}": {
            str(option["option_id"]) for option in attribute.get("options", [])
        }
        for attribute in attributes
        if isinstance(attribute, dict) and attribute.get("attribute_id") is not None
    }
    for index, row in enumerate(prices, 1):
        if not isinstance(row, dict) or not isinstance(row.get("selection"), dict):
            raise ValueError(f"Pricing row {index} has no selection.")
        if any(
            key not in allowed or str(value) not in allowed[key]
            for key, value in row["selection"].items()
        ):
            raise ValueError(
                f"Pricing row {index} references a non-visible option."
            )
        for field in ("price", "unit_price", "quantity"):
            try:
                float(row.get(field))
            except (TypeError, ValueError):
                raise ValueError(f"Pricing row {index} has invalid {field}.") from None


def ensure_export_files() -> tuple[Path, Path]:
    """Builds (or reuses, if already scraped) this product's full pricing
    sweep, returning the json/xlsx paths. Shared by the Download buttons and
    the Upload-to-Printoe flow so there's exactly one place that decides
    when a fresh scrape is needed."""
    safe_code = re.sub(r"[^a-zA-Z0-9_-]+", "-", SCRAPER.catalog.get("product_code") or "product").strip("-").lower()
    base = EXPORT_DIR / f"{safe_code}-{SCRAPER.product_id}"
    json_path = Path(str(base) + ".printoe.json")
    xlsx_path = base.with_suffix(".xlsx")
    fingerprint = export_fingerprint(SCRAPER, VARIANT_SCRAPERS)
    with EXPORT_LOCK:
        # The first download scrapes every finite variation combination;
        # later calls (either format, or an upload) reuse that same export.
        if not export_is_current(json_path, fingerprint) or not xlsx_path.exists():
            prices, errors = [], []
            linked_by_product = {x["product_id"]: x for x in SCRAPER.linked_calculators}
            for product_id, variant in VARIANT_SCRAPERS.items():
                attr_ids = [a["attribute_id"] for a in variant.attributes()]
                # "exhaustive" multiplies every attribute's option count together
                # (thousands of live requests for products with many attributes,
                # taking hours at a polite request rate). "sweep" instead only
                # varies one attribute at a time off the default combo, giving a
                # real scraped price for every individual option value in a
                # fraction of the time - Printoe's import derives each option's
                # price delta from this, so untested joint combinations still
                # price sanely instead of falling back to a flat number.
                variant_prices, variant_errors = variant.scrape_prices("sweep", attr_ids, 50_000, 1, 0.75)
                linked = linked_by_product.get(product_id, {"label": product_id})
                if SCRAPER.linked_calculators and not variant_prices:
                    # A variant that prices zero combinations vanishes from the
                    # merged export with no other trace (its label is dropped
                    # from attr0's options a few lines down) - print so a
                    # broken linked-calculator variant doesn't look identical
                    # to one that simply doesn't exist.
                    print(f"WARNING: linked variant {product_id} ({linked.get('label', product_id)}) produced zero price rows - its type will be missing from the export.")
                if SCRAPER.linked_calculators:
                    for row in variant_prices:
                        row["selection"]["attr0"] = product_id
                        row["display"]["0"] = {"name": SCRAPER.linked_calculators[0]["switch_label"], "label": linked["label"], "option_id": product_id}
                prices.extend(variant_prices); errors.extend(variant_errors)
            data = build_export(SCRAPER, "linked_dependency_pruned_sweep", prices, errors)
            data["attributes"] = calculator_export_attributes(VARIANT_SCRAPERS)
            if len(VARIANT_SCRAPERS) > 1:
                switch = SCRAPER.linked_calculators[0]
                priced_product_ids = {
                    str(row["selection"].get("attr0"))
                    for row in data["prices"]
                    if row["selection"].get("attr0")
                }
                data["attributes"] = [{
                    "attribute_id": "0", "name": switch["switch_label"], "code": "LINKED_CALCULATOR",
                    "field_type": "buttons", "default_option_id": SCRAPER.product_id, "sort_order": 0,
                    "defaults_by_product": {SCRAPER.product_id: SCRAPER.product_id},
                    "hide_rules_by_product": {},
                    "options": [{
                        "option_id": x["product_id"], "source_attr_value_id": x["calc_id"],
                        "label": x["label"], "default": x["product_id"] == SCRAPER.product_id,
                        "sort_order": i + 1, "factors": {"product_id": x["product_id"], "calc_id": x["calc_id"]},
                        "available_product_ids": [],
                        "exclusion_rules_by_product": {},
                    } for i, x in enumerate(SCRAPER.linked_calculators) if x["product_id"] in priced_product_ids],
                }] + data["attributes"]
                data["default_selection"]["attr0"] = SCRAPER.product_id
            for attribute in data["attributes"]:
                key = f"attr{attribute['attribute_id']}"
                default = attribute.get("defaults_by_product", {}).get(
                    SCRAPER.product_id,
                    attribute.get("default_option_id"),
                )
                option_ids = {
                    str(option["option_id"])
                    for option in attribute.get("options", [])
                }
                if str(default) in option_ids:
                    data["default_selection"][key] = str(default)
                else:
                    data["default_selection"].pop(key, None)
            prune_prices_to_exported_options(data)
            data["prices"] = deduplicate_prices(data["prices"])
            data["metadata"]["valid_price_rows"] = len(data["prices"])
            data["metadata"]["schema_version"] = EXPORT_SCHEMA_VERSION
            data["metadata"]["export_fingerprint"] = fingerprint
            validate_printoe_export(data)
            save_json(data, json_path)
            save_xlsx(data, xlsx_path)
    return json_path, xlsx_path


def _printoe_request(path: str, payload: dict, token: str | None = None) -> dict:
    body = json.dumps(payload).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    request = urllib.request.Request(f"{PRINTOE_API_URL}{path}", data=body, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        try:
            message = json.loads(detail).get("message", detail)
        except Exception:
            message = detail
        raise RuntimeError(f"Printoe API error ({exc.code}): {message}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Could not reach Printoe backend at {PRINTOE_API_URL}: {exc}") from exc


def printoe_admin_token() -> str:
    """Logs in once and reuses the token (JWT_EXPIRES_IN=7d server-side) until a call rejects it."""
    global _printoe_token
    if _printoe_token:
        return _printoe_token
    if not PRINTOE_ADMIN_PASSWORD:
        raise RuntimeError(
            "PRINTOE_ADMIN_PASSWORD is not configured. Set it in the preview server environment, then restart."
        )
    result = _printoe_request("/auth/admin/login", {"email": PRINTOE_ADMIN_EMAIL, "password": PRINTOE_ADMIN_PASSWORD})
    _printoe_token = result["data"]["accessToken"]
    return _printoe_token


def upload_to_printoe() -> dict:
    """Ensures this product's full sweep exists, then pushes it straight into
    Printoe as a draft product via /admin/products/import-scrape."""
    json_path, _ = ensure_export_files()
    data = json.loads(json_path.read_text(encoding="utf-8"))
    validate_printoe_export(data)
    token = printoe_admin_token()
    try:
        return _printoe_request("/admin/products/import-scrape", data, token)
    except RuntimeError as exc:
        if "401" in str(exc) or "403" in str(exc):
            global _printoe_token
            _printoe_token = None
            token = printoe_admin_token()
            return _printoe_request("/admin/products/import-scrape", data, token)
        raise


class PreviewHTTPServer(ThreadingHTTPServer):
    """Prevent two preview instances from silently sharing port 8877 on Windows."""
    daemon_threads = True
    allow_reuse_address = False

    def server_bind(self) -> None:
        if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()


def scraper_cache(scraper: UPrintingScraper) -> dict:
    return {
        "product_id": scraper.product_id, "defaults": scraper.defaults,
        "visible_attr_ids": scraper.visible_attr_ids, "catalog": scraper.catalog,
        "product_image": scraper.product_image,
        "price_options": scraper.price_options,
        "page_title": scraper.page_title,
        "hidden_attr_ids": sorted(scraper.hidden_attr_ids),
        "hidden_value_ids": {k: sorted(v) for k, v in scraper.hidden_value_ids.items()},
    }


def scraper_from_cache(parent: UPrintingScraper, data: dict) -> UPrintingScraper:
    item = UPrintingScraper(parent.url, parent.timeout)
    item.api_url, item.auth = parent.api_url, parent.auth
    for key in ("product_id", "defaults", "visible_attr_ids", "catalog", "product_image", "price_options", "page_title"):
        if key in data:
            setattr(item, key, data[key])
    item.hidden_attr_ids = set(data.get("hidden_attr_ids", []))
    item.hidden_value_ids = {k: set(v) for k, v in data.get("hidden_value_ids", {}).items()}
    return item


def migrate_cached_pricing(scraper: UPrintingScraper) -> None:
    """Bring legacy preview caches in line with the current storefront calculator."""
    if "calculator.uprinting.com" in scraper.api_url:
        scraper.api_url = scraper.api_url.replace(
            "calculator.uprinting.com", "calculator.digitalroom.com", 1
        )
    if not scraper.price_options:
        visible = set(map(str, scraper.visible_attr_ids))
        calc_attrs = [
            str(attr_id) for attr_id in scraper.catalog.get("prod_attrs", {})
            if str(attr_id) not in visible and f"attr{attr_id}" in scraper.defaults
        ]
        if calc_attrs:
            scraper.price_options = {
                "calc_attrs_option": 1,
                "calc_attrs": calc_attrs,
                "use_default": "y",
                "override_invalid_spec": True,
                "get_shipping_base_price": True,
            }


def _rule_matches(rule: dict, selection: dict[str, str]) -> bool:
    return all(selection.get(f"attr{attr_id}") == str(value_id) for attr_id, value_id in rule.items())


def normalize_scraper_selection(
    scraper: UPrintingScraper,
    selection: dict[str, str],
    protected_attr: str = "",
) -> dict[str, str]:
    """Apply the same exclusion rules used by UPrinting's shared calculator UI."""
    normalized = dict(scraper.defaults)
    normalized.update(selection)
    attributes = {item["attribute_id"]: item for item in scraper.attributes()}
    raw_attrs = scraper.catalog.get("prod_attrs", {})
    for _ in range(5):
        changed = False
        for attr_id, attr in attributes.items():
            if attr_id == protected_attr:
                continue
            raw_attr = raw_attrs.get(attr_id, {})
            exceptions = raw_attr.get("exceptions", {})
            available = [
                option for option in attr["options"]
                if not any(_rule_matches(rule, normalized) for rule in exceptions.get(option["option_id"], []))
            ]
            current = normalized.get(f"attr{attr_id}")
            # The pricing response can mark a value as hidden because it is
            # only an automatic compatibility fallback, not a user-selectable
            # choice. If every visible value is invalid (for example two-sided
            # printing with "No Folding"), still let normalization choose that
            # hidden-but-valid catalog value exactly as UPrinting does.
            if not available:
                raw_values = raw_attr.get("prod_attr_vals", {})
                hidden_fallbacks = [
                    {
                        "option_id": str(option_id),
                        "sort_order": value.get("sort_order"),
                    }
                    for option_id, value in raw_values.items()
                    if value.get("hide_attribute_value_flag") != "y"
                    and value.get("custom_flag") != "y"
                    and not any(
                        _rule_matches(rule, normalized)
                        for rule in exceptions.get(str(option_id), [])
                    )
                ]
                hidden_fallbacks.sort(
                    key=lambda option: (
                        option["sort_order"] is None,
                        option["sort_order"] or 0,
                    )
                )
                available = hidden_fallbacks
            if available and not any(option["option_id"] == current for option in available):
                fallback = next((o for o in available if o["option_id"] == attr["default_option_id"]), available[0])
                normalized[f"attr{attr_id}"] = fallback["option_id"]
                changed = True
        if not changed:
            break
    return normalized


def normalize_selection(selection: dict[str, str], protected_attr: str = "") -> dict[str, str]:
    return normalize_scraper_selection(SCRAPER, selection, protected_attr)


def load_variants(scraper: UPrintingScraper) -> dict[str, UPrintingScraper]:
    variants = {scraper.product_id: scraper}
    for linked in scraper.linked_calculators:
        if linked["product_id"] != scraper.product_id:
            variants[linked["product_id"]] = scraper.linked_scraper(linked)
    return variants


def live_scrapers(source_url: str) -> tuple[UPrintingScraper, dict[str, UPrintingScraper]]:
    parsed = urlparse(source_url)
    hostname = (parsed.hostname or "").lower()
    if (
        parsed.scheme != "https"
        or not (
            hostname == "uprinting.com"
            or hostname.endswith(".uprinting.com")
        )
    ):
        raise ValueError("Only HTTPS UPrinting product URLs are allowed")
    with LIVE_SCRAPER_LOCK:
        cached = LIVE_SCRAPER_CACHE.get(source_url)
        if cached:
            return cached
        scraper = UPrintingScraper(source_url)
        scraper.load()
        result = (scraper, load_variants(scraper))
        LIVE_SCRAPER_CACHE[source_url] = result
        return result


def live_price_payload(
    scraper: UPrintingScraper,
    variants: dict[str, UPrintingScraper],
    body: dict,
) -> dict:
    selection = body.get("selection", {})
    if not isinstance(selection, dict):
        raise ValueError("selection must be an object")
    raw_selection = {
        str(key): str(value)
        for key, value in selection.items()
        if str(key).startswith("attr") and str(key)[4:].isdigit()
    }
    requested_product = str(
        raw_selection.get("attr0") or body.get("product_id") or scraper.product_id
    )
    price_scraper = variants.get(requested_product)
    if price_scraper is None:
        raise ValueError("Requested linked product is unavailable")
    protected = str(body.get("changed_attribute_id", ""))
    safe_selection = {
        key: value for key, value in raw_selection.items() if key != "attr0"
    }
    safe_selection = normalize_scraper_selection(
        price_scraper,
        safe_selection,
        protected,
    )
    result = price_scraper.price(safe_selection)
    effective_price = result.get("discounted_price")
    if effective_price is None:
        effective_price = result.get("price")
    quantity = result.get("qty")
    try:
        effective_unit_price = (
            float(effective_price) / float(quantity)
            if float(quantity)
            else result.get("unit_price")
        )
    except (TypeError, ValueError, ZeroDivisionError):
        effective_unit_price = result.get(
            "discounted_unit_price",
            result.get("unit_price"),
        )
    return {
        "ok": True,
        "price": effective_price,
        "unit_price": effective_unit_price,
        "quantity": quantity,
        "turnaround_days": result.get("turnaround"),
        "display_specs": result.get("display_specs", []),
        "price_data": result.get("price_data", {}),
        "normalized_selection": safe_selection,
        "product_id": requested_product,
    }


class PreviewHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def end_headers(self) -> None:
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/health":
            self._json(200, {"ok": True})
            return
        if parsed.path == "/api/export":
            try:
                from urllib.parse import parse_qs
                requested = parse_qs(parsed.query).get("format", ["json"])[0]
                if requested not in ("json", "xlsx"):
                    raise ValueError("format must be json or xlsx")
                json_path, xlsx_path = ensure_export_files()
                target = json_path if requested == "json" else xlsx_path
                payload = target.read_bytes()
                self.send_response(200)
                self.send_header("Content-Type", "application/json" if requested == "json" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", f'attachment; filename="{target.name}"')
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
            except Exception as exc:
                self._json(500, {"ok": False, "error": f"Export failed: {exc}"})
            return
        if parsed.path == "/api/config":
            if OFFLINE_EXPORT is not None:
                metadata = OFFLINE_EXPORT["metadata"]
                visible_keys = {f"attr{item['attribute_id']}" for item in OFFLINE_EXPORT["attributes"]}
                offline_defaults = {
                    key: value for key, value in OFFLINE_EXPORT["default_selection"].items()
                    if key in visible_keys
                }
                self._json(200, {
                    "product_id": metadata["product_id"], "product_name": metadata["product_name"],
                    "product_code": metadata.get("product_code"), "source_url": metadata["source_url"],
                    "product_image": "https://staticecp.uprinting.com/7319/600x600/BIC_Sticky_Note_3x3_25_Sheets_Marketing_Materials_A.jpg",
                    "default_selection": offline_defaults, "attributes": OFFLINE_EXPORT["attributes"],
                    "description": OFFLINE_EXPORT.get("description", ""),
                    "images": OFFLINE_EXPORT.get("images", []), "video": OFFLINE_EXPORT.get("video", ""),
                    "linked_switch": None, "variants": {},
                }); return
            attrs = []
            for attr in SCRAPER.attributes():
                raw = SCRAPER.catalog.get("prod_attrs", {}).get(attr["attribute_id"], {})
                item = dict(attr)
                item["exceptions"] = raw.get("exceptions", {})
                attrs.append(item)
            linked = SCRAPER.linked_calculators
            variants = {}
            for item in linked:
                variant = VARIANT_SCRAPERS.get(item["product_id"])
                if variant:
                    variants[item["product_id"]] = {
                        "product_id": variant.product_id,
                        "product_name": variant.catalog.get("product_name"),
                        "default_selection": variant.defaults,
                        "attributes": [dict(a, exceptions=variant.catalog.get("prod_attrs", {}).get(a["attribute_id"], {}).get("exceptions", {})) for a in variant.attributes()],
                    }
            self._json(200, {
                "product_id": SCRAPER.product_id,
                "product_name": SCRAPER.page_title or SCRAPER.catalog.get("product_name"),
                "product_code": SCRAPER.catalog.get("product_code"),
                "source_url": SCRAPER.url,
                "product_image": SCRAPER.product_image,
                "default_selection": SCRAPER.defaults,
                "description": SCRAPER.description,
                "images": SCRAPER.images, "video": SCRAPER.video,
                "attributes": attrs,
                "linked_switch": ({"name": linked[0]["switch_label"], "options": [{"label": x["label"], "product_id": x["product_id"]} for x in linked]} if len(linked) > 1 else None),
                "variants": variants,
            })
            return
        super().do_GET()

    def do_POST(self) -> None:
        global SCRAPER, VARIANT_SCRAPERS, OFFLINE_EXPORT
        request_path = urlparse(self.path).path
        if request_path == "/api/load-product":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                body = json.loads(self.rfile.read(length) or b"{}")
                url = str(body.get("url", "")).strip()
                candidate = UPrintingScraper(url)
                candidate.load()
                if not candidate.attributes():
                    raise ValueError("Is product par calculator attributes nahi mile.")
                with STATE_LOCK:
                    SCRAPER = candidate
                    VARIANT_SCRAPERS = load_variants(candidate)
                    LIVE_SCRAPER_CACHE[candidate.url] = (
                        candidate,
                        VARIANT_SCRAPERS,
                    )
                    OFFLINE_EXPORT = None
                    STATE_FILE.write_text(json.dumps({"url": url}, indent=2), encoding="utf-8")
                    CONFIG_CACHE_FILE.write_text(json.dumps({
                        "url": candidate.url, "product_id": candidate.product_id,
                        "api_url": candidate.api_url, "auth": candidate.auth,
                        "defaults": candidate.defaults, "visible_attr_ids": candidate.visible_attr_ids,
                        "catalog": candidate.catalog, "product_image": candidate.product_image,
                        "linked_calculators": candidate.linked_calculators,
                        "price_options": candidate.price_options,
                        "description": candidate.description,
                        "images": candidate.images, "video": candidate.video,
                        "page_title": candidate.page_title,
                        "hidden_attr_ids": sorted(candidate.hidden_attr_ids),
                        "hidden_value_ids": {k: sorted(v) for k, v in candidate.hidden_value_ids.items()},
                        "variants": {key: scraper_cache(value) for key, value in VARIANT_SCRAPERS.items()},
                    }), encoding="utf-8")
                self._json(200, {
                    "ok": True, "product_id": candidate.product_id,
                    "product_name": candidate.page_title or candidate.catalog.get("product_name"),
                    "attributes": len(candidate.attributes()),
                    "redirect": "/index.html?product=" + candidate.product_id,
                })
            except Exception as exc:
                self._json(400, {"ok": False, "error": str(exc)})
            return
        if request_path == "/api/upload-to-printoe":
            try:
                result = upload_to_printoe()
                self._json(200, {"ok": True, **result})
            except Exception as exc:
                self._json(502, {"ok": False, "error": str(exc)})
            return
        if request_path not in ("/api/price", "/api/live-price"):
            self.send_error(404)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = json.loads(self.rfile.read(length) or b"{}")
            if request_path == "/api/live-price":
                source_url = str(body.get("source_url", "")).strip()
                if not source_url:
                    raise ValueError("source_url is required")
                scraper, variants = live_scrapers(source_url)
                self._json(200, live_price_payload(scraper, variants, body))
                return
            selection = body.get("selection", {})
            if not isinstance(selection, dict):
                raise ValueError("selection must be an object")
            raw_selection = {
                str(key): str(value)
                for key, value in selection.items()
                if str(key).startswith("attr") and str(key)[4:].isdigit()
            }
            if OFFLINE_EXPORT is not None:
                visible_keys = {f"attr{item['attribute_id']}" for item in OFFLINE_EXPORT["attributes"]}
                offline_selection = {key: value for key, value in raw_selection.items() if key in visible_keys}
                row = next((item for item in OFFLINE_EXPORT["prices"] if item["selection"] == offline_selection), None)
                if row is None:
                    raise ValueError("This option combination is unavailable")
                self._json(200, {"ok": True, "price": row["price"], "unit_price": row["unit_price"],
                    "quantity": row["quantity"], "turnaround_days": row.get("turnaround_days"),
                    "display_specs": [], "price_data": {}, "normalized_selection": offline_selection}); return
            self._json(200, live_price_payload(SCRAPER, VARIANT_SCRAPERS, body))
        except (ValueError, ScraperError, OSError) as exc:
            self._json(400, {"ok": False, "error": str(exc)})
        except Exception as exc:
            self._json(502, {"ok": False, "error": f"Live price unavailable: {exc}"})

    def _json(self, status: int, payload: dict) -> None:
        data = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def main() -> None:
    global SCRAPER, VARIANT_SCRAPERS, OFFLINE_EXPORT
    if os.environ.get("PRICING_SERVICE_ONLY") == "1":
        server = PreviewHTTPServer(("0.0.0.0", 8877), PreviewHandler)
        print("Live pricing service ready on port 8877")
        server.serve_forever()
        return
    print("Loading UPrinting product configuration...")
    if STATE_FILE.exists():
        try:
            saved_url = json.loads(STATE_FILE.read_text(encoding="utf-8")).get("url", DEFAULT_URL)
            SCRAPER = UPrintingScraper(saved_url)
        except Exception:
            pass
    if CONFIG_CACHE_FILE.exists():
        try:
            cached = json.loads(CONFIG_CACHE_FILE.read_text(encoding="utf-8"))
            if cached.get("url") == SCRAPER.url:
                for key in ("product_id", "api_url", "auth", "defaults", "visible_attr_ids", "catalog", "product_image", "linked_calculators", "price_options", "description", "images", "video", "page_title"):
                    if key in cached:
                        setattr(SCRAPER, key, cached[key])
                SCRAPER.hidden_attr_ids = set(cached.get("hidden_attr_ids", []))
                SCRAPER.hidden_value_ids = {k: set(v) for k, v in cached.get("hidden_value_ids", {}).items()}
                migrate_cached_pricing(SCRAPER)
                cached_variants = cached.get("variants", {})
                VARIANT_SCRAPERS = {
                    str(key): scraper_from_cache(SCRAPER, value)
                    for key, value in cached_variants.items()
                } or {SCRAPER.product_id: SCRAPER}
                VARIANT_SCRAPERS[SCRAPER.product_id] = SCRAPER
                LIVE_SCRAPER_CACHE[SCRAPER.url] = (
                    SCRAPER,
                    VARIANT_SCRAPERS,
                )
                server = PreviewHTTPServer(("127.0.0.1", 8877), PreviewHandler)
                print(f"Cached active product loaded: {SCRAPER.catalog.get('product_name')}")
                print("Preview ready: http://127.0.0.1:8877/index.html")
                server.serve_forever(); return
        except Exception as exc:
            print(f"Active cache could not be used ({exc}); reloading live configuration.")
    try:
        SCRAPER.load()
        CONFIG_CACHE_FILE.write_text(json.dumps({
            "url": SCRAPER.url, "product_id": SCRAPER.product_id,
            "api_url": SCRAPER.api_url, "auth": SCRAPER.auth,
            "defaults": SCRAPER.defaults, "visible_attr_ids": SCRAPER.visible_attr_ids,
            "catalog": SCRAPER.catalog, "product_image": SCRAPER.product_image,
            "linked_calculators": SCRAPER.linked_calculators,
            "price_options": SCRAPER.price_options,
            "description": SCRAPER.description,
            "images": SCRAPER.images, "video": SCRAPER.video,
            "page_title": SCRAPER.page_title,
            "hidden_attr_ids": sorted(SCRAPER.hidden_attr_ids),
            "hidden_value_ids": {k: sorted(v) for k, v in SCRAPER.hidden_value_ids.items()},
        }), encoding="utf-8")
    except Exception as exc:
        export_path = EXPORT_DIR / "bic3x3stickynotes-2539.printoe.json"
        if export_path.exists():
            candidate = json.loads(export_path.read_text(encoding="utf-8"))
            if "sticky-notepad-3x3" in SCRAPER.url:
                OFFLINE_EXPORT = candidate
                print(f"Live reload failed ({exc}); offline exact-price export loaded.")
        if not CONFIG_CACHE_FILE.exists():
            if OFFLINE_EXPORT is None: raise
            server = PreviewHTTPServer(("127.0.0.1", 8877), PreviewHandler)
            print("Preview ready: http://127.0.0.1:8877/index.html")
            server.serve_forever(); return
        if OFFLINE_EXPORT is None:
            cached = json.loads(CONFIG_CACHE_FILE.read_text(encoding="utf-8"))
            SCRAPER = UPrintingScraper(cached["url"])
            for key in ("product_id", "api_url", "auth", "defaults", "visible_attr_ids", "catalog", "product_image", "linked_calculators", "price_options", "description", "images", "video", "page_title"):
                if key in cached:
                    setattr(SCRAPER, key, cached[key])
            SCRAPER.hidden_attr_ids = set(cached.get("hidden_attr_ids", []))
            SCRAPER.hidden_value_ids = {k: set(v) for k, v in cached.get("hidden_value_ids", {}).items()}
            migrate_cached_pricing(SCRAPER)
            print(f"Live reload failed ({exc}); cached product configuration loaded.")
    VARIANT_SCRAPERS = load_variants(SCRAPER)
    LIVE_SCRAPER_CACHE[SCRAPER.url] = (SCRAPER, VARIANT_SCRAPERS)
    server = PreviewHTTPServer(("127.0.0.1", 8877), PreviewHandler)
    print("Preview ready: http://127.0.0.1:8877/index.html")
    server.serve_forever()


if __name__ == "__main__":
    main()

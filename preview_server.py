"""Local product-preview server with a live UPrinting price endpoint."""

from __future__ import annotations

import json
import re
import socket
import threading
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from uprinting_scraper import DEFAULT_URL, ScraperError, UPrintingScraper, build_export, save_json, save_xlsx

ROOT = Path(__file__).resolve().parent
SCRAPER = UPrintingScraper(DEFAULT_URL)
VARIANT_SCRAPERS: dict[str, UPrintingScraper] = {}
OFFLINE_EXPORT: dict | None = None
STATE_LOCK = threading.RLock()
STATE_FILE = ROOT / "active_product.json"
CONFIG_CACHE_FILE = ROOT / "active_product_cache.json"
EXPORT_DIR = ROOT / "exports"
EXPORT_LOCK = threading.Lock()


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
    }


def scraper_from_cache(parent: UPrintingScraper, data: dict) -> UPrintingScraper:
    item = UPrintingScraper(parent.url, parent.timeout)
    item.api_url, item.auth = parent.api_url, parent.auth
    for key in ("product_id", "defaults", "visible_attr_ids", "catalog", "product_image", "price_options"):
        if key in data:
            setattr(item, key, data[key])
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


def normalize_selection(selection: dict[str, str], protected_attr: str = "") -> dict[str, str]:
    """Apply the same exclusion rules used by UPrinting's shared calculator UI."""
    normalized = dict(SCRAPER.defaults)
    normalized.update(selection)
    attributes = {item["attribute_id"]: item for item in SCRAPER.attributes()}
    raw_attrs = SCRAPER.catalog.get("prod_attrs", {})
    for _ in range(5):
        changed = False
        for attr_id, attr in attributes.items():
            if attr_id == protected_attr:
                continue
            exceptions = raw_attrs.get(attr_id, {}).get("exceptions", {})
            available = [
                option for option in attr["options"]
                if not any(_rule_matches(rule, normalized) for rule in exceptions.get(option["option_id"], []))
            ]
            current = normalized.get(f"attr{attr_id}")
            if available and not any(option["option_id"] == current for option in available):
                fallback = next((o for o in available if o["option_id"] == attr["default_option_id"]), available[0])
                normalized[f"attr{attr_id}"] = fallback["option_id"]
                changed = True
        if not changed:
            break
    return normalized


def load_variants(scraper: UPrintingScraper) -> dict[str, UPrintingScraper]:
    variants = {scraper.product_id: scraper}
    for linked in scraper.linked_calculators:
        if linked["product_id"] != scraper.product_id:
            variants[linked["product_id"]] = scraper.linked_scraper(linked)
    return variants


class PreviewHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def end_headers(self) -> None:
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/export":
            try:
                from urllib.parse import parse_qs
                requested = parse_qs(parsed.query).get("format", ["json"])[0]
                if requested not in ("json", "xlsx"):
                    raise ValueError("format must be json or xlsx")
                safe_code = re.sub(r"[^a-zA-Z0-9_-]+", "-", SCRAPER.catalog.get("product_code") or "product").strip("-").lower()
                base = EXPORT_DIR / f"{safe_code}-{SCRAPER.product_id}"
                json_path = Path(str(base) + ".printoe.json")
                xlsx_path = base.with_suffix(".xlsx")
                with EXPORT_LOCK:
                    # The first download scrapes every finite variation combination;
                    # the second format reuses the same product-specific export.
                    if not json_path.exists() or not xlsx_path.exists():
                        prices, errors = [], []
                        variant_attributes = []
                        linked_by_product = {x["product_id"]: x for x in SCRAPER.linked_calculators}
                        for product_id, variant in VARIANT_SCRAPERS.items():
                            attr_ids = [a["attribute_id"] for a in variant.attributes()]
                            variant_prices, variant_errors = variant.scrape_prices("exhaustive", attr_ids, 50_000, 1, 0.75)
                            linked = linked_by_product.get(product_id, {"label": product_id})
                            for row in variant_prices:
                                row["selection"]["attr0"] = product_id
                                row["display"]["0"] = {"name": SCRAPER.linked_calculators[0]["switch_label"], "label": linked["label"], "option_id": product_id}
                            prices.extend(variant_prices); errors.extend(variant_errors)
                            for attr in variant.attributes():
                                existing = next((x for x in variant_attributes if x["attribute_id"] == attr["attribute_id"]), None)
                                if existing is None:
                                    variant_attributes.append(attr)
                                else:
                                    known = {x["option_id"] for x in existing["options"]}
                                    existing["options"].extend(x for x in attr["options"] if x["option_id"] not in known)
                        data = build_export(SCRAPER, "linked_dependency_pruned_exhaustive", prices, errors)
                        if len(VARIANT_SCRAPERS) > 1:
                            switch = SCRAPER.linked_calculators[0]
                            data["attributes"] = [{
                                "attribute_id": "0", "name": switch["switch_label"], "code": "LINKED_CALCULATOR",
                                "field_type": "buttons", "default_option_id": SCRAPER.product_id, "sort_order": 0,
                                "options": [{"option_id": x["product_id"], "source_attr_value_id": x["calc_id"], "label": x["label"], "default": x["product_id"] == SCRAPER.product_id, "sort_order": i + 1, "factors": {"product_id": x["product_id"], "calc_id": x["calc_id"]}} for i, x in enumerate(SCRAPER.linked_calculators)],
                            }] + variant_attributes
                            data["default_selection"]["attr0"] = SCRAPER.product_id
                            visible_keys = {f"attr{x['attribute_id']}" for x in data["attributes"]}
                            for row in data["prices"]:
                                row["selection"] = {k: v for k, v in row["selection"].items() if k in visible_keys}
                            for attribute in data["attributes"]:
                                key = f"attr{attribute['attribute_id']}"
                                used = {row["selection"].get(key) for row in data["prices"]}
                                attribute["options"] = [option for option in attribute["options"] if option["option_id"] in used]
                        save_json(data, json_path)
                        save_xlsx(data, xlsx_path)
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
                "product_name": SCRAPER.catalog.get("product_name"),
                "product_code": SCRAPER.catalog.get("product_code"),
                "source_url": SCRAPER.url,
                "product_image": SCRAPER.product_image,
                "default_selection": SCRAPER.defaults,
                "attributes": attrs,
                "linked_switch": ({"name": linked[0]["switch_label"], "options": [{"label": x["label"], "product_id": x["product_id"]} for x in linked]} if len(linked) > 1 else None),
                "variants": variants,
            })
            return
        super().do_GET()

    def do_POST(self) -> None:
        global SCRAPER, VARIANT_SCRAPERS, OFFLINE_EXPORT
        if urlparse(self.path).path == "/api/load-product":
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
                    OFFLINE_EXPORT = None
                    STATE_FILE.write_text(json.dumps({"url": url}, indent=2), encoding="utf-8")
                    CONFIG_CACHE_FILE.write_text(json.dumps({
                        "url": candidate.url, "product_id": candidate.product_id,
                        "api_url": candidate.api_url, "auth": candidate.auth,
                        "defaults": candidate.defaults, "visible_attr_ids": candidate.visible_attr_ids,
                        "catalog": candidate.catalog, "product_image": candidate.product_image,
                        "linked_calculators": candidate.linked_calculators,
                        "price_options": candidate.price_options,
                        "variants": {key: scraper_cache(value) for key, value in VARIANT_SCRAPERS.items()},
                    }), encoding="utf-8")
                self._json(200, {
                    "ok": True, "product_id": candidate.product_id,
                    "product_name": candidate.catalog.get("product_name"),
                    "attributes": len(candidate.attributes()),
                    "redirect": "/index.html?product=" + candidate.product_id,
                })
            except Exception as exc:
                self._json(400, {"ok": False, "error": str(exc)})
            return
        if urlparse(self.path).path != "/api/price":
            self.send_error(404)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = json.loads(self.rfile.read(length) or b"{}")
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
            safe_selection = {key: value for key, value in raw_selection.items() if key != "attr0"}
            protected = str(body.get("changed_attribute_id", ""))
            requested_product = str(body.get("product_id", SCRAPER.product_id))
            price_scraper = VARIANT_SCRAPERS.get(requested_product, SCRAPER)
            if price_scraper is SCRAPER:
                safe_selection = normalize_selection(safe_selection, protected)
            else:
                normalized = dict(price_scraper.defaults); normalized.update(safe_selection); safe_selection = normalized
            result = price_scraper.price(safe_selection)
            effective_price = result.get("discounted_price")
            if effective_price is None:
                effective_price = result.get("price")
            quantity = result.get("qty")
            try:
                effective_unit_price = float(effective_price) / float(quantity) if float(quantity) else result.get("unit_price")
            except (TypeError, ValueError, ZeroDivisionError):
                effective_unit_price = result.get("discounted_unit_price", result.get("unit_price"))
            payload = {
                "ok": True,
                "price": effective_price,
                "unit_price": effective_unit_price,
                "quantity": quantity,
                "turnaround_days": result.get("turnaround"),
                "display_specs": result.get("display_specs", []),
                "price_data": result.get("price_data", {}),
                "pricing_debug": {
                    key: value for key, value in result.items()
                    if "price" in str(key).lower() or "discount" in str(key).lower()
                },
                "normalized_selection": safe_selection,
            }
            self._json(200, payload)
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
                for key in ("product_id", "api_url", "auth", "defaults", "visible_attr_ids", "catalog", "product_image", "linked_calculators", "price_options"):
                    if key in cached:
                        setattr(SCRAPER, key, cached[key])
                migrate_cached_pricing(SCRAPER)
                cached_variants = cached.get("variants", {})
                VARIANT_SCRAPERS = {
                    str(key): scraper_from_cache(SCRAPER, value)
                    for key, value in cached_variants.items()
                } or {SCRAPER.product_id: SCRAPER}
                VARIANT_SCRAPERS[SCRAPER.product_id] = SCRAPER
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
            for key in ("product_id", "api_url", "auth", "defaults", "visible_attr_ids", "catalog", "product_image", "linked_calculators", "price_options"):
                if key in cached:
                    setattr(SCRAPER, key, cached[key])
            migrate_cached_pricing(SCRAPER)
            print(f"Live reload failed ({exc}); cached product configuration loaded.")
    VARIANT_SCRAPERS = load_variants(SCRAPER)
    server = PreviewHTTPServer(("127.0.0.1", 8877), PreviewHandler)
    print("Preview ready: http://127.0.0.1:8877/index.html")
    server.serve_forever()


if __name__ == "__main__":
    main()

"""Local product-preview server with a live UPrinting price endpoint."""

from __future__ import annotations

import json
import threading
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from uprinting_scraper import DEFAULT_URL, ScraperError, UPrintingScraper

ROOT = Path(__file__).resolve().parent
SCRAPER = UPrintingScraper(DEFAULT_URL)
STATE_LOCK = threading.RLock()
STATE_FILE = ROOT / "active_product.json"


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


class PreviewHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def end_headers(self) -> None:
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_GET(self) -> None:
        if urlparse(self.path).path == "/api/config":
            attrs = []
            for attr in SCRAPER.attributes():
                raw = SCRAPER.catalog.get("prod_attrs", {}).get(attr["attribute_id"], {})
                item = dict(attr)
                item["exceptions"] = raw.get("exceptions", {})
                attrs.append(item)
            self._json(200, {
                "product_id": SCRAPER.product_id,
                "product_name": SCRAPER.catalog.get("product_name"),
                "product_code": SCRAPER.catalog.get("product_code"),
                "source_url": SCRAPER.url,
                "product_image": SCRAPER.product_image,
                "default_selection": SCRAPER.defaults,
                "attributes": attrs,
            })
            return
        super().do_GET()

    def do_POST(self) -> None:
        global SCRAPER
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
                    STATE_FILE.write_text(json.dumps({"url": url}, indent=2), encoding="utf-8")
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
            safe_selection = {
                str(key): str(value)
                for key, value in selection.items()
                if str(key).startswith("attr") and str(key)[4:].isdigit()
            }
            protected = str(body.get("changed_attribute_id", ""))
            safe_selection = normalize_selection(safe_selection, protected)
            result = SCRAPER.price(safe_selection)
            payload = {
                "ok": True,
                "price": result.get("price"),
                "unit_price": result.get("unit_price"),
                "quantity": result.get("qty"),
                "turnaround_days": result.get("turnaround"),
                "display_specs": result.get("display_specs", []),
                "price_data": result.get("price_data", {}),
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
    global SCRAPER
    print("Loading UPrinting product configuration...")
    if STATE_FILE.exists():
        try:
            saved_url = json.loads(STATE_FILE.read_text(encoding="utf-8")).get("url", DEFAULT_URL)
            SCRAPER = UPrintingScraper(saved_url)
        except Exception:
            pass
    SCRAPER.load()
    server = ThreadingHTTPServer(("127.0.0.1", 8877), PreviewHandler)
    print("Preview ready: http://127.0.0.1:8877/index.html")
    server.serve_forever()


if __name__ == "__main__":
    main()

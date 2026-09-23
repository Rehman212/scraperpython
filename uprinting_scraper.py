"""Export UPrinting product variation IDs and live prices to JSON/XLSX."""

from __future__ import annotations

import argparse
import base64
import html as html_lib
import itertools
import json
import logging
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urlparse

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

LOG = logging.getLogger("uprinting")
DEFAULT_URL = "https://www.uprinting.com/brochure-printing.html"


class ScraperError(RuntimeError):
    pass


def _first(pattern: str, text: str, description: str, flags: int = 0) -> str:
    match = re.search(pattern, text, flags)
    if not match:
        raise ScraperError(f"Page se {description} nahi mila; site markup shayad change ho gaya hai.")
    return match.group(1)


class UPrintingScraper:
    def __init__(self, url: str, timeout: float = 30, retries: int = 4) -> None:
        parsed = urlparse(url)
        if parsed.scheme not in {"http", "https"} or not parsed.hostname:
            raise ValueError("Valid http(s) product URL dein.")
        if not parsed.hostname.endswith("uprinting.com"):
            raise ValueError("Yeh scraper sirf uprinting.com URLs accept karta hai.")
        self.url = url
        self.timeout = timeout
        retry = Retry(
            total=retries,
            backoff_factor=0.7,
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=frozenset({"GET", "POST"}),
        )
        self.session = requests.Session()
        self.session.mount("https://", HTTPAdapter(max_retries=retry))
        self.session.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/131.0.0.0 Safari/537.36"
                ),
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "en-US,en;q=0.9",
            }
        )
        self.product_id = ""
        self.api_url = ""
        self.auth = ""
        self.defaults: dict[str, str] = {}
        self.visible_attr_ids: list[str] = []
        self.hidden_attr_ids: set[str] = set()
        self.hidden_value_ids: dict[str, set[str]] = {}
        self.catalog: dict[str, Any] = {}
        self.product_image = ""
        self.page_html = ""
        self.linked_calculators: list[dict[str, Any]] = []
        self.price_options: dict[str, Any] = {}
        self.description = ""
        self.images: list[str] = []
        self.video = ""
        self.page_title = ""
        # Storefront dropdown labels (e.g. "18pt Cardstock") can differ from
        # calculator catalog names (e.g. "Classic"). Filled from page HTML.
        self.option_labels: dict[str, str] = {}
        # Page-level product family tiles (Mailer / Product / Shipping Boxes).
        self.product_family_switch: list[dict[str, Any]] = []
        # Label for redirect-dropdown family switches (e.g. "A-Frame Style").
        self.product_family_switch_label: str = ""
        # "cards" (sms-redirect tiles) or "dropdown" (A-Frame Style redirect menu).
        self.product_family_switch_display: str = "cards"
        # Calculator field order as rendered on the PDP (attr_container_*).
        self.attr_display_order: list[str] = []
        # multiCalcConfig.multicalc_switch_display: "button" | "dropdown"
        self.linked_switch_display: str = ""
        # Icon URLs for box-list options (Shape tiles on Lip Balm Labels, etc.).
        self.option_icons: dict[str, str] = {}
        # Attribute IDs rendered as icon tiles (class box-attr-N on the PDP).
        self.box_attr_ids: set[str] = set()

    def load(self) -> None:
        LOG.info("Product page download ho raha hai: %s", self.url)
        response = self.session.get(self.url, timeout=self.timeout)
        response.raise_for_status()
        html = response.text
        self.page_html = html
        page_product_id = _first(
            r"var\s+page_product_id\s*=\s*[\"']?(\d+)", html, "product ID"
        )
        # Prefer the live calculator widget when it points at a different ID.
        # Some PDPs (e.g. Car Decals) set page_product_id to a stub catalog that
        # only exposes Quantity and returns $1.00, while the active gallery /
        # shared-calc params use the real product (Material, Size, $82.94…).
        calc_product_id = None
        active_gallery = re.search(
            r'class="[^"]*active-gallery-calc[^"]*"[^>]*data-product-id=["\'](\d+)["\']'
            r'|data-product-id=["\'](\d+)["\'][^>]*class="[^"]*active-gallery-calc[^"]*"',
            html,
        )
        if active_gallery:
            calc_product_id = active_gallery.group(1) or active_gallery.group(2)
        if not calc_product_id:
            params_match = re.search(
                r"setAttribute\(\s*[\"']params[\"']\s*,\s*[\"'](\{.*?\})[\"']\s*\)",
                html,
                re.S,
            )
            if params_match:
                try:
                    params = json.loads(params_match.group(1))
                    candidate = str(params.get("product_id") or "").strip()
                    if candidate.isdigit() and candidate != "0":
                        calc_product_id = candidate
                except (TypeError, ValueError, json.JSONDecodeError):
                    calc_product_id = None
        self.product_id = calc_product_id or page_product_id
        if self.product_id == "0":
            raise ScraperError("Yeh category/landing page hai, configurable product page nahi. Specific product URL dein.")
        # A landing page's <h1> (e.g. "Square Rounded Corner Business Cards")
        # names this specific product, while the shared calculator's catalog
        # name (e.g. "Die-Cut Business Cards") only names the generic parent
        # it was configured from.
        title_match = re.search(r"<h1[^>]*>\s*([^<]+?)\s*</h1>", html)
        self.page_title = title_match.group(1).strip() if title_match else ""
        image_match = re.search(r"var\s+thumbnail_image\s*=\s*image_domain\s*\+\s*['\"]\/['\"]\s*\+\s*['\"]([^'\"]+)", html)
        image_domain_match = re.search(r"var\s+image_domain\s*=\s*['\"]([^'\"]+)", html)
        if image_match and image_domain_match:
            self.product_image = image_domain_match.group(1).rstrip("/") + "/" + image_match.group(1).lstrip("/")
        else:
            og = re.search(r'<meta[^>]+property=["\']og:image["\'][^>]+content=["\']([^"\']+)', html, re.I)
            self.product_image = og.group(1) if og else ""
        self.api_url = _first(
            r"api_compute_price_url\s*:\s*[\"'](https?:\\?/\\?/[^\"']+/v1)/computePrice",
            html,
            "calculator API URL",
        ).replace("\\/", "/")
        # The page's legacy inline config can still name calculator.uprinting.com,
        # while the live calculator bundle routes requests to the current pricing
        # service.  The legacy host can return an older grid (e.g. $46.76 instead
        # of the storefront's $40.86 for Foam Boards).
        if self.api_url.startswith("https://calculator.uprinting.com/"):
            self.api_url = self.api_url.replace(
                "https://calculator.uprinting.com/",
                "https://calculator.digitalroom.com/",
                1,
            )
        key = _first(r"clients\s*:\s*\{\s*key\s*:\s*[\"']([^\"']+)", html, "API client key")
        secret = _first(r"secret\s*:\s*[\"']([^\"']+)", html, "API client secret")
        self.auth = "Basic " + base64.b64encode(f"{key}:{secret}".encode()).decode()

        pricing_match = re.search(r"var\s+CalcPricingData\s*=\s*(\{.*?\});\s*window\.addEventListener", html, re.S)
        pricing: dict[str, Any] = {}
        if pricing_match:
            pricing = json.loads(pricing_match.group(1))
            initial = pricing.get("request", {}).get("initial_price_data", {})
            normalized = pricing.get("response", {}).get("price", {}).get("price_data", {})
            # These calculator switches are not regular attributes, but some
            # products require them to reproduce the storefront price.
            self.price_options = {
                key: normalized[key]
                for key in (
                    "calc_attrs_option", "calc_attrs", "use_default",
                    "override_invalid_spec", "get_shipping_base_price",
                )
                if key in normalized
            }
            if normalized:
                initial = {**initial, **{k: v for k, v in normalized.items() if str(k).startswith("attr")}}
            # Landing pages for a single shape/size (e.g. Square Rounded Corner
            # Business Cards) share their calculator with a generic parent
            # product (Die-Cut Business Cards), which lists every attribute
            # (Shape, Width, Height, ...) as visible. The pricing engine's own
            # value_exceptions for the page's default selection is what the
            # storefront actually uses to hide those fields, so mirror it here
            # instead of trusting the generic attribute list alone.
            value_exceptions = pricing.get("response", {}).get("value_exceptions", {})
            if not isinstance(value_exceptions, dict):
                value_exceptions = {}
            hidden_attributes = value_exceptions.get("hidden_attributes", [])
            if not isinstance(hidden_attributes, list):
                hidden_attributes = []
            self.hidden_attr_ids = {str(x) for x in hidden_attributes}
            # Most products use {"attrId": [valueIds...]}; Stretched Canvas and
            # a few others send [] when nothing is value-hidden.
            hidden_values = value_exceptions.get("hidden_values", {})
            if not isinstance(hidden_values, dict):
                hidden_values = {}
            self.hidden_value_ids = {
                str(attr_id): {str(v) for v in (values if isinstance(values, (list, set, tuple)) else [])}
                for attr_id, values in hidden_values.items()
            }
        else:
            override_match = re.search(r'"price_data_override"\s*:\s*(\{[^}]*\})', html)
            initial = json.loads(override_match.group(1)) if override_match else {}
        self.defaults = {str(k): str(v) for k, v in initial.items() if str(k).startswith("attr")}

        visible_match = re.search(r'"visible_attrs"\s*:\s*\{\s*"attr_ids"\s*:\s*(\[[^]]*])', html)
        self.visible_attr_ids = [str(x) for x in json.loads(visible_match.group(1))] if visible_match else []
        self.catalog = self._post(f"getData/{self.product_id}", self._base_payload(include_product=False))
        # Page HTML can contain retired hidden option IDs. Keep overrides only
        # when they still exist in the current calculator catalog.
        # Free-text attrs (field_type "t", e.g. Vinyl Lettering Width/Height)
        # have empty prod_attr_vals — keep the page/catalog numeric default.
        self.defaults = self._clean_attr_defaults(self.defaults, self.catalog)
        self.option_labels = self._parse_option_labels(html)
        # Initial computePrice display_specs also carry customer-facing labels
        # for the default selection (covers options missing data-display).
        if pricing_match:
            specs = (
                pricing.get("response", {}).get("price", {}).get("display_specs")
                or pricing.get("response", {}).get("display_specs")
                or []
            )
            if isinstance(specs, list):
                for item in specs:
                    if not isinstance(item, dict):
                        continue
                    option_id = str(item.get("prod_attr_val_id") or "").strip()
                    label = str(item.get("attr_value") or "").strip()
                    if option_id and option_id != "0" and label:
                        self.option_labels.setdefault(option_id, label)
        self.linked_calculators = self._parse_linked_calculators(html)
        self.product_family_switch = self._parse_product_family_switch(html)
        self.attr_display_order = self._parse_attr_display_order(html)
        self.option_icons = self._parse_option_icons(html)
        self.box_attr_ids = self._parse_box_attr_ids(html)
        self.description = self._extract_description(html)
        self.images, self.video = self._extract_gallery(html)

    @staticmethod
    def _parse_option_icons(html: str) -> dict[str, str]:
        """Map option ids to icon URLs from storefront box-list tiles."""
        icons: dict[str, str] = {}
        for match in re.finditer(
            r'id="box_(\d+)"[\s\S]{0,500}?<img[^>]+src="([^"]+)"',
            html,
            re.I,
        ):
            icons[match.group(1)] = match.group(2).strip()
        for match in re.finditer(
            r'data-value="(\d+)"[\s\S]{0,400}?class="attr-icon"[\s\S]{0,120}?src="([^"]+)"',
            html,
            re.I,
        ):
            icons.setdefault(match.group(1), match.group(2).strip())
        return icons

    @staticmethod
    def _parse_box_attr_ids(html: str) -> set[str]:
        """Return attribute ids UPrinting renders as icon box tiles."""
        return {str(x) for x in re.findall(r"\bbox-attr-(\d+)\b", html)}

    @staticmethod
    def _extract_balanced_div(html: str, open_tag_pattern: str) -> str:
        """Return the inner HTML of the first div matching open_tag_pattern, tracking nested <div> depth."""
        match = re.search(open_tag_pattern, html)
        if not match:
            return ""
        pos = match.end()
        depth = 1
        content_end = -1
        while depth > 0:
            next_open = html.find("<div", pos)
            next_close = html.find("</div>", pos)
            if next_close == -1:
                return ""
            if next_open != -1 and next_open < next_close:
                depth += 1
                pos = next_open + len("<div")
            else:
                depth -= 1
                content_end = next_close
                pos = next_close + len("</div>")
        return html[match.end():content_end].strip()

    def _extract_description(self, html: str) -> str:
        """Return cleaned Overview copy, without widget CSS dumps."""
        raw = self._extract_balanced_div(html, r'<div\s+class="overview-product-region">')
        if not raw:
            return ""
        cleaned = re.sub(r"<style[\s\S]*?</style>", "", raw, flags=re.I)
        cleaned = re.sub(r"<script[\s\S]*?</script>", "", cleaned, flags=re.I)
        return cleaned.strip()

    def _extract_gallery(self, html: str) -> tuple[list[str], str]:
        """Return the product's gallery image URLs (highest-res) and a video URL, if present."""
        block = self._extract_balanced_div(html, r"<div\s+product-gallery-widget\s+class=\"product-gallery-widget[^\"]*\">")
        if not block:
            return [], ""
        images: list[str] = []
        seen = set()
        for url in re.findall(r'image-zoom="([^"]+)"', block):
            if url not in seen:
                seen.add(url)
                images.append(url)
        video_match = re.search(r'(https?:[^\s"\']+(?:youtube\.com/embed|player\.vimeo\.com/video|\.mp4)[^\s"\']*)', block, re.I)
        video = video_match.group(1).replace("\\/", "/") if video_match else ""
        return images, video

    @staticmethod
    def _parse_option_labels(html: str) -> dict[str, str]:
        """Map option IDs to the storefront dropdown labels from page HTML."""
        labels: dict[str, str] = {}
        for display, value_a, value_b, display_b in re.findall(
            r'data-display="([^"]*)"[^>]*data-value="(\d+)"'
            r'|data-value="(\d+)"[^>]*data-display="([^"]*)"',
            html,
        ):
            if display:
                option_id, raw = value_a, display
            else:
                option_id, raw = value_b, display_b
            text = html_lib.unescape(raw.replace("&#x20;", " ")).strip()
            if option_id and text:
                labels[str(option_id)] = text
        return labels

    def _storefront_option_label(self, option_id: str, catalog_label: object) -> str:
        """Prefer page dropdown text; strip catalog flute suffixes for linked calcs."""
        labeled = self.option_labels.get(str(option_id))
        if labeled:
            return labeled
        text = html_lib.unescape(str(catalog_label or "")).strip()
        # Linked Black calculator catalog appends "E - Flute" / "B - Flute" to
        # Material names; the Mailer Boxes storefront never shows that suffix.
        return re.sub(r"\s+[EB]\s*-\s*Flute\b.*$", "", text, flags=re.I).strip() or text

    @staticmethod
    def _parse_attr_display_order(html: str) -> list[str]:
        """Return calculator attribute IDs in the order the storefront renders them."""
        ordered: list[str] = []
        seen: set[str] = set()
        for attr_id in re.findall(r'id="attr_container_(\d+)"', html):
            if attr_id not in seen:
                seen.add(attr_id)
                ordered.append(attr_id)
        return ordered

    def _parse_product_family_switch(self, html: str) -> list[dict[str, Any]]:
        """Parse product-family switches above the calculator.

        Mailer/Product/Shipping use sms-redirect image tiles. A-Frame Style (and
        similar) use a redirect-dropdown-menu that navigates between PDPs.
        """
        options = self._parse_sms_redirect_family(html)
        if options:
            self.product_family_switch_label = ""
            self.product_family_switch_display = "cards"
            return options
        options = self._parse_redirect_dropdown_family(html)
        if options:
            self.product_family_switch_display = "dropdown"
            return options
        self.product_family_switch_label = ""
        self.product_family_switch_display = "cards"
        return []

    def _parse_sms_redirect_family(self, html: str) -> list[dict[str, Any]]:
        """Parse Mailer / Product / Shipping style sms-redirect tiles."""
        match = re.search(
            r'<div class="sms-redirect">\s*<div class="content-item">(.*?)</div>\s*</div>\s*<div class="sms-wgt-text-image',
            html,
            re.S | re.I,
        )
        if not match:
            return []
        block = match.group(1)
        options: list[dict[str, Any]] = []
        # Active tile has no <a href>; linked tiles wrap label+icon in an anchor.
        for chunk in re.finditer(
            r'<div class="(active|)[^"]*"\s*[^>]*>\s*(?:<a href="([^"]+)"[^>]*>)?\s*'
            r'<div class="img-wrap">\s*<img src="([^"]+)"[^>]*>\s*</div>\s*'
            r'<div class="item-label">\s*([^<]+)',
            block,
            re.S | re.I,
        ):
            active = chunk.group(1) == "active"
            href = (chunk.group(2) or "").strip()
            image = chunk.group(3).strip()
            label = re.sub(r"\s+", " ", chunk.group(4)).strip()
            if not label:
                continue
            if href and href.startswith("/"):
                href = "https://www.uprinting.com" + href
            options.append({
                "label": label,
                "url": href or self.url,
                "image": image,
                "active": active or (not href),
            })
        return self._finalize_family_options(options)

    def _parse_redirect_dropdown_family(self, html: str) -> list[dict[str, Any]]:
        """Parse A-Frame Style style redirect dropdowns between related PDPs."""
        match = re.search(
            r'<div class="multi-calc-items redirect-dropdown-menu">([\s\S]*?)</ul>\s*</div>',
            html,
            re.I,
        )
        if not match:
            return []
        prefix = html[max(0, match.start() - 500) : match.start()]
        label_match = re.search(
            r"<span>\s*([^<]+?)\s*</span>\s*(?:<span class=\"multical-tooltip-container[^\"]*\"|</label>)",
            prefix,
            re.I,
        )
        self.product_family_switch_label = re.sub(
            r"\s+", " ", html_lib.unescape(label_match.group(1) if label_match else "Style")
        ).strip().rstrip(":")
        block = match.group(1)
        options: list[dict[str, Any]] = []
        for chunk in re.finditer(
            r'<li[^>]*class="([^"]*)"[^>]*>\s*<a href="([^"]+)"[^>]*>\s*<span class="val">([^<]+)',
            block,
            re.I,
        ):
            href = chunk.group(2).strip()
            label = re.sub(r"\s+", " ", html_lib.unescape(chunk.group(3))).strip()
            if not label:
                continue
            if href.startswith("/"):
                href = "https://www.uprinting.com" + href
            active = "active" in chunk.group(1).split()
            options.append({
                "label": label,
                "url": href,
                "image": "",
                "active": active,
            })
        return self._finalize_family_options(options)

    def _finalize_family_options(self, options: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """Ensure exactly one active family option."""
        if options and not any(item["active"] for item in options):
            current = self.url.rstrip("/").lower()
            for item in options:
                item["active"] = item["url"].rstrip("/").lower() == current
        if sum(1 for item in options if item["active"]) != 1 and options:
            for item in options:
                item["active"] = item["url"].rstrip("/").lower() == self.url.rstrip("/").lower()
            if not any(item["active"] for item in options):
                options[0]["active"] = True
                for item in options[1:]:
                    item["active"] = False
        return options

    def _parse_linked_calculators(self, html: str) -> list[dict[str, Any]]:
        """Return page-level calculator switches such as 25/50 Sheets."""
        widget_match = re.search(r"var\s+calculator_widget\s*=\s*JSON\.parse\('((?:\\.|[^'])*)'\)", html, re.S)
        config_match = re.search(r"var\s+multiCalcConfig\s*=\s*(\{.*?\});\s*multiCalcConfig\.display_type", html, re.S)
        if not widget_match or not config_match:
            return []
        try:
            encoded = json.loads('"' + widget_match.group(1).replace('"', '\\"').replace('\\"', '\\"') + '"')
            widgets = json.loads(encoded)
            config = json.loads(config_match.group(1))
        except (ValueError, json.JSONDecodeError):
            # JSON.parse uses a JavaScript string. unicode_escape handles its
            # escaped quotes on pages where JSON's string decoder is stricter.
            try:
                widgets = json.loads(bytes(widget_match.group(1), "utf-8").decode("unicode_escape"))
                config = json.loads(config_match.group(1))
            except Exception:
                return []
        result = []
        self.linked_switch_display = str(
            config.get("multicalc_switch_display")
            or config.get("display_type")
            or ""
        ).strip().lower()
        for option in config.get("calc_switch", []):
            calc_id = str(option.get("calc_id", ""))
            widget = widgets.get(calc_id, {})
            product_id = str(widget.get("product_id", ""))
            if product_id:
                icon = option.get("icon") if isinstance(option.get("icon"), dict) else {}
                icon_url = str(icon.get("path") or "").strip()
                result.append({
                    "calc_id": calc_id,
                    "product_id": product_id,
                    "label": str(option.get("label", "")).strip(),
                    "switch_label": str(config.get("switch_label", "Option")).strip(),
                    "icon": icon_url,
                    "defaults": {f"attr{x['attribute_id']}": str(x["default_prod_attr_val_id"])
                                 for x in widget.get("prod_attrs", [])
                                 if str(x.get("default_prod_attr_val_id", "0")) not in ("0", "")},
                    "visible_attr_ids": [str(x["attribute_id"]) for x in widget.get("prod_attrs", []) if x.get("hide_flag") != "y"],
                })
        # Storefront renders icon-less "button" switches as dropdowns (A-Frame
        # Display Options). Only keep button/cards when at least one icon exists.
        if self.linked_switch_display in ("button", "") and result and not any(
            str(x.get("icon") or "").strip() for x in result
        ):
            self.linked_switch_display = "dropdown"
        return result

    def linked_scraper(self, linked: dict[str, Any]) -> "UPrintingScraper":
        child = UPrintingScraper(self.url, self.timeout)
        child.api_url, child.auth, child.product_id = self.api_url, self.auth, linked["product_id"]
        child.product_image, child.page_html = self.product_image, self.page_html
        child.option_labels = dict(self.option_labels)
        child.option_icons = dict(self.option_icons)
        child.box_attr_ids = set(self.box_attr_ids)
        child.attr_display_order = list(self.attr_display_order)
        # computePrice requires the same pricing flags as a fully load()ed page;
        # without them linked variants (e.g. Address Labels → Sheet) return a
        # stub $1.00 instead of the real matrix price.
        child.price_options = dict(self.price_options)
        child.visible_attr_ids = list(linked["visible_attr_ids"])
        child.catalog = child._post(f"getData/{child.product_id}", child._base_payload(include_product=False))
        # linked["defaults"] comes from the PARENT page's embedded switcher
        # config, not from this child product's own catalog - load() already
        # re-validates every default against prod_attr_vals for the top-level
        # product (a few lines up); a linked variant needs the exact same
        # treatment or a single stale/renumbered default (e.g. "0", or an ID
        # that existed when the switcher config was authored but has since
        # been retired) poisons every selection in a sweep, since sweep mode
        # always starts from these defaults unmodified and only varies one
        # other attribute at a time - one bad key means the whole variant
        # 412s on every request and silently produces zero price rows.
        clean_defaults: dict[str, str] = {}
        values_by_attr: dict[str, dict[str, Any]] = {}
        catalog_default: dict[str, str] = {}
        text_attr_keys: set[str] = set()
        for attr_id, attr in child.catalog.get("prod_attrs", {}).items():
            key = f"attr{attr_id}"
            values = self._attr_values_map(attr.get("prod_attr_vals", {}))
            values_by_attr[key] = values
            catalog_default[key] = str(attr.get("default_value", ""))
            if str(attr.get("field_type") or "") == "t":
                text_attr_keys.add(key)
        # Keep page switcher defaults when valid (landing pages often set qty
        # differently than the standalone calculator catalog).
        for key, candidate in linked["defaults"].items():
            if key in text_attr_keys:
                if candidate not in (None, ""):
                    clean_defaults[key] = str(candidate)
                continue
            values = values_by_attr.get(key, {})
            if candidate in values:
                clean_defaults[key] = candidate
        # Fill gaps from this calculator's own catalog defaults (linked Sheet
        # switcher snapshots are often incomplete and omit material/sheets).
        for key, values in values_by_attr.items():
            if key in clean_defaults:
                continue
            candidate = catalog_default.get(key, "")
            if key in text_attr_keys:
                if candidate not in (None, ""):
                    clean_defaults[key] = str(candidate)
                continue
            # Seed list-backed Width/Height the same way as top-level load().
            attr_id = key[4:]
            attr = child.catalog.get("prod_attrs", {}).get(attr_id, {})
            is_width = attr_id == "width" or attr.get("width_flag") == "y"
            is_height = attr_id == "height" or attr.get("height_flag") == "y"
            if candidate not in values and is_width:
                candidate = str(child.catalog.get("start_width") or "")
            if candidate not in values and is_height:
                candidate = str(child.catalog.get("start_height") or "")
            if candidate not in values and values:
                candidate = next(iter(values))
            if candidate and (candidate in values or (is_width or is_height) and candidate not in (None, "")):
                clean_defaults[key] = str(candidate)
        # Material/finish on linked calculators: prefer this product's catalog
        # default over a stale parent-page switcher snapshot (Address Labels
        # Roll was stuck on White BOPP instead of White Paper).
        for key in ("attr1", "attr25", "attr17"):
            values = values_by_attr.get(key, {})
            candidate = catalog_default.get(key, "")
            if candidate and candidate in values:
                clean_defaults[key] = candidate
        child.defaults = clean_defaults
        return child

    @staticmethod
    def _iter_attr_values(raw_values: Any) -> list[tuple[str, dict[str, Any]]]:
        """Normalize prod_attr_vals whether the API returned a dict or a list.

        Packaging Sleeves (and similar dynamic-size products) ship Width/Height as
        a list of `{prod_attr_val_id, attr_value, ...}` instead of an id→value map.
        """
        if isinstance(raw_values, dict):
            return [
                (str(value_id), value)
                for value_id, value in raw_values.items()
                if isinstance(value, dict)
            ]
        if isinstance(raw_values, list):
            items: list[tuple[str, dict[str, Any]]] = []
            for value in raw_values:
                if not isinstance(value, dict):
                    continue
                value_id = value.get("prod_attr_val_id", value.get("id", ""))
                if value_id in (None, ""):
                    continue
                items.append((str(value_id), value))
            return items
        return []

    @classmethod
    def _attr_values_map(cls, raw_values: Any) -> dict[str, dict[str, Any]]:
        return {value_id: value for value_id, value in cls._iter_attr_values(raw_values)}

    @classmethod
    def _clean_attr_defaults(
        cls,
        raw_defaults: dict[str, str],
        catalog: dict[str, Any],
    ) -> dict[str, str]:
        """Keep option-id defaults that exist in the catalog; keep free-text as-is."""
        clean: dict[str, str] = {}
        for attr_id, attr in catalog.get("prod_attrs", {}).items():
            key = f"attr{attr_id}"
            candidate = raw_defaults.get(key)
            values = cls._attr_values_map(attr.get("prod_attr_vals", {}))
            if str(attr.get("field_type") or "") == "t":
                if candidate in (None, ""):
                    candidate = str(attr.get("default_value", "") or "")
                if candidate not in (None, ""):
                    clean[key] = str(candidate)
                continue
            # Dynamic Width/Height (list-backed selects): page HTML often omits
            # them from initial_price_data; seed from catalog start_* so pricing
            # matches the storefront default (e.g. Packaging Sleeves 4" × 2").
            # Car Magnets keeps width/height hidden until Size=Custom — don't
            # seed those into defaults or they override preset Size factors.
            is_width = str(attr_id) == "width" or attr.get("width_flag") == "y"
            is_height = str(attr_id) == "height" or attr.get("height_flag") == "y"
            if (is_width or is_height) and str(attr_id) in {"width", "height"}:
                if candidate in (None, ""):
                    candidate = str(
                        catalog.get("start_width" if is_width else "start_height") or ""
                    )
                # Keep a seed for Custom Size pricing, but callers that pick a
                # preset Size must drop these via _api_selection(size_dims).
                if candidate not in (None, ""):
                    clean[key] = str(candidate)
                continue
            if candidate not in values:
                candidate = str(attr.get("default_value", "") or "")
            if candidate not in values and is_width:
                candidate = str(catalog.get("start_width") or "")
            if candidate not in values and is_height:
                candidate = str(catalog.get("start_height") or "")
            if candidate not in values and values:
                candidate = next(iter(values))
            if candidate and (candidate in values or (is_width or is_height) and candidate not in (None, "")):
                clean[key] = str(candidate)
        return clean

    def _base_payload(self, include_product: bool = True) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "productType": "offset",
            "publishedVersion": True,
            "disableDataCache": False,
            "disablePriceCache": False,
        }
        if include_product:
            payload["product_id"] = self.product_id
        return payload

    def _post(self, endpoint: str, payload: dict[str, Any]) -> dict[str, Any]:
        headers = {
            "Authorization": self.auth,
            "Origin": "https://www.uprinting.com",
            "Accept": "application/json, text/plain, */*",
            "Referer": "https://www.uprinting.com/",
        }
        response = self.session.post(f"{self.api_url}/{endpoint}", json=payload, headers=headers, timeout=self.timeout)
        if response.status_code >= 400:
            detail = response.text[:500].replace("\n", " ")
            raise ScraperError(f"API {response.status_code}: {detail}")
        return response.json()

    def _trivial_hidden_attr_ids(self) -> set[str]:
        """Hidden attrs with ≤1 selectable value cannot drive real exclusions.

        Foil Labels (and similar) attach exception rules to Printed Side /
        Print Process / Core Size even though those attrs are fixed to a
        single value. Treating those rules as permanent would drop the live
        defaults (Square/Rectangle + 0.75\" x 1.5\"). Landing-page Shape locks
        still work: those hidden attrs keep multiple catalog values.
        """
        trivial: set[str] = set()
        for attr_id in self.hidden_attr_ids:
            attr = self.catalog.get("prod_attrs", {}).get(str(attr_id), {})
            values = [
                value_id
                for value_id, value in self._iter_attr_values(attr.get("prod_attr_vals", {}))
                if value.get("hide_attribute_value_flag") != "y"
                and value.get("custom_flag") != "y"
            ]
            if len(values) <= 1:
                trivial.add(str(attr_id))
        return trivial

    def _sanitize_exceptions(self, exceptions: dict[str, Any]) -> dict[str, Any]:
        """Drop exception clauses that only reference trivial hidden attrs."""
        if not isinstance(exceptions, dict):
            return {}
        trivial = self._trivial_hidden_attr_ids()
        if not trivial:
            return {
                str(option_id): [
                    {str(k): str(v) for k, v in rule.items()}
                    for rule in (rules if isinstance(rules, list) else [])
                    if isinstance(rule, dict)
                ]
                for option_id, rules in exceptions.items()
            }
        cleaned: dict[str, Any] = {}
        for option_id, rules in exceptions.items():
            if not isinstance(rules, list):
                continue
            new_rules = []
            for rule in rules:
                if not isinstance(rule, dict):
                    continue
                filtered = {
                    str(k): str(v)
                    for k, v in rule.items()
                    if str(k) not in trivial
                }
                if filtered:
                    new_rules.append(filtered)
            cleaned[str(option_id)] = new_rules
        return cleaned

    def _rule_permanently_excludes(self, rule: dict[str, Any]) -> bool:
        """True when a rule is always active because only non-trivial hidden attrs lock it."""
        if not isinstance(rule, dict) or not rule:
            return False
        trivial = self._trivial_hidden_attr_ids()
        effective = {
            str(k): str(v)
            for k, v in rule.items()
            if str(k) not in trivial
        }
        if not effective:
            return False
        return all(
            attr_id in self.hidden_attr_ids
            and self.defaults.get(f"attr{attr_id}") == value
            for attr_id, value in effective.items()
        )

    def _is_storefront_custom_size_attr(self, attr_id: str, attr: dict[str, Any]) -> bool:
        """True for Size=Custom Size rows the storefront still renders.

        Folding-carton calculators (Roll-End Tuck, etc.) mark Size with
        hide_attribute_flag=y and omit it from visible_attrs, but the DOM still
        shows the Custom Size control above Length/Width/Depth.
        """
        if str(attr_id) not in {str(x) for x in self.attr_display_order}:
            return False
        name = str(
            attr.get("product_attribute_name") or attr.get("attribute_name") or ""
        ).strip().lower()
        if name != "size":
            return False
        for _value_id, value in self._iter_attr_values(attr.get("prod_attr_vals", {})):
            label = str(value.get("attr_value", "")).strip().lower()
            if value.get("custom_flag") == "y" and "custom" in label:
                return True
            if label == "custom size":
                return True
        return False

    def _is_storefront_readout_attr(self, attr_id: str, attr: dict[str, Any]) -> bool:
        """True for DOM-ordered attrs shown as fixed readouts (Frame Size, A-Frame Type).

        These are often omitted from visible_attrs but still rendered beside the
        interactive fields. Skip the multi-calc twin (e.g. Display Options) when
        linked calculators already own that switch.
        """
        if str(attr_id) not in {str(x) for x in self.attr_display_order}:
            return False
        if attr.get("hide_attribute_flag") == "y":
            return False
        name = str(
            attr.get("product_attribute_name") or attr.get("attribute_name") or ""
        ).strip().lower()
        if self.linked_calculators:
            switch = str(self.linked_calculators[0].get("switch_label") or "").strip().lower()
            if switch and name == switch:
                return False
        if not any(True for _ in self._iter_attr_values(attr.get("prod_attr_vals", {}))):
            return False
        # Prefer the live DOM: hidden containers without a visible
        # single-option-attribute readout are not customer-facing (e.g. Color /
        # Printed Side stay in display_order but only as hidden calc state).
        shown = self._dom_shows_attribute(str(attr_id))
        if shown is False:
            return False
        return True

    def _dom_shows_attribute(self, attr_id: str) -> bool | None:
        """Return whether the PDP currently renders this attribute to customers.

        None means the page HTML is unavailable so callers should fall back to
        catalog visibility rules.
        """
        html = getattr(self, "page_html", "") or ""
        if not html:
            return None
        match = re.search(
            rf'id="attr_container_{re.escape(str(attr_id))}"([^>]*)>',
            html,
            re.I,
        )
        if not match:
            return False
        tag_attrs = match.group(1)
        hidden = bool(re.search(r'\bclass="[^"]*\bhidden\b', tag_attrs, re.I))
        if not hidden:
            return True
        prefix = html[max(0, match.start() - 280) : match.start()]
        return "single-option-attribute" in prefix

    def _hidden_value_depends_on_visible_attr(
        self, value_id: str, exceptions: dict[str, Any]
    ) -> bool:
        """True when a pricing hidden_value can reappear after changing a visible option.

        Box Style on packaging products is a common case: Material=24pt unlocks a
        different identical-looking style id that landing-page hidden_values omit.
        """
        visible = {str(x) for x in self.visible_attr_ids}
        for rule in exceptions.get(str(value_id), []) or []:
            if isinstance(rule, dict) and any(str(key) in visible for key in rule):
                return True
        return False

    def uses_boxes_price_layout(self) -> bool:
        """Match UPrinting packaging calculators: unit price large, subtotal under it."""
        # Image-tile family switches (Mailer/Product/Shipping) use the boxes price
        # layout. Redirect dropdowns (A-Frame Style) do not.
        if self.product_family_switch and self.product_family_switch_display != "dropdown":
            return True
        attrs = self.catalog.get("prod_attrs", {})
        if not isinstance(attrs, dict):
            return False
        has_width = any(a.get("width_flag") == "y" for a in attrs.values() if isinstance(a, dict))
        has_height = any(a.get("height_flag") == "y" for a in attrs.values() if isinstance(a, dict))
        has_depth = any(a.get("depth_flag") == "y" for a in attrs.values() if isinstance(a, dict))
        return has_width and has_height and has_depth

    def attributes(self, visible_only: bool = True) -> list[dict[str, Any]]:
        result = []
        visible = set(self.visible_attr_ids)
        for attr_id, attr in self.catalog.get("prod_attrs", {}).items():
            # Attributes listed as currently hidden can still belong to the
            # calculator (White Ink until Clear BOPP, Width/Height until Custom).
            # Only drop attributes that are not in the page's visible_attr_ids —
            # except free-size Width/Height on dynamic custom-size products,
            # which the storefront reveals when Size=Custom.
            is_free_size = (
                str(attr_id) in {"width", "height"}
                and self.catalog.get("dynamic_size") == "c"
            )
            force_custom_size = self._is_storefront_custom_size_attr(str(attr_id), attr)
            force_readout = self._is_storefront_readout_attr(str(attr_id), attr)
            if (
                visible_only
                and str(attr_id) in self.hidden_attr_ids
                and str(attr_id) not in visible
                and not is_free_size
                and not force_custom_size
                and not force_readout
            ):
                continue
            if (
                visible_only
                and visible
                and str(attr_id) not in visible
                and not is_free_size
                and not force_custom_size
                and not force_readout
            ):
                continue
            if attr.get("hide_attribute_flag") == "y" and not force_custom_size:
                continue
            values = []
            value_items = self._iter_attr_values(attr.get("prod_attr_vals", {}))
            hidden_values = self.hidden_value_ids.get(str(attr_id), set()) if visible_only else set()
            default_value_id = self.defaults.get(f"attr{attr_id}")
            exceptions = attr.get("exceptions", {}) if isinstance(attr.get("exceptions"), dict) else {}
            if is_free_size and "-1" not in exceptions:
                # Hide Width/Height unless Size is the Custom option.
                size_vals = self._attr_values_map(
                    self.catalog.get("prod_attrs", {}).get("3", {}).get("prod_attr_vals", {})
                )
                preset_rules = [
                    {"3": option_id}
                    for option_id, option in size_vals.items()
                    if "custom" not in str(option.get("attr_value", "")).lower()
                    and str(option_id).lower() != "custom"
                ]
                if preset_rules:
                    exceptions = {**exceptions, "-1": preset_rules}
            for value_id, value in value_items:
                if value.get("hide_attribute_value_flag") == "y":
                    continue
                # custom_flag marks free-entry options such as "Custom Size".
                # The storefront still shows those; only skip unlabeled custom
                # sentinels that are not customer-facing.
                if value.get("custom_flag") == "y":
                    custom_label = str(value.get("attr_value", "")).strip().lower()
                    if "custom" not in custom_label:
                        continue
                # The pricing engine's hidden_values also lists the attribute's
                # own current default (e.g. Printing Time's "6 Business Days"
                # alongside the faster options it hides) - it only means those
                # OTHER values aren't switchable-to, not that the default itself
                # should disappear from the dropdown.
                # Dependency-driven hides (Box Style ids gated by Material) must
                # stay in the option list so the UI can unlock them on change.
                if str(value_id) != default_value_id and (
                    str(value_id) in hidden_values or str(value.get("attr_val_id", "")) in hidden_values
                ):
                    if not self._hidden_value_depends_on_visible_attr(str(value_id), exceptions):
                        continue
                # A value's "exceptions" rules mark it invalid whenever some
                # other attribute holds a given value. When every *meaningful*
                # attribute named in a rule is itself hidden (fixed for this
                # landing page, e.g. Shape stuck at "Square Rounded Corner"),
                # that rule can never NOT hold, so the value is permanently
                # invalid here. Ignore single-option hidden attrs (Printed
                # Side=Front Only on Foil Labels) — those rules are catalog
                # noise and the storefront still sells the option.
                if visible_only and any(
                    self._rule_permanently_excludes(rule)
                    for rule in exceptions.get(str(value_id), [])
                ):
                    continue
                # Dynamic-size products expose an internal max-dimension sentinel
                # alongside the customer-facing "Custom" option. The website hides it.
                raw_factors = value.get("factors", {})
                factors = raw_factors if isinstance(raw_factors, dict) else {}
                if (
                    str(attr_id) == "3"
                    and self.catalog.get("dynamic_size") == "c"
                    and str(factors.get("width")) == str(self.catalog.get("end_width"))
                    and str(factors.get("height")) == str(self.catalog.get("end_height"))
                ):
                    continue
                values.append(
                    {
                        "option_id": str(value_id),
                        "source_attr_value_id": str(value.get("attr_val_id", "")),
                        "label": self._storefront_option_label(
                            str(value_id),
                            value.get("attr_value", ""),
                        ),
                        "default": self.defaults.get(f"attr{attr_id}") == str(value_id),
                        "sort_order": value.get("sort_order"),
                        "factors": factors,
                        "icon": self.option_icons.get(str(value_id), ""),
                    }
                )
            values.sort(key=lambda v: (v["sort_order"] is None, v["sort_order"] or 0, v["label"]))
            field_type = str(attr.get("field_type") or "")
            # Storefront box-list attrs (Shape on Lip Balm Labels) must export as
            # buttons so Printoe keeps the icon tiles — including when Style
            # filters the list down to a single Circle option.
            if str(attr_id) in self.box_attr_ids and values:
                field_type = "buttons"
            min_value = None
            max_value = None
            if field_type == "t":
                if attr.get("width_flag") == "y":
                    min_value = self.catalog.get("start_width")
                    max_value = self.catalog.get("end_width")
                elif attr.get("height_flag") == "y":
                    min_value = self.catalog.get("start_height")
                    max_value = self.catalog.get("end_height")
                elif attr.get("depth_flag") == "y":
                    min_value = self.catalog.get("start_depth")
                    max_value = self.catalog.get("end_depth")
                else:
                    min_value = attr.get("default_min_value")
                    max_raw = attr.get("default_max_value")
                    max_value = None if max_raw in (None, "", 0, "0") else max_raw
            default_id = self.defaults.get(f"attr{attr_id}") or str(attr.get("default_value", "") or "")
            result.append(
                {
                    "attribute_id": str(attr_id),
                    "name": attr.get("product_attribute_name") or attr.get("attribute_name") or f"Attribute {attr_id}",
                    "code": attr.get("attribute_code", ""),
                    "field_type": field_type,
                    "data_type": attr.get("data_type", ""),
                    "min_value": None if min_value in (None, "") else str(min_value),
                    "max_value": None if max_value in (None, "") else str(max_value),
                    "default_option_id": str(default_id),
                    "sort_order": attr.get("sort_order"),
                    "options": values,
                    "exceptions": self._sanitize_exceptions(exceptions),
                }
            )
        # Prefer the storefront DOM order (attr_container_*), then visible_attrs,
        # then catalog sort_order. visible_attrs alone is often wrong (e.g. Metal
        # Business Cards lists Material before Size while the page shows Size first).
        display_rank = {str(aid): idx for idx, aid in enumerate(self.attr_display_order)}
        visible_rank = {str(aid): idx for idx, aid in enumerate(self.visible_attr_ids)}
        result.sort(
            key=lambda a: (
                display_rank.get(a["attribute_id"], len(display_rank) + 1),
                visible_rank.get(a["attribute_id"], len(visible_rank) + 1),
                a["sort_order"] is None,
                a["sort_order"] or 0,
            )
        )
        # Product boxes (and similar) list Length/Width/Depth before Size in
        # visible_attrs, but the storefront always shows Size first so the
        # custom-dimension inputs appear under it.
        dim_ids = {
            str(aid)
            for aid, attr in self.catalog.get("prod_attrs", {}).items()
            if attr.get("width_flag") == "y"
            or attr.get("height_flag") == "y"
            or attr.get("depth_flag") == "y"
        }
        size_indexes = [
            index
            for index, attr in enumerate(result)
            if str(attr.get("name", "")).strip().lower() == "size"
        ]
        if size_indexes and dim_ids and not display_rank:
            size_attr = result.pop(size_indexes[0])
            insert_at = next(
                (index for index, attr in enumerate(result) if attr["attribute_id"] in dim_ids),
                len(result),
            )
            result.insert(insert_at, size_attr)
        return result

    def price(self, selection: dict[str, str]) -> dict[str, Any]:
        payload = self._base_payload()
        payload.update(self.price_options)
        payload.update(self._api_selection(selection))
        return self._post("computePrice", payload)

    def _size_attribute_id(self) -> str | None:
        """Return the Size attribute id used for preset width/height/depth factors."""
        prod_attrs = self.catalog.get("prod_attrs", {})
        if "3" in prod_attrs:
            return "3"
        for attr_id, attr in prod_attrs.items():
            name = str(attr.get("product_attribute_name") or attr.get("attribute_name") or "").strip().lower()
            if name == "size":
                return str(attr_id)
        return None

    def _size_dimension_factors(self, selection: dict[str, str]) -> dict[str, str]:
        """Return width/height/depth factors from the selected Size option, if any."""
        size_attr_id = self._size_attribute_id()
        if not size_attr_id:
            return {}
        size_option_id = selection.get(f"attr{size_attr_id}")
        if not size_option_id:
            return {}
        size_attr = self.catalog.get("prod_attrs", {}).get(size_attr_id, {})
        values = self._attr_values_map(size_attr.get("prod_attr_vals", {}))
        option = values.get(str(size_option_id), {})
        factors = option.get("factors", {})
        if not isinstance(factors, dict):
            return {}
        dims: dict[str, str] = {}
        for key in ("width", "height", "depth"):
            value = factors.get(key)
            if value not in (None, ""):
                dims[key] = str(value)
        # Preset sizes always carry width+height; custom size has empty factors.
        if "width" not in dims or "height" not in dims:
            return {}
        return dims

    def _api_selection(self, selection: dict[str, str]) -> dict[str, str]:
        """Translate selection keys/values into the shape computePrice expects."""
        translated = dict(selection)
        size_dims = self._size_dimension_factors(selection)
        for key, option_id in selection.items():
            if not key.startswith("attr"):
                continue
            attr_id = key[4:]
            attribute = self.catalog.get("prod_attrs", {}).get(attr_id, {})
            # Packaging Sleeves uses attribute ids "width"/"height" and prices off
            # bare width/height keys (attrwidth must be rewritten). DTF Transfers
            # and similar products use numeric attrs 247/248 with width_flag /
            # height_flag — those option keys must stay on the payload or the
            # API silently falls back to the catalog default height (e.g. 3).
            # Car Magnets: when a preset Size is selected, its factors already
            # carry width/height. Sending hidden free-size defaults (3x3) on
            # top overrides the Size and underprices (e.g. $13.22 instead of $43.96).
            # Product Boxes also send depth the same way for Custom Size.
            if str(attr_id) == "width" or attribute.get("width_flag") == "y":
                if size_dims:
                    translated.pop(key, None)
                    continue
                translated["width"] = str(option_id)
                if str(attr_id) == "width":
                    translated.pop(key, None)
                continue
            if str(attr_id) == "height" or attribute.get("height_flag") == "y":
                if size_dims:
                    translated.pop(key, None)
                    continue
                translated["height"] = str(option_id)
                if str(attr_id) == "height":
                    translated.pop(key, None)
                continue
            if str(attr_id) == "depth" or attribute.get("depth_flag") == "y":
                if size_dims:
                    translated.pop(key, None)
                    continue
                translated["depth"] = str(option_id)
                if str(attr_id) == "depth":
                    translated.pop(key, None)
                continue
            values = self._attr_values_map(attribute.get("prod_attr_vals", {}))
            if not values:
                continue
            option = values.get(str(option_id), {})
            factors = option.get("factors", {}) if isinstance(option.get("factors"), dict) else {}
            qty_factor = factors.get("display_qty")
            if qty_factor in (None, ""):
                qty_factor = factors.get("qty")
            attr_name = str(
                attribute.get("product_attribute_name")
                or attribute.get("attribute_name")
                or ""
            ).strip().lower()
            is_qty_attr = attribute.get("qty_flag") == "y" or attr_name == "quantity"
            # Candle Labels (and similar) keep quantity option IDs on attr5 but
            # still need qty_var=250 or computePrice reports qty=100 and a wrong
            # unit price ($0.38 instead of $0.15).
            if is_qty_attr and qty_factor not in (None, ""):
                translated["qty_var"] = str(qty_factor).replace(",", "")
                for other_id, other_attr in self.catalog.get("prod_attrs", {}).items():
                    if str(other_id) == str(attr_id):
                        continue
                    other_name = str(
                        other_attr.get("product_attribute_name")
                        or other_attr.get("attribute_name")
                        or ""
                    ).strip().lower()
                    if other_name != "quantity":
                        continue
                    other_key = f"attr{other_id}"
                    if other_key in translated or other_key in self.defaults:
                        translated[other_key] = str(qty_factor).replace(",", "")
            # Most calculators expect a prod_attr_val option ID. Range-based
            # quantity calculators are different: their default_value is an
            # actual quantity (not a key in prod_attr_vals), and sending the
            # option ID makes the API interpret e.g. 1487661 as 1,487,661.
            if str(attribute.get("default_value", "")) in values:
                continue
            display_qty = factors.get("display_qty")
            if display_qty not in (None, ""):
                translated[key] = str(display_qty).replace(",", "")
            elif factors.get("qty") not in (None, ""):
                translated[key] = str(factors.get("qty")).replace(",", "")
        return translated

    def _priced_row(self, selection: dict[str, str], changed: str = "") -> dict[str, Any]:
        response = self.price(selection)
        effective_price = response.get("discounted_price")
        if effective_price is None:
            effective_price = response.get("price")
        quantity = response.get("qty")
        try:
            effective_unit_price = float(effective_price) / float(quantity) if float(quantity) else response.get("unit_price")
        except (TypeError, ValueError, ZeroDivisionError):
            effective_unit_price = response.get("discounted_unit_price", response.get("unit_price"))
        display = {
            str(item["attribute_id"]): {
                "name": item.get("attribute_name", ""),
                "label": item.get("attr_value", ""),
                "option_id": str(item.get("prod_attr_val_id", "")),
            }
            for item in response.get("display_specs", [])
        }
        # UPrinting's calculator can silently substitute an incompatible
        # request with the nearest valid combination (e.g. a Spot UV finish
        # that only exists on a heavier paper at a higher quantity) and price
        # THAT instead. display_specs reports what was actually priced; the
        # requested `selection` is not necessarily what the price/turnaround
        # below describe, so resolve it against display_specs wherever the
        # API confirmed a (possibly different) value, rather than exporting
        # the row under a selection key nothing was truly priced at.
        resolved_selection = dict(selection)
        for attr_id, info in display.items():
            key = f"attr{attr_id}"
            if key not in resolved_selection:
                continue
            option_id = str(info.get("option_id") or "")
            # Range/qty fields often report prod_attr_val_id "0" even though the
            # storefront dropdown options use real IDs (100 → 1488116). Writing
            # "0" into the export makes prune_prices_to_exported_options drop
            # every row ("Export has no valid pricing rows").
            if option_id and option_id not in ("0", ""):
                resolved_selection[key] = option_id
                continue
            attr = self.catalog.get("prod_attrs", {}).get(str(attr_id), {})
            if attr.get("qty_flag") == "y":
                mapped = self._quantity_option_id(
                    str(attr_id),
                    info.get("label") or resolved_selection.get(key) or quantity,
                )
                if mapped:
                    resolved_selection[key] = mapped
        return {
            "changed_attribute_id": changed,
            "selection": resolved_selection,
            "display": display,
            "price": effective_price,
            "original_price": response.get("orig_price"),
            "total_price": response.get("total_price"),
            "unit_price": effective_unit_price,
            "quantity": quantity,
            "turnaround_days": response.get("turnaround"),
            "in_stock": response.get("in_stock_flag"),
            "currency": "USD",
            "order_specs": response.get("order_specs", []),
        }

    def _quantity_option_id(self, attr_id: str, raw: Any) -> str | None:
        """Map a display quantity (100 / 1,000) onto the catalog option id."""
        needle = str(raw or "").replace(",", "").strip()
        if not needle:
            return None
        attr = self.catalog.get("prod_attrs", {}).get(str(attr_id), {})
        for value_id, value in self._iter_attr_values(attr.get("prod_attr_vals", {})):
            factors = value.get("factors", {}) if isinstance(value.get("factors"), dict) else {}
            candidates = {
                str(factors.get("display_qty") or "").replace(",", "").strip(),
                str(value.get("attr_value") or "").replace(",", "").strip(),
            }
            if needle in candidates:
                return str(value_id)
        return None

    def selection_is_valid(self, selection: dict[str, str]) -> bool:
        """Return whether every selected visible value satisfies catalog rules.

        Hidden calculator attrs (Die-Cutting, secondary Quantity, etc.) often
        carry exception rows that do not match the landing-page defaults the
        storefront still prices. Checking those blocked whole sweeps — Name /
        Candle Labels produced 1–2 rows and then pruned to an empty export.
        """
        visible_ids = {str(attr["attribute_id"]) for attr in self.attributes()}
        for attr_id, attr in self.catalog.get("prod_attrs", {}).items():
            if str(attr_id) not in visible_ids:
                continue
            key = f"attr{attr_id}"
            selected = selection.get(key)
            if selected is None:
                continue
            rules = attr.get("exceptions", {}).get(str(selected), [])
            if any(
                all(selection.get(f"attr{rule_attr}") == str(value) for rule_attr, value in rule.items())
                for rule in rules
            ):
                return False
        return True

    def selections(self, mode: str, vary: list[str] | None, max_combinations: int) -> Iterable[tuple[dict[str, str], str]]:
        attrs = self.attributes()
        by_id = {a["attribute_id"]: a for a in attrs}
        if mode == "default":
            yield dict(self.defaults), ""
            return
        if mode == "sweep":
            yield dict(self.defaults), ""
            seen = {tuple(sorted(self.defaults.items()))}
            for attr in attrs:
                for option in attr["options"]:
                    selection = dict(self.defaults)
                    selection[f"attr{attr['attribute_id']}"] = option["option_id"]
                    if not self.selection_is_valid(selection):
                        continue
                    key = tuple(sorted(selection.items()))
                    if key not in seen:
                        seen.add(key)
                        yield selection, attr["attribute_id"]
            return

        chosen = list(by_id) if not vary or vary == ["all"] else vary
        unknown = [x for x in chosen if x not in by_id]
        if unknown:
            raise ValueError(f"Unknown --vary attribute IDs: {', '.join(unknown)}")
        option_lists = [[o["option_id"] for o in by_id[attr_id]["options"]] for attr_id in chosen]
        count = 1
        for options in option_lists:
            count *= len(options)
        if max_combinations and count > max_combinations:
            raise ValueError(
                f"{count:,} combinations banti hain, limit {max_combinations:,} hai. "
                "--vary mein kam attribute IDs dein ya --max-combinations barhayein (0 = unlimited)."
            )
        def rule_matches(rule: dict[str, Any], selection: dict[str, str]) -> bool:
            return all(selection.get(f"attr{k}") == str(v) for k, v in rule.items())

        def walk(index: int, selection: dict[str, str]) -> Iterable[tuple[dict[str, str], str]]:
            if index == len(chosen):
                yield dict(selection), ",".join(chosen)
                return
            attr_id = chosen[index]
            raw_attr = self.catalog.get("prod_attrs", {}).get(attr_id, {})
            exceptions = raw_attr.get("exceptions", {})
            if any(rule_matches(rule, selection) for rule in exceptions.get("-1", [])):
                yield from walk(index + 1, selection)
                return
            for option_id in option_lists[index]:
                if any(rule_matches(rule, selection) for rule in exceptions.get(option_id, [])):
                    continue
                selection[f"attr{attr_id}"] = option_id
                yield from walk(index + 1, selection)
            selection.pop(f"attr{attr_id}", None)

        yield from walk(0, dict(self.defaults))

    def scrape_prices(
        self,
        mode: str,
        vary: list[str] | None,
        max_combinations: int,
        workers: int,
        delay: float,
    ) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
        selections = list(self.selections(mode, vary, max_combinations))
        LOG.info("%d price configurations process hongi", len(selections))
        prices: list[dict[str, Any]] = []
        errors: list[dict[str, Any]] = []

        def run(item: tuple[dict[str, str], str]) -> dict[str, Any]:
            if delay:
                time.sleep(delay)
            return self._priced_row(*item)

        with ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
            futures = {pool.submit(run, item): item for item in selections}
            for index, future in enumerate(as_completed(futures), 1):
                selection, changed = futures[future]
                try:
                    prices.append(future.result())
                except Exception as exc:  # invalid combos are expected in exhaustive mode
                    errors.append({"changed_attribute_id": changed, "selection": selection, "error": str(exc)})
                if index % 50 == 0 or index == len(selections):
                    LOG.info("Progress: %d/%d (valid=%d, invalid=%d)", index, len(selections), len(prices), len(errors))
        prices.sort(key=lambda row: tuple(sorted(row["selection"].items())))
        return prices, errors


def build_export(scraper: UPrintingScraper, mode: str, prices: list[dict[str, Any]], errors: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "metadata": {
            "source_url": scraper.url,
            "scraped_at_utc": datetime.now(timezone.utc).isoformat(),
            "product_id": scraper.product_id,
            "product_code": scraper.catalog.get("product_code"),
            "product_name": scraper.page_title or scraper.catalog.get("product_name"),
            "mode": mode,
            "currency": "USD",
            "valid_price_rows": len(prices),
            "invalid_rows": len(errors),
        },
        # Export callers may add synthetic fields (for example a linked
        # calculator selector). Never let that mutate the live scraper state.
        "default_selection": dict(scraper.defaults),
        "description": scraper.description,
        "product_image": scraper.product_image,
        "images": list(scraper.images),
        "video": scraper.video,
        "attributes": scraper.attributes(),
        "prices": prices,
        "errors": errors,
    }


def save_json(data: dict[str, Any], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def _style_sheet(ws: Any) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = fill
    for column in ws.columns:
        letter = get_column_letter(column[0].column)
        width = min(60, max(11, *(len(str(c.value or "")) + 2 for c in column[:300])))
        ws.column_dimensions[letter].width = width


def save_xlsx(data: dict[str, Any], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    summary.append(["Field", "Value"])
    for key, value in data["metadata"].items():
        if isinstance(value, (dict, list)):
            summary.append([key, json.dumps(value, ensure_ascii=False)])
        else:
            summary.append([key, value])
    summary.append(["default_selection", json.dumps(data["default_selection"], ensure_ascii=False)])

    attributes = wb.create_sheet("Attributes")
    attributes.append(["attribute_id", "attribute_name", "attribute_code", "option_id", "source_attr_value_id", "option_label", "is_default", "factors"])
    for attr in data["attributes"]:
        for option in attr["options"]:
            attributes.append([
                attr["attribute_id"], attr["name"], attr["code"], option["option_id"], option["source_attr_value_id"],
                option["label"], option["default"], json.dumps(option["factors"], ensure_ascii=False),
            ])

    attr_ids = [a["attribute_id"] for a in data["attributes"]]
    attr_names = {a["attribute_id"]: a["name"] for a in data["attributes"]}
    prices_ws = wb.create_sheet("Prices")
    fixed_headers = ["price", "original_price", "total_price", "unit_price", "currency", "quantity", "turnaround_days", "in_stock"]
    variation_headers = list(itertools.chain.from_iterable((f"{attr_names[x]} ID", attr_names[x]) for x in attr_ids))
    prices_ws.append(fixed_headers + variation_headers)
    for row in data["prices"]:
        cells = [row.get(x) for x in fixed_headers]
        for attr_id in attr_ids:
            shown = row.get("display", {}).get(attr_id, {})
            cells.extend([shown.get("option_id", row["selection"].get(f"attr{attr_id}", "")), shown.get("label", "")])
        prices_ws.append(cells)

    errors_ws = wb.create_sheet("Errors")
    errors_ws.append(["changed_attribute_id", "selection", "error"])
    for row in data["errors"]:
        errors_ws.append([row["changed_attribute_id"], json.dumps(row["selection"]), row["error"]])

    for ws in wb.worksheets:
        _style_sheet(ws)
    wb.save(path)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="UPrinting variation IDs aur live prices JSON/XLSX mein export karein.")
    parser.add_argument("url", nargs="?", default=DEFAULT_URL, help="UPrinting product URL")
    parser.add_argument("--mode", choices=("default", "sweep", "exhaustive"), default="sweep")
    parser.add_argument("--vary", default="", help="Exhaustive mode: comma-separated attribute IDs, ya 'all'")
    parser.add_argument("--max-combinations", type=int, default=10000, help="Safety cap; 0 = unlimited")
    parser.add_argument("--workers", type=int, default=4, help="Concurrent API calls (recommended <= 4)")
    parser.add_argument("--delay", type=float, default=0.05, help="Har request se pehle delay seconds")
    parser.add_argument("--output", default="uprinting_export", help="Output path without extension")
    parser.add_argument("--timeout", type=float, default=30)
    parser.add_argument("--verbose", action="store_true")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format="%(levelname)s: %(message)s")
    try:
        scraper = UPrintingScraper(args.url, timeout=args.timeout)
        scraper.load()
        vary = [x.strip() for x in args.vary.split(",") if x.strip()] or None
        prices, errors = scraper.scrape_prices(args.mode, vary, args.max_combinations, args.workers, args.delay)
        data = build_export(scraper, args.mode, prices, errors)
        base = Path(args.output)
        json_path, xlsx_path = base.with_suffix(".json"), base.with_suffix(".xlsx")
        save_json(data, json_path)
        save_xlsx(data, xlsx_path)
        print(f"Done: {len(prices)} valid prices, {len(errors)} invalid combinations")
        print(f"JSON : {json_path.resolve()}")
        print(f"Excel: {xlsx_path.resolve()}")
        return 0 if prices else 2
    except (requests.RequestException, ScraperError, ValueError, json.JSONDecodeError) as exc:
        LOG.error("%s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

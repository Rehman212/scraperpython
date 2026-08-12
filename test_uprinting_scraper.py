from pathlib import Path

from openpyxl import load_workbook

from uprinting_scraper import UPrintingScraper, build_export, save_json, save_xlsx
from preview_server import SCRAPER as PREVIEW_SCRAPER, normalize_selection


URL = "https://www.uprinting.com/brochure-printing.html"


def test_live_page_catalog_and_default_price(tmp_path: Path) -> None:
    scraper = UPrintingScraper(URL)
    scraper.load()
    assert scraper.product_id == "4"
    attributes = scraper.attributes()
    assert any(a["name"] == "Size (before folding)" for a in attributes)
    assert all(a["options"] for a in attributes)
    response = scraper.price(scraper.defaults)
    assert float(response["price"]) > 0
    assert response["display_specs"]

    row = scraper._priced_row(scraper.defaults)
    data = build_export(scraper, "default", [row], [])
    json_path = tmp_path / "result.json"
    xlsx_path = tmp_path / "result.xlsx"
    save_json(data, json_path)
    save_xlsx(data, xlsx_path)
    assert json_path.stat().st_size > 1000
    workbook = load_workbook(xlsx_path, read_only=True)
    assert {"Summary", "Attributes", "Prices", "Errors"}.issubset(workbook.sheetnames)
    assert workbook["Prices"].max_row == 2


def test_live_two_quantity_prices() -> None:
    scraper = UPrintingScraper(URL)
    scraper.load()
    quantity = next(a for a in scraper.attributes() if a["name"] == "Quantity")
    values = quantity["options"][:2]
    prices = []
    for option in values:
        selection = dict(scraper.defaults)
        selection[f"attr{quantity['attribute_id']}"] = option["option_id"]
        prices.append(float(scraper.price(selection)["price"]))
    assert len(prices) == 2
    assert all(price > 0 for price in prices)


def test_preview_matches_uprinting_reference_configuration() -> None:
    PREVIEW_SCRAPER.load()
    size = next(a for a in PREVIEW_SCRAPER.attributes() if a["attribute_id"] == "3")
    assert len(size["options"]) == 10
    selection = {
        "attr1": "69028", "attr3": "143", "attr4": "165", "attr5": "169",
        "attr6": "199", "attr7": "151", "attr16": "194516",
        "attr400": "68456", "attr635": "68458",
    }
    normalized = normalize_selection(selection, protected_attr="7")
    assert normalized["attr4"] == "162"  # UPrinting changes this to Front Only.
    result = PREVIEW_SCRAPER.price(normalized)
    assert result["price"] == "268.20"
    assert result["unit_price"] == "0.5364"


def test_preview_matches_large_brochure_single_and_double_side_prices() -> None:
    PREVIEW_SCRAPER.load()
    base = {
        "attr1": "69028", "attr3": "14731", "attr5": "169", "attr6": "199",
        "attr7": "69030", "attr16": "194516", "attr400": "68456", "attr635": "68458",
    }
    expected = {"164": ("Outside Only", "1089.50"), "165": ("Outside and Inside", "1939.80")}
    for side_id, (side_label, price) in expected.items():
        result = PREVIEW_SCRAPER.price({**base, "attr4": side_id})
        selected_side = next(x for x in result["display_specs"] if x["attribute_id"] == "4")
        assert selected_side["attr_value"] == side_label
        assert result["price"] == price

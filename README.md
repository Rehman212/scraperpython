# UPrinting Variation & Price Scraper

Yeh Python CLI kisi UPrinting product URL se:

- product/attribute IDs aur tamam option IDs/labels nikalti hai;
- UPrinting calculator API se live, exact USD price aur unit price leti hai;
- valid results JSON aur Excel (`.xlsx`) mein export karti hai;
- rejected/invalid combinations ko `Errors` sheet mein rakhti hai.

## Setup

```powershell
python -m pip install -r requirements.txt
```

## Recommended run (har option ka price, baqi settings default)

```powershell
python uprinting_scraper.py "https://www.uprinting.com/brochure-printing.html" --mode sweep --output brochure
```

Output: `brochure.json` aur `brochure.xlsx`.

## Full combinations

Pehle `Attributes` sheet/JSON se attribute IDs dekhein. Misal: Size `3`, Quantity `5`, Printing Time `6`:

```powershell
python uprinting_scraper.py "URL" --mode exhaustive --vary 3,5,6 --output brochure_matrix
```

Har visible attribute ki Cartesian combinations ke liye:

```powershell
python uprinting_scraper.py "URL" --mode exhaustive --vary all --max-combinations 0 --output all_prices
```

`--vary all` bohat bari request count bana sakta hai. Pehle limited attributes use karein; default safety limit 10,000 combinations hai. Site par load kam rakhne ke liye workers 4 se zyada na karein.

## Useful options

```text
--mode default|sweep|exhaustive
--vary 3,5,6
--max-combinations 10000
--workers 4
--delay 0.05
--output result_name
--verbose
```

Prices live hain aur UPrinting kabhi bhi change kar sakta hai. Export metadata mein exact UTC scrape time included hota hai. Site/API markup badalne par scraper update ki zaroorat ho sakti hai.

## Browser preview

```powershell
python preview_server.py
```

Open `http://127.0.0.1:8877/index.html`. Exact configuration IDs query string mein bhi di ja sakti hain, for example `?attr1=69028&attr3=14731&attr4=164`.

## Printoe exact-pricing import

Complete crawler run karein:

```powershell
python full_scrape.py "UPRINTING_PRODUCT_URL" --database product_prices.sqlite
```

Crawler resumable SQLite checkpoint rakhta hai. Complete hone par:

- `product_prices.printoe.json` — Printoe admin ke Customer Fields step mein upload karein.
- `product_prices.xlsx` — review/reporting ke liye.

Printoe mein product Edit → Customer Fields → **Import pricing JSON** → file select → **Update**. Import fields ko populate karta hai aur exact combination prices 500-row chunks mein backend database mein store karta hai. Simple Extra-$ formula imported products ke liye use nahi hota.

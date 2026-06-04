# WB Parser

A command-line tool that collects product data from Wildberries by article number and exports it to Excel and Google Sheets.

Given a list of article numbers in an Excel file, the parser fetches each product's name, brand, price, discount, rating, review count and stock, then writes the results to a local `.xlsx` file and, optionally, to a Google Sheet.

## Features

- Pulls data from two independent Wildberries sources with automatic fallback
- Recovers from transient network errors with batched retries
- Polite request pacing (randomised delays and pauses between batches) to avoid hammering the server
- Deduplicates input and preserves the original article order in the output
- Exports to both a local Excel file and a Google Sheet

## How it works

Wildberries does not expose a single clean API, so the parser talks to two different backends and switches between them:

1. **Card API** (`card.wb.ru`) — the dynamic source for price, discount, rating and stock. Queried first across several delivery regions, since price and availability vary by region.
2. **Basket CDN** (`basket-NN.wbbasket.ru`) — static product cards (name, brand) plus price history. Used as a fallback when the Card API returns nothing.

The basket server number is derived from the article via a lookup table reverse-engineered from how Wildberries distributes products across its CDN. Because that mapping drifts over time, the parser also probes neighbouring servers.

## Tech stack

Python 3.10+, `requests` for HTTP, `openpyxl` for Excel, `gspread` + `google-auth` for Google Sheets.

## Installation

```bash
git clone https://github.com/<your-username>/wb-parser.git
cd wb-parser
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # macOS / Linux
pip install -r requirements.txt
```

## Configuration

Copy `.env.example` to `.env` and fill in your values:

```bash
cp .env.example .env
```

Google Sheets export is **optional**. To enable it:

1. Create a Google Cloud service account and download its JSON key.
2. Save the key as `credentials.json` in the project folder (it is git-ignored).
3. Share your target spreadsheet with the service account's email.
4. Set `WB_SPREADSHEET_ID` in `.env`.

If no spreadsheet ID is set, the parser simply skips the Sheets step and still writes the Excel file.

## Usage

1. Put your article numbers in column A of `articles.xlsx`, one per row, with a header in the first row.
2. Run:

```bash
python wb_parser.py
```

Results are written to `wb_results.xlsx`.

## Output

| Column | Description |
| --- | --- |
| Артикул | Article number |
| Название | Product name |
| Бренд | Brand |
| Цена (руб) | Current price, RUB |
| Цена без скидки | Original price, RUB |
| Скидка % | Discount |
| Рейтинг | Rating |
| Отзывы | Review count |
| Остаток | Stock |
| Дата обновления | Last updated |
| Ссылка | Product link |

## Project structure

```
wb-parser/
├── wb_parser.py        # the parser
├── articles.xlsx       # input: article numbers
├── requirements.txt    # dependencies
├── .env.example        # config template
├── .gitignore
└── README.md
```

## Note

This is a learning project built to study how product data is distributed across a large e-commerce backend. It reads only publicly available data and paces its requests to stay gentle on the server.

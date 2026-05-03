# Cisco EoX Lookup

Web app to bulk-query Cisco End-of-Life status for product SKUs using the Cisco Support APIs. Produces a formatted Excel report with EoX dates, replacement PIDs, and software suggestions.

## Requirements

- Python 3.10+
- A Cisco API Console account at https://apiconsole.cisco.com with access to:
  - EoX V5 API
  - Product Information API
  - Software Suggestion API  
  *(requires active SNTC or PSS contract)*

## Setup

```bash
# 1. Create a .env file with your Cisco API credentials
echo "CLIENT_ID=your-client-id" > .env
echo "CLIENT_SECRET=your-client-secret" >> .env

# 2. Install dependencies (use the venv — requests/flask are not in system Python)
uv sync
# or: pip install -r requirements.txt

# 3. Run the app
.venv/bin/python3 app.py
```

Open **http://localhost:5001** in your browser.

## Usage

1. Enter (or confirm pre-filled) **Client ID** and **Client Secret**
2. Paste PIDs in the text area (comma or newline separated) **or** upload an Excel file with PIDs in column A
3. Optionally select a column preset to control which columns appear in the output
4. Click **Run** — a live progress bar shows each pipeline phase
5. Download the generated `.xlsx` when complete

## Features

- **Bulk lookup** — batches up to 20 PIDs per EoX/Info call, 10 per SW call
- **Replacement chain resolution** — follows `MigrationProductId` / `MigrationProductName` chains to the terminal replacement PID
- **SW suggestions** — recommended software release, train, and lifecycle per PID (subscription required)
- **Excel append mode** — when an Excel file is uploaded, EoX columns are appended to the right of the source sheet
- **Column presets** — save and reorder output column sets via the UI
- **History** — past runs are saved locally with summary stats
- **Compare** — diff two generated Excel files to highlight changed, added, and removed rows

## Output columns

| Column | Source |
|---|---|
| Product Name | Product Info API |
| Product Description | EoX API |
| Product Type / Series / Release Date | Product Info API |
| EoX Status | Derived (`End of Life` / `End of Sale` / `Active` / `Invalid SKU` / `No Record`) |
| End of Sale / SW Maintenance / Support / Vuln Support | EoX API |
| Replacement PID | EoX API (terminal node of migration chain) |
| Bulletin URL | EoX API |
| SW Type / Suggested Release / Train / Lifecycle / SW Release Date | SW Suggestion API |

## Status logic

| Status | Condition |
|---|---|
| End of Life | `LastDateOfSupport` date is set |
| End of Sale | `EndOfSaleDate` date is set (no LDoS) |
| Active | No EoX record, but product confirmed via Product Info |
| Invalid SKU | API returned `SSA_ERR_021` and product not found in Product Info |
| No Record | No EoX data and no Product Info entry |

## Debug routes

```
GET /api/debug/eox/<sku>   — raw EoX API response
GET /api/debug/hw/<sku>    — raw Product Info response
GET /api/debug/sw/<sku>    — raw SW Suggestion response
```

## Notes

- OAuth2 tokens are cached server-side until expiry
- PIDs containing `/` (e.g. `ASA5506/K9`) are resolved via MDF ID lookup before SW suggestion
- Persistent data stored in `data/history.json` and `data/presets.json`

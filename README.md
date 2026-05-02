# Cisco EoX Lookup — Flask App

Query Cisco End-of-Life status for any product SKU using the EoX V5 Support API.

## Requirements

- Python 3.10+
- A Cisco API Console account with **EoX V5 API** access  
  (requires active SNTC or PSS contract)

## Setup

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the app
python app.py
```

Then open **http://localhost:5000** in your browser.

## Usage

1. Enter your **Client ID** and **Client Secret** from https://apiconsole.cisco.com
2. Enter one or more SKUs (comma or newline separated, up to 20 per request)
3. Click **Query EoX API**

## API Endpoints

| Method | Path       | Description                          |
|--------|------------|--------------------------------------|
| GET    | `/`        | Web UI                               |
| POST   | `/api/eox` | Query EoX by SKU(s), returns JSON    |

### POST /api/eox — Request body

```json
{
  "client_id":     "your-client-id",
  "client_secret": "your-client-secret",
  "skus":          "ASA5505-BUN-K9, WS-C3750X-48P-S"
}
```

### Response

```json
{
  "records": [
    {
      "EOLProductID": "ASA5505-BUN-K9",
      "ProductIDDescription": "ASA 5505 Appliance...",
      "EndOfSaleDate":            { "value": "2017-08-25" },
      "EndOfSWMaintenanceReleases": { "value": "2018-08-25" },
      "LastDateOfSupport":         { "value": "2022-08-31" },
      "EOXMigrationDetails": {
        "MigrationProductId": "ASA5506-K9",
        "MigrationProductName": "ASA 5506-X..."
      },
      ...
    }
  ]
}
```

## Status Logic

| Status       | Condition                              |
|--------------|----------------------------------------|
| End of Life  | `LastDateOfSupport` is in the past     |
| End of Sale  | `EndOfSaleDate` is in the past         |
| Active       | Neither date has passed yet            |
| No Record    | SKU not found in Cisco EoX database    |

## Notes

- OAuth2 tokens are cached server-side until expiry (no repeated auth calls)
- Requests are batched automatically (max 20 PIDs per Cisco API call)
- The app proxies Cisco's API server-side, solving the CORS issue from browser-direct calls

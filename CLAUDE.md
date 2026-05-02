# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Run the app (must use the .venv interpreter — requests/flask are not in system Python)
.venv/bin/python3 app.py

# Install / sync dependencies
.venv/bin/pip install -r requirements.txt
# or with uv:
uv sync
```

The app runs on **http://localhost:5001** (debug mode, port set in `app.run`).

Credentials are loaded from `.env` (`CLIENT_ID`, `CLIENT_SECRET`) and pre-filled in the UI. There are no tests and no lint step.

## Architecture

Single-file Flask backend (`app.py`) + one Jinja template (`templates/index.html`). No database — persistent data is stored in `data/history.json` and `data/presets.json`.

### Request flow

1. **`POST /api/start`** — parses PIDs (text or Excel upload), starts a background `threading.Thread` running `run_job`, returns `job_id`.
2. **`GET /api/progress/<job_id>`** — SSE stream; browser polls this while `run_job` appends `(event, data)` tuples to `_job_events[job_id]`.
3. **`GET /api/download/<job_id>`** — serves the in-memory `.xlsx` buffer and cleans up job state.

### `run_job` pipeline (sequential phases)

| Phase | Progress range | What happens |
|---|---|---|
| Auth | 5% | `get_access_token` — OAuth2 client_credentials, cached server-side |
| 3 parallel API calls | 5–85% | `fetch_eox`, `fetch_product_info`, `fetch_sw_suggestions` via `ThreadPoolExecutor(3)` |
| SW suffix fallback | 86–89% | `_sw_suffix_fallback` retries HW PIDs that got no SW hit with `-A`/`-E` suffixes |
| Chain resolution | 89–92% | `enrich_replacement_chains` BFS-prefetches all replacement nodes, then resolves to `_final_replacement` |
| Excel build | 92%+ | `build_excel` — either standalone workbook or columns appended to source workbook |

### Key data structures

- **`eox_map`** `dict[str, dict]` — keyed by `EOLProductID.upper()` (the canonical key the Cisco API returns). **Never** store a record under a different queried PID as an alias — this corrupts chain resolution with wrong data (see `_fetch_one_eox` comments).
- **`info_map`** `dict[str, dict]` — keyed by `product_id.upper()`.
- **`sw_map`** `dict[str, dict]` — keyed by `basePID.upper()`; `None` means the SW Suggestion API returned 403 (no subscription).
- **`_final_replacement`** — synthetic key written into `eox_map[pid]` by `enrich_replacement_chains`; contains the terminal PID in the `MigrationProductId` chain.

### Cisco API details

| Constant | URL | Batch limit |
|---|---|---|
| `EOX_BASE` | `apix.cisco.com/supporttools/eox/rest/5` | 20 PIDs |
| `PRODUCT_INFO` | `apix.cisco.com/product/v1/information/product_ids` | 20 PIDs |
| `SW_SUGGEST` | `apix.cisco.com/software/suggestion/v2/.../productIds` | 10 PIDs |
| `SW_SUGGEST_MDF` | `…/mdfIds` | 10 MDF IDs |
| `PRODUCT_MDF` | `apix.cisco.com/product/v1/information/product_ids_mdf` | 5 PIDs |

PIDs containing `/` (e.g. `ASA5506/K9`) cannot be batch-queried for SW — they go through an MDF ID lookup first (`PRODUCT_MDF` → `SW_SUGGEST_MDF`).

All API calls use `_get_with_retry` (3 attempts, backoff on 429/5xx). `fetch_eox` and `fetch_product_info` also run `_fetch_one_eox` / `_fetch_one_product` as a sequential per-PID fallback for anything the batch missed.

### Excel output (`build_excel`)

Column catalogue lives in `ALL_COLUMNS` — a list of `(name, extractor_fn)` pairs where the extractor receives `(pid, eox_rec, info_rec, sw_rec)`. `_COL_WIDTHS` and `_CENTER_COLS` control formatting. Column selection (the "Columns" preset in the UI) arrives as a JSON array of `{col, label}` objects in `col_spec`.

When an Excel file was uploaded as input, `build_excel` appends EoX columns to the right of the source sheet (source workbook passed as `source_wb`). For text input it creates a standalone workbook.

Output filename pattern: `eox_results_YYYYMMDD_HHMM.xlsx` (text) or `<original>_eox_YYYYMMDD_HHMM.xlsx` (file upload).

### Status logic

`_resolved_status(eox_rec, info_rec)` is the single source of truth. It promotes `SSA_ERR_021_Pid` ("Invalid SKU") and missing EoX records to `"Active"` when `product_info` confirms the product exists. Status values: `"End of Life"`, `"End of Sale"`, `"Active"`, `"Invalid SKU"`, `"No Record"`.

### Debug routes

- `GET /api/debug/eox/<sku>` — raw EoX API response
- `GET /api/debug/hw/<sku>` — raw Product Info response
- `GET /api/debug/sw/<sku>` — raw SW Suggestion response

These use credentials from `.env` directly (no UI form needed).

### Frontend (`templates/index.html`)

Single-page app with vanilla JS — no build step, no npm. Glass-morphism design (backdrop-filter, aurora glow, CSS variables). Key JS concepts:

- `setProgress(pct, msg, completed)` — updates the progress bar and step chips (`completed` is the set of finished API names).
- `renderHistEntry(e)` — renders a history card including collapsible "Recommended Hardware" chips (sorted by replacement frequency).
- `toggleHistRepl(id)` — show/hide the replacement chips list.
- Column preset editor: drag-reorderable rows stored in `data/presets.json` via `POST /api/presets`.
- History stored via `POST /api/history`; cleared via `DELETE /api/history`.

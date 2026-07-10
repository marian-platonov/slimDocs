# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

SlimDocs is a single-file Streamlit app (`app.py`, ~2900 lines) that extracts text from files or URLs and reformats it as Markdown, plain text, chunked JSON, or a queryable DuckDB database — for feeding document content into LLM workflows without blowing up the context window. There is no package structure, no test suite, and no build step; everything lives in `app.py`.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run the app (Streamlit dev server, port 8087 per .streamlit/config.toml)
streamlit run app.py
```

OCR requires the Tesseract binary installed separately from `pytesseract` (see README.md's "Install Tesseract" section) — the app degrades gracefully (`"OCR unavailable"` messages) when it's missing, it does not error out.

There are no automated tests or linter configs in this repo.

## Architecture

Everything is in `app.py`. It has four functional layers, read top to bottom:

1. **EVTX native parser (top of file)** — `wevtapi.dll` bindings via `ctypes` for Windows Event Log parsing, loaded once at import time. `_wevt` is `None` on non-Windows or if the DLL load fails, at which point `_extract_evtx` falls back to the `evtx` Python library. Any change to EVTX handling needs to account for both code paths.

2. **Extractor registry** — `_FILE_EXTRACTORS: dict` maps a file extension to a `_extract_*(path) -> str` function. Extensions are grouped in `_EXT_CATEGORIES` (Documents, Spreadsheets, Web, Archives, Images, Logs, Code, Config/Data, Videos), and `_SUPPORTED_EXTENSIONS` is derived from that dict. **To add support for a new file type**: add the extension to the right `_EXT_CATEGORIES` set, write an `_extract_xxx(path: Path) -> str` function, and register it in the dispatch table near line 810. `extract_from_file()` is the single entry point that looks up the dispatcher.
   - PDFs get special treatment: `_extract_pdf` uses `pdfplumber` for text/tables, then calls `_extract_pdf_images_ocr` which uses PyMuPDF (`fitz`) to pull embedded raster images, OCRs each via `pytesseract`, and stashes the raw image bytes in the module-level `_pdf_img_cache: dict` (keyed by `str(path)`) so the caller can write them alongside the output later. This cache is cleared at the start of every `process_files*` run and popped per-file as it's consumed — any new code path that processes PDFs needs to drain this cache or images leak into the wrong output.
   - Archives (`_extract_archive`) recursively extract text from supported inner file types (`_ARCHIVE_TEXT_EXTS`) or route them back through `_FILE_EXTRACTORS` for rich formats (e.g. a `.pdf` inside a `.zip`); nested archives are skipped to avoid recursion.

3. **Processing pipelines** — four parallel functions handle the cross product of {files, URLs} × {file-output (zip/folder), DuckDB}: `process_files`, `process_urls`, `process_files_duckdb`, `process_urls_duckdb`. All are near-duplicates that: iterate inputs, call the extractor/fetcher, compute a token estimate (`len(bytes) // 4` heuristic, not a real tokenizer), format the output (`_format_output` for md/txt/json chunking), and accumulate `successes`/`errors` lists plus a `logs` list via `_build_logs`. Keep these four in sync when changing the shared token/stats/logging shape.

4. **Streamlit UI (`main()`, bottom of file)** — sidebar drives input mode (Single File / Folder / URLs), output format (`_FORMAT_OPTIONS`: md/txt/json/duckdb), and output destination; three or four tabs render results (File Processing, Statistics & Reports, Logs, and DuckDB Explorer — the last only appears when `duckdb` format is selected). All app state lives in `st.session_state`, initialized once in `_init_session()`.
   - `_render_original_file` and the `_render_docx_preview`/`_render_pptx_preview` helpers implement an in-app "original file" viewer that is separate from and richer than the plain-text extraction shown in the other preview tabs.
   - The DuckDB output mode writes to a `documents` table (schema in `_make_duckdb_db`) and unlocks `_render_duckdb_explorer`, which has its own keyword search (with highlighted snippets via `_find_snippets`/`_render_snippet_html`), a raw SQL editor, and a generic table browser/visualizer — these query the `.duckdb` file directly with `duckdb.connect(..., read_only=True)`.
   - Every run's full result dict is kept in `st.session_state.session_history` (keyed by `run_ts`), not just the latest — Statistics & Reports shows `session_history[selected_session_ts]` when set, else the latest `results`. Clicking a Logs row sets `selected_session_ts` and calls `_switch_to_tab("Statistics & Reports")`, which reaches into the parent DOM via a `components.html` script to click that tab, since `st.tabs` has no public API for switching tabs from Python — depends on Streamlit's internal tab markup and may need updating on a future Streamlit upgrade.

## Key conventions

- Private/internal helpers are prefixed with `_`; only `extract_from_file`, `chunk_text`, `discover_files`, `process_files`, `process_urls`, `process_files_duckdb`, `process_urls_duckdb`, and `main` are unprefixed.
- All extractors raise on missing optional dependencies with an actionable `pip install ...` message (see `_extract_pdf`, `_extract_docx`, etc.) rather than failing silently — follow this pattern for any new extractor.
- Confluence URLs (`atlassian.net/.../pages/...`) are special-cased in `_fetch_url`/`_fetch_confluence` to use the REST API with Basic Auth from `ATLASSIAN_EMAIL`/`ATLASSIAN_API_TOKEN` env vars instead of a plain GET.
- Salesforce URLs (`*.salesforce.com` / `*.lightning.force.com`) are special-cased in `_fetch_url`/`_fetch_salesforce` to shell out to the `sf` CLI instead of a plain GET, since Salesforce requires an authenticated session the app doesn't have. Three URL shapes are handled: a single record (`sf data get record`), a bare classic record ID (object type resolved via a `KeyPrefix` lookup against `EntityDefinition`), and a custom saved list view (`ListView` lookup by `DeveloperName` + `sf api request rest .../listviews/<id>/results`, rendered as a Markdown table). Raises an actionable error (with install/login instructions) if `sf` isn't on PATH or has no authenticated default org.
- Byte decoding always goes through `_decode_bytes`, which tries `utf-8-sig`, `utf-16`, `utf-8`, then falls back to `latin-1` — use it instead of a bare `.decode()` anywhere new text is read from disk.

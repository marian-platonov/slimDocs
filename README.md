# 📄 SlimDocs

Extract text from any file or URL and reformat it as Markdown, plain text, or chunked Claude-ready JSON - with token reduction stats.

---

## What it does

SlimDocs is a Streamlit web app that strips away binary file formats and gives you clean, readable text. It is designed for AI workflows where you need to feed document content into a language model without blowing up your context window.

- Supports 50+ file extensions across documents, spreadsheets, archives, images, logs, code, and more
- Extracts embedded images from PDFs and saves them alongside the output with OCR text where available
- Fetches and extracts content directly from URLs, including Confluence and Salesforce pages
- Outputs Markdown, plain text, or chunked JSON ready for Claude or any other LLM
- Shows token counts before and after, reduction percentage, and per-file statistics

---

## Requirements

Tested with **Python 3.14.6 (64-bit)** and **Streamlit 1.59.1**. Other 3.x/1.3x+ versions likely work but haven't been verified.

---

## Getting started

### 1. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 2. Install Tesseract (for image and PDF image OCR)

> **Important:** `pytesseract` (the Python package installed above) and **Tesseract-OCR** (the actual OCR engine) are two separate things. Installing `pytesseract` alone is not enough - you must also install the Tesseract binary or OCR will silently fail.

| Component | What it is | Status after `pip install` |
|-----------|------------|---------------------------|
| `pytesseract` | Python wrapper that calls Tesseract | ✅ installed |
| Tesseract-OCR | The actual executable that does the OCR work | ❌ must be installed separately |

**Steps:**

1. Go to: https://github.com/UB-Mannheim/tesseract/wiki
2. Download and run the Windows installer (e.g. `tesseract-ocr-w64-setup-5.x.x.exe`)
3. During install, note the path - default is `C:\Program Files\Tesseract-OCR\`
4. Add `C:\Program Files\Tesseract-OCR\` to your system `PATH`:
   - Open **Start** → search **"Environment Variables"** → **Edit the system environment variables**
   - Under **System variables**, select `Path` → **Edit** → **New** → paste the path
   - Click OK on all dialogs
5. Open a **new** terminal and verify:
   ```
   tesseract --version
   ```
6. Re-run the extraction tool

> **Still seeing the error after following the steps?** The tool was likely launched before you added Tesseract to PATH. Close and reopen the tool completely (not just a page refresh) so it picks up the updated PATH environment.

### 3. Run the app

```bash
streamlit run app.py
```

This starts the Streamlit dev server and opens the app in your default browser (see [Configuration](#configuration) for the port).

### 4. (Optional) One-click launch shortcut

Instead of opening a terminal every time, create a Windows shortcut that starts the server in the background and opens the app in your default browser with a single double-click:

1. Right-click on your Desktop (or inside the project folder) → **New → Shortcut**
2. Set the **target** to:
   ```
   "C:\Program Files\Python314\pythonw.exe" -c "import subprocess; subprocess.Popen('streamlit run app.py', creationflags=subprocess.CREATE_NEW_CONSOLE)"
   ```
   (adjust the `pythonw.exe` path to match your own Python install location)
3. Set **Start in** to the project folder:
   ```
   C:\Marian\Python Projects\AI projects\slimDocs
   ```
4. Name it `slimDocs` - double-clicking it launches Streamlit in a new console window and opens the app in your browser, no manual `streamlit run` needed.

---

## Configuration

Server behavior (port, auto-reload on save, etc.) is controlled by [`.streamlit/config.toml`](.streamlit/config.toml):

```toml
[server]
port = 8087
runOnSave = true
```

Edit this file to change the port the app runs on or to disable auto-reload when `app.py` changes.

For the full list of available `config.toml` options (server, browser, theme, client, logger, and more), see the official Streamlit reference: [docs.streamlit.io/develop/api-reference/configuration/config.toml](https://docs.streamlit.io/develop/api-reference/configuration/config.toml).

---

## How to use

| Step | Action |
|------|--------|
| **1. Input mode** | Choose **Single File**, **Folder**, or **URLs** from the sidebar |
| **2. Format** | Pick output format: `md` Markdown · `txt` Plain text · `json` Chunked JSON |
| **3. Output dir** | Set destination folder (defaults to Downloads) |
| **4. Zip toggle** | ON = single `.zip` file · OFF = timestamped folder |
| **5. Extract** | Click **🚀 Extract** |

After extraction, three tabs are available for each processed file:

- **📂 Original file** - in-app rendered view of the source file
- **📄 Extracted text** - raw parsed content with token count
- **📝 Output** - final formatted output (Markdown / plain text / JSON)

---

## Supported file types

| Category | Extensions |
|----------|------------|
| **Documents** | `.pdf` `.docx` `.doc` `.pptx` `.ppt` `.odt` `.rtf` `.md` `.txt` |
| **Spreadsheets** | `.xlsx` `.xls` `.csv` `.tsv` `.ods` |
| **Web** | `.html` `.htm` |
| **Archives** | `.zip` `.tar` `.gz` `.bz2` `.rar` `.xz` `.7z` |
| **Images (OCR)** | `.png` `.jpg` `.jpeg` `.gif` `.bmp` `.tiff` `.webp` |
| **Logs** | `.log` `.out` `.err` `.evtx` |
| **Code** | `.py` `.js` `.ts` `.java` `.go` `.rs` `.cs` `.cpp` `.sql` `.sh` `.ps1` and more |
| **Config / Data** | `.json` `.yaml` `.toml` `.xml` `.ini` `.env` `.tf` `.tfvars` and more |
| **Videos** | `.mp4` `.avi` `.mov` `.mkv` and more (file metadata only) |

---

## Output formats

### Markdown (`.md`)
Human-readable output with headings, tables, and structure preserved. Best for reading and sharing.

### Plain text (`.txt`)
Flat text with no formatting. Lowest token overhead.

### Chunked JSON (`.json`)
Array of objects with `id`, `content`, and `metadata.source` fields. Each chunk is ~2 000 characters with 200-character overlap. Designed for direct use in retrieval pipelines and Claude prompts.

```json
[
  {
    "id": "report_chunk_0",
    "content": "...",
    "metadata": { "source": "report.pdf" }
  }
]
```

---

## PDF extraction

PDFs get full treatment:

- Text extracted page by page via **pdfplumber**, including structured tables as Markdown
- Embedded raster images extracted via **PyMuPDF** and saved to a `{filename}_images/` subfolder alongside the output file, with Markdown image references embedded in the text
- OCR applied to each embedded image via **Tesseract** (if installed)
- Images smaller than 50×50 px (icons, bullets) are skipped automatically

---

## URL mode

Paste one or more URLs - one per line or comma-separated. SlimDocs fetches the page and extracts its main content using **trafilatura**.

> **Note:** trafilatura parses static HTML only - it doesn't execute JavaScript and doesn't follow client-side redirects. Two kinds of URLs return placeholder text instead of real content, and SlimDocs detects both and reports them as extraction errors rather than silently "succeeding" with junk:
> - A page that renders entirely client-side (a React/Angular/Vue single-page app with no server-rendered content) - only its `<noscript>` "JavaScript is required" fallback is visible.
> - An automatic-redirect interstitial (e.g. Google Search's own "click here if you are not redirected" stub, or any `<meta http-equiv="refresh">` bounce page) - the real destination is never reached.

**Confluence support:** if the `ATLASSIAN_EMAIL` and `ATLASSIAN_API_TOKEN` environment variables are set, Confluence page URLs are fetched via the REST API with authentication.

```bash
set ATLASSIAN_EMAIL=you@company.com
set ATLASSIAN_API_TOKEN=your_token_here
```

**Salesforce support:** Salesforce URLs (`*.salesforce.com` / `*.lightning.force.com`) are fetched via the [Salesforce CLI](https://developer.salesforce.com/tools/salesforcecli) (`sf`) instead of a plain HTTP request. A plain request only ever returns Salesforce's login page (there's no browser session to reuse), so SlimDocs shells out to the `sf` CLI's already-authenticated org instead.

**Setup - one-time, per machine:**

1. **Install the CLI** (requires [Node.js](https://nodejs.org/)):
   ```bash
   npm install --global @salesforce/cli
   ```
   Verify it's on PATH:
   ```bash
   sf --version
   ```
2. **Authenticate to your org** - this opens a browser window for the standard Salesforce login/SSO flow, then stores the session under the alias you choose:
   ```bash
   sf org login web -a UiPath --set-default
   ```
   - `-a UiPath` names this org alias `UiPath` (pick any name)
   - `--set-default` makes it the CLI's default org, so SlimDocs doesn't need to know the alias
3. **Verify the session:**
   ```bash
   sf org display
   ```
   Should print `connectedStatus: Connected` and the org's instance URL. Re-run `sf org login web` any time the session expires.

If the CLI isn't installed, or there's no authenticated default org, SlimDocs shows an actionable error naming exactly which step above is missing - it won't silently fail or show a login-page scrape.

**Supported URL shapes:**
- A single record, e.g. `.../lightning/r/Case/500XXXXXXXXXXXXAAA/view` - fetched as a Markdown field list.
- A bare classic record ID, e.g. `https://yourorg.my.salesforce.com/500XXXXXXXXXXXXAAA` - object type is resolved automatically.
- A **custom** saved list view, e.g. `.../lightning/o/Case/list?filterName=My_Open_Cases` - fetched as a Markdown table (up to 2,000 rows). Built-in standard views (e.g. "Recent", "All Open Cases") aren't queryable this way - open a custom list view, or paste a single record URL instead.

---

## Statistics & Reports

The **Statistics & Reports** tab shows, for the run currently being viewed:

- Total tokens before and after, per-file reduction breakdown (bar chart), and tokens-saved-by-file-type (pie chart)
- A full results table for successfully extracted files
- An **Errors** table with the file/URL and error message for anything that failed in that run - shown even for a run that produced only errors and no successes
- An exportable extraction report (Markdown) and a running **Session Total** across every run so far, once you've done more than one

By default this shows the **latest** run. To go back and re-examine an older one, use the Logs tab below.

The **Logs** tab shows a full history of every extraction run, filterable by success / error and searchable by filename. Every run's full results stay available for the life of the browser session (not just the summary line in Logs) - **click any log row** to jump straight to that run's report in Statistics & Reports, with a banner showing which session you're viewing and a **⬅️ Back to latest** button to return.

---

## Optional dependencies reference

| Package | Version | Purpose |
|---------|---------|---------|
| `streamlit` | `>=1.38.0` | Python framework for the interactive app |
| `pdfplumber` | `>=0.10.0` | PDF text and table extraction |
| `PyMuPDF` | `>=1.24.0` | PDF page rendering and embedded image extraction |
| `python-docx` | `>=1.1.0` | DOCX and ODT parsing |
| `openpyxl` | `>=3.1.2` | XLSX parsing |
| `odfpy` | `>=1.4.1` | ODS spreadsheet parsing |
| `python-pptx` | `>=0.6.21` | PPTX slide extraction |
| `pillow` | `>=10.0.0` | Image handling |
| `pytesseract` | `>=0.3.10` | OCR (requires Tesseract binary) |
| `trafilatura` | `>=1.6.0` | HTML and URL content extraction |
| `rarfile` | `>=4.1` | RAR archive support |
| `py7zr` | `>=0.21.0` | 7-Zip archive support |
| `evtx` | `>=0.8.0` | Windows Event Log fallback parser |
| `defusedxml` | `>=0.7.1` | Safe XML parsing for ODT/EVTX (XXE / entity-expansion protection) |
| `requests` | `>=2.31.0` | URL fetching |
| `pandas` | `>=2.1.0` | DataFrames and CSV export |
| `plotly` | `>=5.18.0` | Interactive charts in the Statistics tab |

> Versions above match [`requirements.txt`](requirements.txt) - that file is the source of truth if the two ever drift.

---

## Troubleshooting

### OCR not working / "Tesseract binary not found" in output

**Symptom:** Extracted `.md` files contain lines like `OCR unavailable - Tesseract binary not found` or `OCR skipped - Tesseract not installed`.

**Cause:** `pytesseract` (the Python wrapper) and **Tesseract-OCR** (the actual OCR engine binary) are two separate installs. Having only `pytesseract` via pip is not enough.

**Fix:** Follow the Tesseract install steps in [Getting started → Step 2](#2-install-tesseract-for-image-and-pdf-image-ocr) above. After adding Tesseract to PATH, **close and reopen the tool completely** - a page refresh alone does not reload environment variables.

### App doesn't open automatically in the browser

**Symptom:** Running `streamlit run app.py` (or double-clicking the `slimDocs.exe.lnk` shortcut) starts the server in the terminal, but no browser window opens.

**Cause:** Streamlit opens whichever browser is set as the OS **default browser**. If no default browser is configured (or it's set to an app that can't handle the URL), there's nothing for it to launch.

**Fix:**
1. Check **Settings → Apps → Default apps → Web browser** (Windows) and confirm one is set
2. Recommended: set **Microsoft Edge** or **Google Chrome** as the default browser
3. If it still doesn't open, manually browse to `http://localhost:8087` (or whatever port is set in [`.streamlit/config.toml`](.streamlit/config.toml))

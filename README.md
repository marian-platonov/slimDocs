# 📄 SlimDocs

Extract text from any file or URL and reformat it as Markdown, plain text, or chunked Claude-ready JSON — with token reduction stats.

---

## What it does

SlimDocs is a Streamlit web app that strips away binary file formats and gives you clean, readable text. It is designed for AI workflows where you need to feed document content into a language model without blowing up your context window.

- Supports 50+ file extensions across documents, spreadsheets, archives, images, logs, code, and more
- Extracts embedded images from PDFs and saves them alongside the output with OCR text where available
- Fetches and extracts content directly from URLs, including Confluence pages
- Outputs Markdown, plain text, or chunked JSON ready for Claude or any other LLM
- Shows token counts before and after, reduction percentage, and per-file statistics

---

## Getting started

### 1. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 2. Install Tesseract (for image and PDF image OCR)

> **Important:** `pytesseract` (the Python package installed above) and **Tesseract-OCR** (the actual OCR engine) are two separate things. Installing `pytesseract` alone is not enough — you must also install the Tesseract binary or OCR will silently fail.

| Component | What it is | Status after `pip install` |
|-----------|------------|---------------------------|
| `pytesseract` | Python wrapper that calls Tesseract | ✅ installed |
| Tesseract-OCR | The actual executable that does the OCR work | ❌ must be installed separately |

**Steps:**

1. Go to: https://github.com/UB-Mannheim/tesseract/wiki
2. Download and run the Windows installer (e.g. `tesseract-ocr-w64-setup-5.x.x.exe`)
3. During install, note the path — default is `C:\Program Files\Tesseract-OCR\`
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
4. Name it `slimDocs` — double-clicking it launches Streamlit in a new console window and opens the app in your browser, no manual `streamlit run` needed.

---

## Configuration

Server behavior (port, auto-reload on save, etc.) is controlled by [`.streamlit/config.toml`](.streamlit/config.toml):

```toml
[server]
port = 8087
runOnSave = true
```

Edit this file to change the port the app runs on or to disable auto-reload when `app.py` changes.

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

- **📂 Original file** — in-app rendered view of the source file
- **📄 Extracted text** — raw parsed content with token count
- **📝 Output** — final formatted output (Markdown / plain text / JSON)

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

Paste one or more URLs — one per line or comma-separated. SlimDocs fetches the page and extracts its main content using **trafilatura**.

**Confluence support:** if the `ATLASSIAN_EMAIL` and `ATLASSIAN_API_TOKEN` environment variables are set, Confluence page URLs are fetched via the REST API with authentication.

```bash
set ATLASSIAN_EMAIL=you@company.com
set ATLASSIAN_API_TOKEN=your_token_here
```

---

## Statistics & Reports

The **Statistics & Reports** tab shows:

- Total tokens before and after across all processed files
- Per-file reduction breakdown (bar chart)
- Exportable extraction report (Markdown)

The **Logs** tab shows a full session history of all extraction runs, filterable by success / error and searchable by filename.

---

## Optional dependencies reference

| Package | Purpose |
|---------|---------|
| `pdfplumber` | PDF text and table extraction |
| `PyMuPDF` | PDF page rendering and embedded image extraction |
| `python-docx` | DOCX and ODT parsing |
| `openpyxl` | XLSX parsing |
| `odfpy` | ODS spreadsheet parsing |
| `python-pptx` | PPTX slide extraction |
| `pillow` | Image handling |
| `pytesseract` | OCR (requires Tesseract binary) |
| `trafilatura` | HTML and URL content extraction |
| `rarfile` | RAR archive support |
| `py7zr` | 7-Zip archive support |
| `evtx` | Windows Event Log fallback parser |
| `requests` | URL fetching |
| `pandas` | DataFrames and CSV export |
| `plotly` | Interactive charts in the Statistics tab |

---

## Troubleshooting

### OCR not working / "Tesseract binary not found" in output

**Symptom:** Extracted `.md` files contain lines like `OCR unavailable — Tesseract binary not found` or `OCR skipped — Tesseract not installed`.

**Cause:** `pytesseract` (the Python wrapper) and **Tesseract-OCR** (the actual OCR engine binary) are two separate installs. Having only `pytesseract` via pip is not enough.

**Fix:** Follow the Tesseract install steps in [Getting started → Step 2](#2-install-tesseract-for-image-and-pdf-image-ocr) above. After adding Tesseract to PATH, **close and reopen the tool completely** — a page refresh alone does not reload environment variables.

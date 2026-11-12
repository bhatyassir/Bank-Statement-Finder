# Statement Finder

An offline Windows desktop utility that searches your system, drives, or folders to automatically discover, classify, and organize scattered **bank, credit card, and loan statements** into clean per-bank folders — without ever moving, renaming, or modifying your originals.

100% local processing: no AI APIs, no cloud uploads, and no internet connection required when scanning files.

---

## Key Capabilities

* **System-wide or folder-targeted discovery:** Recursively scans entire drives or selected directories while automatically skipping OS, program, and cache folders (`Windows`, `Program Files`, `AppData`, `$Recycle.Bin`, `node_modules`, `.git`, `.venv`, etc.).
* **Non-destructive by design:** Original files are never moved, renamed, or altered; discovered statements are copied cleanly into a separate destination tree.
* **Multi-format parsing:**
  * **PDF** (`.pdf`) via PyMuPDF — text-based PDFs only; scanned/image and password-protected PDFs are detected and skipped cleanly.
  * **Word** (`.docx`) via `python-docx` (paragraphs and tables).
  * **Text & Bank Exports** (`.rpt`, `.rptnfs`, `.nfs`, `.txt`, `.tsv`, `.htm`, `.html`).
  * **Spreadsheets skipped intentionally:** `.xls`, `.xlsx`, `.xlsm`, and `.csv` are ignored so personal workbooks and exports are never swept in as bank statements.
* **Smart non-statement filtering:** Separates genuine bank/NBFC/co-op/credit-card statements from common accounting and financial noise — invoices, Tally/accounting books & ledgers (`Day Book`, `Vch Type`, sundry debtors/creditors, purchase/sales registers), GST portal ledgers & GSTR schedules, Form 26AS / TDS statements, equity/demat reports, loan repayment schedules, restructuring proposals, and telecom recharge/usage bills.
* **Accurate bank identification & deduplication:** Identifies 45+ Indian banks, NBFCs, rural/grameen banks, and regional co-operative banks using heading IFSC/name matching, sponsor-bank IFSC filtering, bank footer URLs/phrases, and path fallback. Identical files (matching SHA-256) across folders are copied once and logged.
* **Fast, crash-isolated worker pool:** Runs parallel worker processes across CPU threads with configurable per-file timeouts so a hung or malformed file is skipped and logged without stopping the scan.
* **Two ways to run:** Desktop GUI window (`Statement Finder.exe` / `Statement Finder.bat`) plus optional headless CLI batch mode (`--run <source> <dest> [--scan-only]`).

---

## Quick Start

### Option 1: Standalone Windows Executable (Recommended — No Python Needed)

Download **`Statement Finder.exe`** from the repository's **Releases** page (or build it yourself using the spec below) and double-click it to launch.

### Option 2: Run from Python Source

1. Requires **Python 3** with `tkinter` included.
2. Double-click **`Statement Finder.bat`** (automatically installs `pymupdf` and `python-docx` on first launch), or run manually:
   ```powershell
   python -m pip install -r requirements.txt
   python statement_finder.py
   ```

### Using the App

1. **Look in:** Pick the folder or drive on your system to search.
2. **Copy to:** Pick an empty destination folder (if placed inside the source folder, it is automatically excluded from scanning).
3. *(Optional)* Tick **Scan only** to preview findings and generate the CSV report without copying any files, or uncheck `(not sure)` copying if you only want high-confidence matches.
4. Click **Find statements**. Progress, per-bank counts, and a live log appear in the window.
5. Click **Open the copies folder** or **Open the report** when finished.

### Headless / Scripted Mode

Run a scan directly without opening the window:

```powershell
"Statement Finder.exe" --run "C:\Path\To\Search" "C:\Path\To\Output"
"Statement Finder.exe" --run "C:\Path\To\Search" "C:\Path\To\Output" --scan-only
```

---

## Organized Output Structure

```text
Copy to\
├── HDFC Bank\
│   ├── statement_apr.pdf
│   └── (not sure)\                     low-row or borderline matches for quick review
│       └── short_excerpt.rpt
├── ICICI Bank\
├── J&K Bank\
├── State Bank of India\
├── _Bank not identified\               valid statement with no bank clue in text or path
├── _Credit card statements\
│   └── HDFC Bank\                      credit card statements kept separate
└── Statement Finder report 2026-11-12_120000.csv
```

The generated UTF-8 BOM CSV report opens cleanly in Excel and records for every found or unreadable file: result status, detected bank, certainty (`Sure` / `Not sure`), account vs. credit card type, heuristic score, matched transaction row count, file name, original folder location, destination path, classification reasons or error, size, and file modification date.

---

## How Classification Works

1. Reads the first 4 pages of each PDF (or up to 400 KB of text / up to 6 Word tables).
2. Counts rows that start with a transaction date (`DD/MM/YYYY`, `DD-Mon-YYYY`, ISO dates, comma-separated `MM,DD,YYYY`, wrapped-year rows) and end with a 2-decimal currency amount (including multi-line wrapped narrations).
3. Scores statement headers (`Statement of Account`, column labels like `Narration`, `Withdrawal`, `Deposit`, `Balance`, account metadata, and banking keywords like `UPI`, `NEFT`, `RTGS`, `IMPS`, `CHQ`, `Cr`/`Dr`), allowing headerless raw dumps when dense transaction tables contain unmistakable banking signals.
4. Applies negative rules and document-level exclusions against invoices, ledgers, tax statements, telecom bills, and repayment schedules before assigning `'statement'`, `'likely'`, or `'not'`.
5. Resolves the bank in priority order: heading IFSC & bank name (ignoring customer address mentions and sponsor-bank sub-member/virtual UPI IFSC prefixes) -> bank websites & template fingerprints on page headers/footers -> non-transaction footer lines -> filename / parent folder fallback.

---

## Building the Standalone Executable

Built with PyInstaller (`6.x`) from [Statement Finder.spec](Statement%20Finder.spec):

```powershell
python -m pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --clean --workpath "%TEMP%\sf_build" --distpath . "Statement Finder.spec"
```

---

## Running Tests

The unit and integration suite uses synthetic, made-up statements only (no real client data):

```powershell
python -W default test_statement_finder.py
```

---

## Project Structure

| File | Purpose |
|---|---|
| `statement_finder.py` | Tkinter desktop GUI & optional `--run` headless entrypoint |
| `finder_core.py` | File walker, PDF/DOCX/HTML/text extractors, classifier engine, worker pool, deduplicating copier & CSV reporter |
| `Statement Finder.spec` | PyInstaller single-file Windows build configuration |
| `Statement Finder.bat` | One-click Windows launcher from source (auto-installs dependencies) |
| `requirements.txt` | Runtime Python dependencies (`pymupdf`, `python-docx`) |
| `test_statement_finder.py` | Automated test suite (48 tests covering classification, readers, copying, timeouts/crashes, CLI & GUI) |

User folder preferences and unexpected error logs are stored locally in `%APPDATA%\StatementFinder\`.

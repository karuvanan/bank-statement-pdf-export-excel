# PDF Export Excel

Maybank / MBB PDF bank statement to Excel converter for Windows. The project can run as a Tkinter desktop GUI, a Gradio Web UI, or a command-line converter.

## 1. Project Name Suggestion

Recommended project name:

`PDF Export Excel`

Other possible names:

- `MBB PDF Statement to Excel`
- `Bank Statement OCR to Excel`
- `Maybank PDF Excel Exporter`

## 2. GitHub Repository Name Suggestion

Recommended repository name:

`mbb-pdf-statement-to-excel`

Other possible repository names:

- `pdf-export-excel`
- `bank-statement-ocr-excel`
- `maybank-pdf-to-excel`

## 3. Project Introduction

This project converts Maybank / MBB PDF bank statements into formatted Excel workbooks. It uses Windows OCR through PowerShell to read statement pages, parses transaction rows, splits descriptions into clean Excel columns, and exports a workbook with summary and transaction sheets.

The desktop GUI starts the Gradio Web UI in the background automatically. The GUI shows the Web UI IP address, port number, conversion logs, and backend Web UI logs in one window.

## 4. Main Features

- Convert one folder of PDF bank statements into one Excel workbook.
- Tkinter desktop GUI entry point: `run_tk_gui.py`.
- Gradio Web UI entry point: `run_gradio.py`.
- Command-line entry point: `convert_cli.py`.
- Uses Windows OCR through PowerShell.
- Outputs Excel workbook with `Summary`, `All Transactions`, and one worksheet per PDF file.
- Splits transaction description into `Desc` and `Details Payee Reference`.
- If `Details Payee Reference` is empty, it is filled with the same value as `Desc`.
- Uses one Excel `Date` column formatted as `dd/mm/yyyy`.
- Infers missing amount values from statement balance movement when possible.
- Highlights missing `Amount`, `Debit`, and `Credit` cells in yellow when review is needed.
- Writes output to the project `outputs/pdf_export_excel` folder.

## 5. Main Files And Folders

| Path | Description |
| --- | --- |
| `run_tk_gui.py` | Starts the desktop GUI. This is the easiest entry point for normal use. |
| `run_gradio.py` | Starts the Gradio Web UI directly. |
| `convert_cli.py` | Runs conversion from command line. |
| `requirements.txt` | Python package list. |
| `pdf_excel_app/converter.py` | Main conversion workflow: render PDF pages, run OCR, parse transactions, export Excel. |
| `pdf_excel_app/excel_export.py` | Excel workbook creation, columns, formatting, formulas, highlighting. |
| `pdf_excel_app/gradio_app.py` | Gradio Web UI. |
| `pdf_excel_app/tk_gui.py` | Tkinter desktop GUI and background Gradio server log window. |
| `scripts/ocr_pages.ps1` | PowerShell OCR script using Windows OCR. |
| `scripts/parse_mbb_ocr.py` | OCR text parser for Maybank / MBB statements. |
| `scripts/render_pdfs.py` | Helper script for rendering PDF pages. |
| `scripts/build_mbb_excel.mjs` | Development helper for workbook building. Normal users do not need this. |
| `outputs/` | Generated Excel output folder. Do not upload to GitHub. |
| `tmp/` | Temporary conversion files. Do not upload to GitHub. |
| `tmp_pdf_images/` | Temporary rendered PDF images. Do not upload to GitHub. |

## 6. Environment Requirements

### Required

- Windows 10 or Windows 11
- Windows PowerShell
- Python 3.11 or newer recommended
- VS Code recommended for editing
- Internet connection for first-time `pip install`

### Python Packages

Install from `requirements.txt`:

- `pypdfium2`
- `Pillow`
- `openpyxl`
- `gradio`
- `httpx`
- other Gradio support packages

### Not Required

- PHP is not required.
- MySQL is not required.
- SQLite is not required.
- Composer is not required.
- npm is not required for normal use.
- XAMPP is not required.

### Optional

- Node.js is only useful if you want to run or modify the development helper `scripts/build_mbb_excel.mjs`.
- XAMPP can be installed on the same PC, but this project does not use Apache, PHP, or MySQL.

## 7. Installation Steps

Open PowerShell in the project folder:

```powershell
cd "C:\Users\karuv\OneDrive\Documents\pdf export excel"
```

Check Python:

```powershell
py --version
```

Install Python dependencies:

```powershell
py -m pip install -r requirements.txt
```

Start the desktop GUI:

```powershell
py run_tk_gui.py
```

Start only the Gradio Web UI:

```powershell
py run_gradio.py
```

Run by command line:

```powershell
py convert_cli.py "C:\path\to\pdf-folder"
```

## 8. First Upload To GitHub

Before uploading, delete or move all real customer documents, bank statements, Excel outputs, database backups, and password files from the project folder.

Initialize Git:

```powershell
cd "C:\Users\karuv\OneDrive\Documents\pdf export excel"
git init
git status
```

Add project files:

```powershell
git add README.md .gitignore requirements.txt run_tk_gui.py run_gradio.py convert_cli.py pdf_excel_app scripts
git status
```

Create the first commit:

```powershell
git commit -m "Initial commit"
```

Set the branch name:

```powershell
git branch -M main
```

Connect to your GitHub repository:

```powershell
git remote add origin https://github.com/YOUR_USERNAME/mbb-pdf-statement-to-excel.git
```

Push to GitHub:

```powershell
git push -u origin main
```

## 9. Update GitHub After Future Changes

Check changed files:

```powershell
cd "C:\Users\karuv\OneDrive\Documents\pdf export excel"
git status
```

Add changes:

```powershell
git add .
```

Commit changes:

```powershell
git commit -m "Update project"
```

Push changes:

```powershell
git push
```

## 10. Suggested `.gitignore`

This project should not upload real PDF bank statements, generated Excel files, local databases, logs, cache, temporary OCR files, local secrets, `vendor`, or `node_modules`.

The repository should keep source code, `README.md`, `requirements.txt`, and public configuration examples only.

## 11. Files That Should Not Be Uploaded

Do not upload:

- `.env`
- API keys, tokens, passwords
- real customer PDF files
- real bank statements
- real Excel financial exports
- database files such as `.sqlite`, `.db`
- database backups
- logs
- cache files
- temporary OCR files
- `outputs/`
- `tmp/`
- `tmp_pdf_images/`
- `vendor/`
- `node_modules/`
- ZIP / RAR files containing real data

## 12. Usage

Desktop GUI:

```powershell
py run_tk_gui.py
```

Gradio Web UI:

```powershell
py run_gradio.py
```

Command line:

```powershell
py convert_cli.py "C:\path\to\pdf-folder"
```

Generated files are saved under:

```text
outputs/pdf_export_excel
```

## 13. Sync If GitHub Repository Already Exists

If the GitHub repository already exists and your local folder is not a Git repo yet:

```powershell
cd "C:\Users\karuv\OneDrive\Documents\pdf export excel"
git init
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/mbb-pdf-statement-to-excel.git
git pull origin main --allow-unrelated-histories
git add .
git commit -m "Sync local project"
git push -u origin main
```

## 14. Fix `remote origin already exists`

Check current remote:

```powershell
git remote -v
```

Change the remote URL:

```powershell
git remote set-url origin https://github.com/YOUR_USERNAME/mbb-pdf-statement-to-excel.git
```

If you want to remove and add it again:

```powershell
git remote remove origin
git remote add origin https://github.com/YOUR_USERNAME/mbb-pdf-statement-to-excel.git
```

## 15. Fix `push rejected` Or `non-fast-forward`

Pull GitHub changes first:

```powershell
git pull --rebase origin main
```

If there are conflicts, open the files in VS Code, fix them, then run:

```powershell
git status
git add .
git rebase --continue
git push
```

If Git says histories are unrelated:

```powershell
git pull origin main --allow-unrelated-histories
git push
```

Avoid force push unless you fully understand that it can overwrite GitHub history.

## 16. If Local Project Already Has A `.git` Folder

Check current Git state:

```powershell
cd "C:\Users\karuv\OneDrive\Documents\pdf export excel"
git status
git branch
git remote -v
```

If no remote is set:

```powershell
git remote add origin https://github.com/YOUR_USERNAME/mbb-pdf-statement-to-excel.git
```

If remote is wrong:

```powershell
git remote set-url origin https://github.com/YOUR_USERNAME/mbb-pdf-statement-to-excel.git
```

Then push:

```powershell
git push -u origin main
```

## 17. If GitHub Repository Is Empty

Use this flow:

```powershell
cd "C:\Users\karuv\OneDrive\Documents\pdf export excel"
git init
git add README.md .gitignore requirements.txt run_tk_gui.py run_gradio.py convert_cli.py pdf_excel_app scripts
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/mbb-pdf-statement-to-excel.git
git push -u origin main
```

## 18. If GitHub Repository Already Has README.md

Pull GitHub README first:

```powershell
cd "C:\Users\karuv\OneDrive\Documents\pdf export excel"
git init
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/mbb-pdf-statement-to-excel.git
git pull origin main --allow-unrelated-histories
```

If there are conflicts, fix them in VS Code. Then:

```powershell
git add .
git commit -m "Merge GitHub README with local project"
git push -u origin main
```

## GitHub Upload Checklist

- README.md has been generated or updated.
- `.gitignore` has been generated.
- Real customer documents have been removed.
- Real financial documents have been removed.
- Database backups have been removed.
- `.env` is not uploaded.
- API keys, tokens, and passwords are not uploaded.
- `outputs/`, `tmp/`, and `tmp_pdf_images/` are not uploaded.
- GitHub repository visibility has been confirmed as public or private.
- If the repository is public, only sanitized sample data is included.

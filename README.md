# UniTrans Waybill Sorter 1.0

This Windows tool sorts downloaded PDF waybills into the same order as the Excel workbook. It can also create one bulk PDF for printing later in Chrome, Adobe Reader, or another PDF viewer, fill blank driver numbers from scanned waybill images, and remove permissions-style PDF locks from folders of PDFs.

## Use

1. Open `UniTrans Waybill Sorter.exe`.
2. Choose `Sort Waybill PDFs`.
3. Click `Choose Excel` and select the Excel file.
4. Click `Choose PDF Folder` and select any PDF inside the folder containing the PDF waybills. The app will use that PDF's folder.
5. Confirm the downloaded PDF check matches the Excel supplementary count.
6. Click `Sort Waybills`.
7. Confirm the sort summary shows the copied PDFs were verified as exact copies.
8. Click `Create Bulk PDF`.
9. Open the bulk PDF in your preferred PDF viewer and print from there.

Optional PDF extras are hidden behind `Options` and are off by default:

- Separator pages between waybills.
- Digital sequence marks on forms 1 and 4.

The program creates a new folder inside the PDF folder:

`sorted_waybills_YYYYMMDD_HHMMSS`

Inside it:

- `Sorted` contains copied PDFs named like `001_XAPU1814169.pdf`.
- `Unmatched` contains PDFs that could not be matched.
- `Bulk PDF` contains the combined PDF when you create it.
- `sort_report.csv` shows every match, missing container, duplicate, and unmatched PDF.
- `sort_report.csv` also includes SHA-256 hashes proving each sorted PDF is byte-for-byte identical to the downloaded source PDF.
- `summary.txt` shows the copy amount detected from the Excel name and the final counts.

The original Excel and PDF files are not changed.

The files in `Sorted` are exact copies of the downloaded PDFs. The bulk PDF is a new combined PDF, but when optional extras are off it uses the original PDF pages directly; it does not turn pages into screenshots or change margins.

When separator pages are enabled, the bulk PDF order is: all pages for `001`, separator, all pages for `002`, separator, and so on.

The sequence mark option stamps the Excel sequence number on pages detected as form `1` and form `4`, using the form number printed at the top-left of each waybill page.

## Unlock PDFs

Choose `Unlock PDFs` from the first screen when a folder contains PDFs that open but cannot be edited, printed, or processed because of permissions-style encryption.

1. Click `Choose PDF Folder` and select the folder.
2. Click `Unlock PDFs`.
3. Use `Open Output Folder` to view the new timestamped output folder.

The program creates:

`unlocked_pdfs_YYYYMMDD_HHMMSS`

Inside it:

- `Unlocked PDFs` contains unlocked/copied PDF files.
- `unlock_report.csv` lists every source PDF, output PDF, status, page count, and note.
- `summary.txt` shows the totals.

PDFs that open with a blank password are rewritten without encryption. PDFs that are already unlocked are copied unchanged. PDFs that require an unknown open password are skipped and listed in the report.

## Build

Run PowerShell from this folder:

```powershell
.\build.ps1
```

The standalone executable is created at:

`dist\UniTrans Waybill Sorter.exe`

# Client List Generator

A desktop Python program that reads stamping certificates from one or more searchable PDFs and creates an Excel workbook with one row per PDF. Machine details found across its certificate pages are listed in page order on separate lines within the matching cell. Blank pages are ignored.

The workbook contains one row per PDF with these columns:

1. Date of Stamping
2. Address
3. Machine Details (one line per machine: make, capacity, simplified type/class, and machine number)
4. SI (Stamping Inspector)

Rows are sorted by stamping date, oldest first. For records with the same date, rows are sorted alphabetically by Address. Records without a readable stamping date are placed after dated records.

## Requirements

- Python 3.10 or newer
- Searchable/text-based PDFs using the Tamil Nadu Legal Metrology certificate layout shown by the reference PDF

Scanned image-only PDFs need OCR before this program can read them. The extractor reads the date and client address from the certificate text and machine details from the table columns. It can handle multiple equipment certificate pages for the same client PDF, including a blank final page.

## Run

### Windows app

Double-click `run.exe` in this folder. The packaged app includes its Python runtime and required libraries, so Python does not need to be installed on the computer where you run it.

### Run from Python source

Open PowerShell in this folder and run:

```powershell
py -m pip install -r requirements.txt
py client_list_generator.py
```

To rebuild the Windows executable after changing the program, install the requirements and run:

```powershell
py -m PyInstaller --onefile --windowed --name run client_list_generator.py
Copy-Item dist\run.exe .\run.exe
```

Use **Add PDF file(s)…** to select one or more PDFs. The selected file paths appear in the list. Select one or more rows and choose **Remove selected** to remove them from both the list and the generation input; **Clear all** removes every selected PDF. Choose **Create list** to combine records from the remaining files into one `.xlsx` workbook.

The app follows the Windows light/dark app theme preference and checks for changes while it is open. It creates one combined workbook for the full selection. Machine details are combined in the format `Make, Capacity, Type/Class, M/C: Machine Number`. Weight-only certificates without machine make/number use the item description and denomination instead. Inspector labels are normalized, for example, `Stamping Inspector-2, Salem` becomes `SI 2, Salem`, `Stamping Inspector/Inspector of Legal Metrology, 1st Circle, Salem` becomes `SI 1, Salem`, and `Deputy Inspector of Labour-II, Salem` becomes `DIL 2, Salem`. Distinct inspector values found across pages are listed on separate lines in the SI cell. After conversion, the app's **Failed / partial files** panel shows PDFs marked failed or partially extracted, with reasons; successful files are omitted. If there are no failures or partial extractions, the panel says so. Failed files are omitted from the Excel client list, and blank trailing pages are ignored. Review the generated workbook against the source PDFs before using it.

# Skoda Export CSV Generator

A Python GUI tool that parses Skoda commercial invoice PDFs and maps the extracted items strictly to a 114-column Logisys Export CSV template, handling calculations, item variations, and necessary trade fields (Exim Scheme, DBK Sr. No., RoDTEP).

## Tech Stack
- Python 3.12
- Tkinter (GUI)
- pdfplumber (PDF tabular extraction)
- Pillow (Logo management)
- openpyxl (HS Code Master loading)

---

## Installation

### Clone
```bash
git clone https://github.com/username/skoda-export-generator.git
cd skoda-export-generator
```

---

## Python Setup (MANDATORY)

⚠️ **IMPORTANT:** You must use a virtual environment.

1. Create a virtual environment:
```bash
python -m venv venv
```

2. Activate (REQUIRED)
**Windows:**
```cmd
venv\Scripts\activate
```
**Mac/Linux:**
```bash
source venv/bin/activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Run the application:
```bash
python skoda_export_csv.py
```

---

## Build Executable (For Desktop Windows Use)

To generate a portable `.exe` folder containing everything you need:

1. Ensure your `venv` is active.
2. Build the app using the provided Spec file. It enforces asset bundling (Logos, HS Code Master).
```bash
pyinstaller SkodaExportGenerator.spec
```

The resulting application executable will be placed in the `dist\SkodaExportGenerator\` folder.

---

## Usage

1. **Select Invoice PDF:** Pick the Skoda commercial invoice.
2. **Setup Fields:** Select the required settings (TOI, Exim Scheme, DBK, Origin District, etc.).
3. **Parse Invoice:** The system extracts the multi-line part numbers, computes unit price (`Price/100 ÷ 100`), validates HS Code dimensions, calculates SQC based on HS mappings, and populates the 114-column structure.
4. **Export:** Click `Export Logisys CSV` to save your fully compliant CSV.

Notes: If the tool fails to auto-load the 12000+ line HS Code Master, manually upload the Excel map via the designated upload button.

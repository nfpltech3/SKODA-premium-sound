# SKODA Invoice to CSV Converter

Automated tool to extract line-item data from Premium Sound Solutions (PSS) invoice PDFs for SKODA and export them to a consolidated CSV file.

## Features
- Extracts Pose, Article No, Qty, Price, Amount, HS Code, and Country of Origin.
- Handles multi-page invoices.
- Nagarkot branded graphical user interface.
- Batch processing of multiple PDFs.

## Tech Stack
- **Python 3.10+**
- **Tkinter** (GUI)
- **pdfplumber** (PDF Extraction)
- **Pandas** (Data Formatting & CSV Export)
- **Pillow** (Logo support)

---

## Installation

### 1. Clone or Download
```bash
git clone https://github.com/nfpltech3/SKODA-premium-sound.git
cd "Pune - Premium Sound"
```

### 2. Setup Virtual Environment (MANDATORY)
```bash
python -m venv venv
```

### 3. Activate Virtual Environment
**Windows:**
```pwsh
venv\Scripts\activate
```

**Mac/Linux:**
```bash
source venv/bin/activate
```

### 4. Install Dependencies
```bash
pip install -r requirements.txt
```

---

## Usage
1. Run the application:
   ```bash
   python SKODA_Invoice_to_CSV.py
   ```
2. Click **Select Files** to pick one or more invoice PDFs.
3. Review the **Data Preview** in the table.
4. Click **Process & Export** to generate the CSV.
5. The output CSV will be saved in the selected output folder (defaults to the source folder).

---

## Building the Executable
To create a standalone `.exe` for Windows:

1. Ensure venv is active and `pyinstaller` is installed.
2. Run the build command using the spec file:
   ```bash
   pyinstaller SKODA_Invoice_to_CSV.spec
   ```
3. Find the executable in the `dist/` folder.

---

## Project Structure
- `SKODA_Invoice_to_CSV.py`: Main application logic and GUI.
- `Nagarkot Logo.png`: Brand asset.
- `SKODA_Invoice_to_CSV.spec`: PyInstaller configuration.
- `.gitignore`: System files exclusion.
- `requirements.txt`: Project dependencies.

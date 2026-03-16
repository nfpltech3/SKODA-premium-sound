# SKODA Invoice to CSV Converter — User Guide

## Introduction
The **SKODA Invoice to CSV Converter** is a dedicated desktop utility designed for Nagarkot Forwarders. It automates the extraction of line-item details from **Premium Sound Solutions (PSS)** invoice PDFs specifically for SKODA inventory, transforming complex PDF data into a structured CSV format ready for import or analysis.

## How to Use

### 1. Launching the App
- **For End-Users:** Locate the `dist` folder and double-click `SKODA_Invoice_to_CSV.exe`.
- **For Developers:** Activate the virtual environment and run `python SKODA_Invoice_to_CSV.py`.

### 2. The Workflow (Step-by-Step)
1.  **Select Invoices**: Click the **📂 Select Files** button. You can select multiple PDFs at once.
    - *Note: Only standard PSS invoice formats are supported. The app looks for specific "Pos", "Article No", and "Commodity Code" patterns.*
2.  **Define Save Location**: By default, the output is saved in the same folder as your first PDF. Click **📁 Change Folder** if you wish to save the results elsewhere.
3.  **Preview Data**: Once files are selected, click **⚙ Process & Export**. The "Data Preview" table will populate, allowing you to verify the extracted details before they are finalized.
4.  **Final Export**: After the extraction completes, a success message will appear showing the total items extracted and the location of the generated `.csv` file.

## Interface Reference

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| **Select Files** | Opens a dialog to pick invoice PDFs. | `.pdf` files |
| **Clear List** | Removes all selected files from the queue. | N/A |
| **Change Folder** | Sets the destination for the output CSV. | System Folder Path |
| **Process & Export** | Runs the extraction engine and saves the CSV. | N/A |
| **Data Preview** | Displays Article No, Qty, HS Code, etc. | Tabular view |

## Troubleshooting & Validations

If you see an error, check this table:

| Message | What it means | Solution |
| :--- | :--- | :--- |
| **"Please select at least one PDF file"** | No files were added to the listbox. | Click "Select Files" and choose at least one PDF. |
| **"No line-item data could be extracted"** | The PDF format does not match the PSS layout or is a scanned image (OCR not supported). | Ensure the PDF contains selectable text and follows the expected PSS structure. |
| **"Failed/empty: [filename]"** | A specific file in your list was either corrupt or did not contain valid item blocks. | Check the specific PDF file for data. It might be a cover sheet without line items. |
| **"Processing Error"** | An unexpected system or file access error occurred. | Ensure the output CSV is not currently open in Excel and that you have write permissions to the folder. |

## Notes for Excel Users
- **Leading Zeros:** The app automatically formats Invoice Numbers as `="00..."` to ensure Excel does not strip leading zeros when you open the CSV.
- **Encoding:** Files are saved using `UTF-8 with BOM` to ensure special characters (like currency symbols) display correctly in Excel.

---
**© Nagarkot Forwarders Pvt Ltd**

# Skoda Export CSV Generator User Guide

## Introduction
The **Skoda Export CSV Generator** is a specialized, one-click Windows application built for Nagarkot Forwarders Pvt Ltd. It automatically scans multi-page Skoda commercial invoice PDFs, perfectly extracts and formats complicated data (like multi-line part numbers and split HS Codes), calculates proper pricing, maps everything against the HS Code Master rules, and directly outputs a ready-to-import **114-column CSV file for Logisys**.

## How to Use

### 1. Launching the App
1. Locate the `SkodaExportGenerator.exe` file on your computer (typically inside the `dist` folder if you built it, or wherever it was provided to you).
2. Double-click the `.exe` file to launch the application.

### 2. The Workflow (Step-by-Step)
1. **Select Invoice PDF**: Click the `📄 Select Invoice PDF` button and choose your Skoda commercial invoice 
   - *Note: The file must be a standard readable PDF from Skoda.*
2. **Review HS Master Options**: The tool automatically loads the included `HSN Code - Master_Data_Feb_2026.xlsx`. If it successfully loads, you will see green text saying `HS Master: XXXXX codes loaded ✓`.
   - *Note: If this fails, click `🔄 Reload / Change HS Code Master` to manually locate the Excel file.*
3. **Configure Export Options**: Fill out the fields in the Left Panel.
   - Example Options: **Terms of Invoice (TOI)**, **Exim Scheme**, and **DBK Under**.
   - **Tip:** If the Country of Origin (COO) on the invoice is `IN`, the Exim Scheme automatically defaults to `19`.
4. **Parse Invoice**: Click the blue `▶ Parse Invoice` button.
   - *Note: If a selected HS code is split incorrectly in the PDF, the tool safely reconstructs both halves into a perfect 8-digit sequence.*
   - *Result*: The grid on the right will visually populate all 114 columns of Logisys data allowing you to scroll both vertically and horizontally.
5. **Export Logisys CSV**: Click the green `💾 Export Logisys CSV` button.
   - *Result*: You will be prompted to save your CSV file. It defaults to the naming convention: `Logisys_Export_[InvoiceNumber].csv`.

## Interface Reference

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| **Select Invoice PDF** | Opens a file dialog to locate the Skoda Invoice | [.pdf] file |
| **Reload HS Code Master** | Opens a window to manually override the HS mapping rules | `.xlsx` or `.xls` file |
| **Terms of Invoice (TOI)** | Radio selection for shipment terms | `C&F` / `FOB` / `CIF` / `C&I` |
| **Exim Scheme** | Radio selection. Used for Logisys Column 11 | `19` / `00` |
| **End Use** | End use of the exported items | `GNX100` / `GNX200` |
| **IGST Pay Status** | Sets Logisys column mapping for IGST | Export Under Bond... / Export Against Payment |
| **IGST Rate** | Sets IGST tax slab | `18` / `5` / `None` |
| **DBK Under** | Drawback Claim type | `Actual` / `Provision` |
| **Origin District** | Numeric code representing the origin district | Default: `490` |
| **PTAFTA Code** | Trade Agreement preference code | `GSTP` / `NCPTI` |

---

## Technical Transformation Rules (Behind the Scenes)
The application automates several Logisys requirements for you:
1. **Rate Calculation**: The tool automatically calculates the exact per-unit rate by dividing the `Price/100` column by 100.
2. **DBK Sr. No Generation**: The application takes the extracted DBK Code and automatically appends a `"B"` (e.g., `870899` becomes `870899B`).
3. **HS Code (UQC) Checking**: The tool references the HS Master Excel. If an item's HS Code requires `NOS` or `PCS`, the **SQC Quantity** uses the `Quantity` value. If the master says `KGS`, the **SQC Quantity** uses the `Net Weight`.
4. **Multi-line Part Numbers**: If a part number spans two lines (e.g. contains a variant or color code like `VKH`), the app cleanly joins them together.

---

## Troubleshooting & Validations

If you see an error, check this table:

| Message | What it means | Solution |
| :--- | :--- | :--- |
| **"Please select an invoice PDF first."** | You clicked Parse without a file. | Click `📄 Select Invoice PDF` to load a [.pdf]file first. |
| **"No data to export. Parse an invoice first."** | You clicked Export before data was ready. | Click `▶ Parse Invoice` to extract data into the preview table, then Export. |
| **"Failed to parse PDF:"** | The PDF format is unreadable, not a valid PDF, or lacks the standard tables spanning 15 columns. | Verify you selected an actual Skoda Commercial Invoice PDF, not an image or scrambled document. |
| **"[Warning Log] Item X: HS code 'YYYY' is not 8 digits"** | Despite reconstruction efforts, the HS Code found on the invoice isn't exactly 8 digits. | Review the generated CSV to ensure the HS Code (RITC) is correct for the mentioned position number. |
| **"HS Master not found — click Reload"** | The default `HSN Code - Master_Data_Feb_2026.xlsx` is missing from the folder. | Ensure the master file is located next to the `.exe`, or click `Reload` to manually select it. |
| **"Failed to write CSV:"** | The location you tried saving to is protected, or the file is open in another program (like Excel). | Close the [.csv] if it is currently open in Excel, and try saving again. |

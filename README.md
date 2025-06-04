# Automated Invoice Generation from Excel

This project provides a Python script to **automate the generation of professional-looking PDF invoices** directly from Excel spreadsheet data. It processes multiple Excel files, creating a separate PDF invoice for each, complete with product details, pricing, and a calculated total.

## Overview

The "Automated Invoice Generation" tool simplifies the invoicing process for businesses or individuals dealing with numerous transactions. Instead of manually creating each invoice, this script reads your sales data from structured Excel files and instantly generates clear, formatted PDF invoices, saving significant time and reducing errors.

## Features

* **Batch Invoice Generation:** Processes all `.xlsx` files found in a specified `Invoices/` directory, generating a unique PDF invoice for each.
* **Dynamic Invoice Details:** Automatically extracts invoice number and date from the Excel filename (e.g., `1001-2023-01-15.xlsx` becomes Invoice Nr. 1001, Date: 2023-01-15).
* **Product Itemization:** Populates the invoice with product ID, name, quantity, price per unit, and total price for each item, directly from the Excel sheet.
* **Automatic Total Calculation:** Calculates and displays the grand total for all items in the invoice.
* **Customizable Header/Footer:** Includes a bold header for invoice number and date, and a footer with a company name and logo (e.g., "PythonHow").
* **Structured PDF Layout:** Uses `fpdf` to create a well-organized PDF with clear tables and consistent formatting.
* **Output Directory:** Saves all generated PDF invoices into a dedicated `PDFs/` folder.

## Technologies Used

* Python
* `pandas` library (for reading and processing Excel files)
* `fpdf` library (for creating PDF documents)
* `glob` module (for finding multiple files based on a pattern)
* `pathlib` module (for object-oriented file path manipulation)
* `openpyxl` (indirectly used by pandas for Excel file handling)

## How It Works

The script systematically generates invoices based on your Excel data:

1.  **File Discovery:**
    * It uses `glob.glob("Invoices/*.xlsx")` to locate all Excel files within the `Invoices/` directory.

2.  **Iterate Through Invoices:**
    * For each Excel file found:
        * A new `FPDF` object is initialized for the current invoice.
        * The filename (e.g., `1001-2023-01-15`) is split to extract the invoice number and date.
        * A new page is added to the PDF.
        * The extracted invoice number and date are printed as bold headers.

3.  **Read Excel Data:**
    * The corresponding Excel file is read into a Pandas DataFrame, specifically from "Sheet 1".
    * Column names from the Excel file (e.g., `product_id`, `product_name`) are processed (replacing underscores and capitalizing) to create clean table headers for the PDF.

4.  **Generate Invoice Table:**
    * A bold header row is added to the PDF table using the processed column names.
    * The script iterates through each row of the Pandas DataFrame (each product item):
        * Each product's details (`product_id`, `product_name`, `amount_purchased`, `price_per_unit`, `total_price`) are added as a row in the PDF table.
        * The `total_price` for each item is summed up to calculate the grand total for the invoice.
    * After all items are added, a final row for the `total` is added to the table.

5.  **Final Touches & Output:**
    * The grand total is printed below the table in bold text.
    * A company name ("PythonHow") and an image (`pythonhow.png`) are added to the bottom of the invoice.
    * The generated PDF is saved to the `PDFs/` directory with a filename derived from the original Excel file (e.g., `1001-2023-01-15.pdf`).

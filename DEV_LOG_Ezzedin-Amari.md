# DYDON AI Internship 2025 Submission
## Overview
This repository contains my solution for the DYDON AI Internship 2025 task. The script extract_text_[ezdin-amari].py extracts text from files (.pdf, .docx, .xls, .xlsx) in the uploads/ folder, adhering to the requirements of processing these file types with a fallback to OCR for PDFs lacking extractable text. The approach focuses on clarity, performance, error handling, and extensibility.
Dependencies

## Python 3.8+
## Libraries:
## PyMuPDF (fitz): For efficient PDF text extraction and image conversion.
## pytesseract: For OCR on scanned PDFs or PDFs without extractable text.
## pdf2image: To convert PDF pages to images for OCR.
## python-docx: For .docx file processing.
## openpyxl: For .xls/.xlsx file processing.
## Pillow: For image handling with OCR.
## tqdm: For progress tracking during OCR.


# External Tools:
## Tesseract OCR (install via system package manager or installer).
## Poppler (required for pdf2image on Windows/Linux/macOS).



# Install dependencies:
pip install PyMuPDF pytesseract pdf2image python-docx openpyxl Pillow tqdm


## Tesseract: Install via brew install tesseract (macOS), sudo apt-get install tesseract-ocr (Linux), or download from GitHub (Windows).
## Poppler: Install via brew install poppler (macOS), sudo apt-get install poppler-utils (Linux), or download from GitHub (Windows).

## Approach
## General Strategy

The script iterates through all files in the uploads/ folder, detecting and processing supported file types (.pdf, .docx, .xls, .xlsx).
It employs a modular design with separate handling for each file type, ensuring extensibility for additional formats.
Error handling is implemented to log issues and continue processing, while performance is optimized by prioritizing native extraction over OCR.

# PDF Processing

Tool: Uses PyMuPDF (fitz) for native text extraction, chosen for its efficiency and ability to handle complex PDF structures.
Two-Pass Approach:
First pass: Extracts text directly using page.get_text() for digital PDFs.
Second pass: If no text is extracted, falls back to OCR using pytesseract on images generated with page.get_pixmap() at 300 DPI.


Progress Tracking: Employs tqdm to display progress during OCR, which can be time-consuming for multi-page PDFs.
Error Handling: Logs errors per page during OCR and overall PDF processing, ensuring the script doesn’t crash.

# DOCX Processing

Tool: Uses python-docx to extract paragraph text.
Approach: Joins all paragraph texts with newline characters for a cohesive output.
Consideration: Assumes well-formed .docx files; errors are caught and logged if files are corrupted.

# Excel Processing

Tool: Uses openpyxl with read_only=True and data_only=True for efficient handling of .xls/.xlsx files.
Approach:
Processes all sheets in the workbook, prefixing each with a header (e.g., --- Sheet: Sheet1 ---).
Iterates over rows using iter_rows(), joining non-null cell values with tabs (\t) for readability.


# Consideration: Focuses on extracting all available data, with error logging for malformed files.

# Error Handling and Logging

Uses Python’s logging module to record successes, errors, and skipped files.
Gracefully handles exceptions (e.g., corrupted files, missing Tesseract) and continues processing remaining files.

Performance Optimizations

# Prioritizes native extraction (PyMuPDF for PDFs, openpyxl for Excel) to minimize OCR usage, which is computationally intensive.
Uses read_only=True in openpyxl to optimize memory usage for large Excel files.
Processes files sequentially to manage resource consumption.

# Extensibility

The script is designed to support additional file types by adding new conditions in the extract_text function.
Libraries can be swapped (e.g., PyMuPDF for another PDF library) with minimal code changes.

# Running the Code

Place test files (.pdf, .docx, .xls, .xlsx) in the uploads/ folder.
Install dependencies as listed above.
Run the script:python extract_text_[ezdin-amari].py


Check the extracted_output/ folder for extracted text files.
Review logs in the terminal for processing details.

# Verification
To verify the script works:

Ensure uploads/ contains test files (e.g., a digital PDF, a scanned PDF, a .docx, and a .xlsx).
Run the script and confirm output files in extracted_output/:
PDFs: Text matches content (OCR for scanned PDFs).
DOCX: All paragraphs are extracted.
XLS/XLSX: Each sheet is separated with headers, and rows are tab-separated.


Test edge cases (empty files, corrupted files, unsupported formats) to confirm error handling.
On Kaggle, upload files, enable internet, install dependencies with !pip and !apt-get, and run with !python extract_text_[ezdin-amari].py.

## Future Improvements

Add support for additional file formats (e.g., .txt, .pptx).
Optimize OCR performance with multi-threading for large PDFs.
Implement data validation for Excel cells to handle special formats.

# This approach balances functionality, performance, and readability while meeting the internship requirements.


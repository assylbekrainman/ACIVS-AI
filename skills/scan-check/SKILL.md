---
name: scan-check
description: This skill analyzes scanned documents (images or PDFs) to verify them against the required supporting documents for 'Сметы расходов' (Expenditure Estimates). It uses OCR to extract text and an LLM to classify the document and check for compliance with predefined requirements.
license: Complete terms in LICENSE.txt
---

# Scan Check Skill

This skill provides a mechanism to automate the verification of supporting documents for Expenditure Estimates. It leverages Optical Character Recognition (OCR) to extract text from scanned images or PDF files and then uses a Large Language Model (LLM) to intelligently match the document content against a predefined set of requirements.

## Functionality

The `scan-check` skill performs the following:

1.  **Text Extraction**: Extracts text from image files (PNG, JPG, JPEG, TIFF, BMP) and PDF documents using OCR and PDF text extraction libraries.
2.  **Document Classification**: Identifies the relevant 'Статья Сметы Расходов' (Expenditure Article) based on the extracted document content.
3.  **Requirement Verification**: Checks if the document fulfills specific requirements outlined for the identified expenditure article.
4.  **Detailed Reporting**: Provides a structured JSON output indicating the identified article, met requirements, a confidence score, and any discrepancies or missing information.

## Usage

To use this skill, you will invoke the `scan_check.sh` script located in the `scripts/` directory, passing the path to the document you wish to verify as an argument.

```bash
/home/ubuntu/skills/scan-check/scripts/scan_check.sh <path_to_document>
```

**Example:**

```bash
/home/ubuntu/skills/scan-check/scripts/scan_check.sh /home/ubuntu/scans/invoice_for_equipment.pdf
```

The script will output a JSON object containing the verification results.

## Requirements Reference

The detailed list of document requirements for each expenditure article is stored in:

-   `/home/ubuntu/skills/scan-check/references/document_requirements.md`

This file is used by the verification script to inform the LLM's assessment. Any updates to the document requirements should be made in this Markdown file.

## Technical Details

The skill utilizes the following components:

-   **`verify_document.py`**: The core Python script that orchestrates the OCR, text extraction, and LLM interaction.
-   **`scan_check.sh`**: A simple shell wrapper to execute the Python script.
-   **`document_requirements.md`**: A reference file containing the structured document requirements.
-   **`pytesseract`**: Python wrapper for Google's Tesseract-OCR Engine for image-based text extraction.
-   **`pdfminer.six`**: For text extraction from PDF documents.
-   **`OpenAI`**: Python client for interacting with the LLM (gemini-2.5-flash) for document classification and verification.

Ensure that Tesseract OCR and its Russian language pack are installed on the system for optimal performance with Russian documents. The necessary Python libraries (`openai`, `pytesseract`, `pdfminer.six`, `Pillow`) are also required.

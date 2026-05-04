import pytesseract
from PIL import Image
from pdfminer.high_level import extract_text as pdf_extract_text
import os
from openai import OpenAI

# Initialize OpenAI client
client = OpenAI()

def ocr_image(image_path):
    """Performs OCR on an image file."""
    try:
        text = pytesseract.image_to_string(Image.open(image_path), lang='rus+eng')
        return text
    except Exception as e:
        return f"Error during OCR for image {image_path}: {e}"

def ocr_pdf(pdf_path):
    """Extracts text from a PDF file using pdfminer.six."""
    try:
        text = pdf_extract_text(pdf_path)
        return text
    except Exception as e:
        return f"Error during text extraction for PDF {pdf_path}: {e}"

def get_document_requirements(requirements_file_path):
    """Reads document requirements from a Markdown file."""
    with open(requirements_file_path, 'r', encoding='utf-8') as f:
        return f.read()

def verify_document_with_llm(document_text, requirements):
    """Verifies the document against requirements using an LLM."""
    prompt = f"""You are an expert document verifier. Your task is to analyze the provided document text and determine which category of 'Сметы расходов' (Expenditure Estimates) it belongs to, and if it fulfills any of the specified document requirements for that category. 

Here are the document requirements:
{requirements}

Here is the document text:
{document_text}

Based on the document text, identify:
1. The most likely 'Статья Сметы Расходов' (Expenditure Article) this document relates to.
2. Which specific document requirements (e.g., '1.1. Копия утвержденного штатного расписания') from that article, if any, are met by this document.
3. A confidence score (0-100) for your assessment.
4. Any discrepancies or missing information.

Provide your answer in a structured JSON format, like this:
{{
    "expenditure_article": "<Identified Article>",
    "met_requirements": [
        "<Requirement 1>",
        "<Requirement 2>"
    ],
    "confidence_score": <Score>,
    "discrepancies": "<Discrepancies or Missing Info>"
}}
If no article is identified, set 'expenditure_article' to 'N/A' and 'met_requirements' to an empty list.
"""

    try:
        response = client.chat.completions.create(
            model="gemini-2.5-flash", # Using a suitable model
            response_format={ "type": "json_object" },
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error during LLM verification: {e}"

def main(file_path):
    requirements_file = "/home/ubuntu/skills/scan-check/references/document_requirements.md"
    requirements = get_document_requirements(requirements_file)

    document_text = ""
    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension in ['.png', '.jpg', '.jpeg', '.tiff', '.bmp']:
        document_text = ocr_image(file_path)
    elif file_extension == '.pdf':
        document_text = ocr_pdf(file_path)
    else:
        return f"Unsupported file type: {file_extension}. Please provide an image or PDF."

    if "Error" in document_text:
        return document_text

    verification_result = verify_document_with_llm(document_text, requirements)
    return verification_result

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python verify_document.py <path_to_document>")
        sys.exit(1)
    
    document_path = sys.argv[1]
    result = main(document_path)
    print(result)

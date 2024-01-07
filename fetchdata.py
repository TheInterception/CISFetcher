import pdfplumber
import pandas as pd
import re

def is_appendix_line(line):
    return "Appendix" in line

def clean_recommendation(recommendation):
    pattern = re.compile(r'\((Automated|Manual|Scored|Not Scored)\)\s*\.{3,}\s*\d+|\.{10,}\s*\d+')
    return pattern.split(recommendation)[0]

def process_cis_pdf(pdf_filename, excel_filename):
    with pdfplumber.open(pdf_filename) as pdf:
        pages = pdf.pages[2:]

        recommendations = []
        extracting = False
        for page in pages:
            text = page.extract_text()
            for line in text.split('\n'):
                if is_appendix_line(line):
                    extracting = False
                    break
                if "Recommendations" in line or extracting:
                    extracting = True
                    if line.strip() != "Recommendations":
                        clean_line = clean_recommendation(line.strip())
                        if clean_line:
                            recommendations.append(clean_line)
            if not extracting:
                break

    df = pd.DataFrame(recommendations, columns=["Recommendations"])
    df.to_excel(excel_filename, index=False)
    print(f"Saved the recommendations to '{excel_filename}'")

pdf_filename = input("Enter the name of the PDF file to extract CIS from: ")
excel_filename = input("Enter the name of the generated XLSX file: ")

process_cis_pdf(pdf_filename, excel_filename)


import pdfplumber
import pandas as pd
import re

def is_appendix_line(line):
    return "Appendix" in line

def clean_recommendation(recommendation):
    # Compile a regex pattern to match the specified sequences and a page number
    pattern = re.compile(r'\((Automated|Manual|Scored|Not Scored)\)\s*\.{3,}\s*\d+|\.{10,}\s*\d+')
    # Use the regex pattern to split the recommendation and return the first part
    return pattern.split(recommendation)[0]

# Function to process the PDF and extract recommendations
def process_cis_pdf(pdf_filename, excel_filename):
    with pdfplumber.open(pdf_filename) as pdf:
        # Start reading from the table of contents
        pages = pdf.pages[2:]  # Assuming table of contents starts at page 3

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
                        if clean_line:  # Only add if the line is not empty
                            recommendations.append(clean_line)
            if not extracting:
                break

    df = pd.DataFrame(recommendations, columns=["Recommendations"])
    df.to_excel(excel_filename, index=False)
    print(f"Saved the recommendations to '{excel_filename}'")

# Prompt the user for the input PDF file and the output Excel file
pdf_filename = input("Enter the name of the PDF file to extract CIS from: ")
excel_filename = input("Enter the name of the generated XLSX file: ")

# Call the function with the provided filenames
process_cis_pdf(pdf_filename, excel_filename)


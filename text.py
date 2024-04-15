import re
import pandas as pd
from docx import Document

def extract_info_from_cv(doc_path):
    doc = Document(doc_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    full_text = "\n".join(full_text)
    
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    phone_pattern = r'(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4})'
    
    emails = re.findall(email_pattern, full_text)
    phones = re.findall(phone_pattern, full_text)
    
    return emails, phones, full_text

def save_to_excel(emails, phones, full_text):
    data = {
        'Email Address': emails,
        'Contact Number': phones,
        'Overall Content': [full_text]      }
    df = pd.DataFrame(data)
    df.to_excel('cv_information.xlsx', index=False)

def main():
    cv_path = "cv.docx"
    emails, phones, full_text = extract_info_from_cv(cv_path)
    save_to_excel(emails, phones, full_text)

if __name__ == "__main__":
    main()

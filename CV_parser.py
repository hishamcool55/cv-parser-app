import streamlit as st
import pdfplumber
import pytesseract
from PIL import Image
import io
import os
import re
import pandas as pd

# Function to extract text from a PDF with fallback to OCR for image-based content
def extract_text_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
            else:
                # Fall back to OCR if no text is detected on the page
                image = page.to_image(resolution=300).original
                img = Image.open(io.BytesIO(image))
                ocr_text = pytesseract.image_to_string(img)
                full_text += ocr_text + "\n"
        return full_text

# Function to clean and extract name, email, and phone from the extracted text
def extract_info_from_text(extracted_text):
    info = {
        "Name": None,
        "Email": None,
        "Phone": None
    }

    EMAIL_REGEX = r'[\w\.-]+@[\w\.-]+\.\w+'
    PHONE_REGEX = r'(\+?\d{1,3}[-/]?\d{1,4}[-/]?\d{7,9})'  # Ensures at least 11 digits in total

    # Extract email
    email_match = re.search(EMAIL_REGEX, extracted_text)
    if email_match:
        info["Email"] = email_match.group(0)

    # Extract phone
    phone_matches = re.findall(PHONE_REGEX, extracted_text)
    if phone_matches:
        # Clean phone number (remove spaces and special characters)
        clean_phone = re.sub(r'[^+\d]', '', phone_matches[0])  # Remove all except digits and the plus sign
        info["Phone"] = clean_phone

    # Clean up and extract name (removing excessive spaces between letters)
    lines = extracted_text.splitlines()
    if len(lines) > 0:
        raw_name = lines[0].strip()  # First line is assumed to be the name
        cleaned_name = re.sub(r'\s+', ' ', raw_name)  # Remove excessive spacing between letters
        info["Name"] = cleaned_name.strip()

    return info

# Function to process a single PDF
def process_pdf_file(uploaded_file):
    extracted_text = extract_text_from_pdf(uploaded_file)
    info = extract_info_from_text(extracted_text)
    return info

# Streamlit interface
st.title("CV Parser App")

# File uploader
uploaded_files = st.file_uploader("Choose PDF files", type="pdf", accept_multiple_files=True)

# Button to start the parsing process
if st.button("Parse Resumes"):
    if uploaded_files:
        data = []
        for uploaded_file in uploaded_files:
            st.write(f"Processing file: {uploaded_file.name}")
            info = process_pdf_file(uploaded_file)
            info['File Name'] = uploaded_file.name
            data.append(info)

        # Convert data to a pandas DataFrame
        df = pd.DataFrame(data)

        # Create an in-memory buffer to store the Excel file
        buffer = io.BytesIO()

        # Save the DataFrame to the buffer using the openpyxl engine
        df.to_excel(buffer, engine='openpyxl', index=False)

        # Move the buffer's pointer to the beginning
        buffer.seek(0)

        # Provide a download button for users to download the Excel file
        st.download_button(
            label="Download Excel file",
            data=buffer,
            file_name='parsed_resumes.xlsx',
            mime='application/vnd.ms-excel'
        )
    else:
        st.error("Please upload at least one PDF file.")

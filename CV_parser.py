import requests
import streamlit as st
import re
import pdfplumber
from PIL import Image
from docx import Document
import io
import pandas as pd

# OCR.space API key (replace with your key)
OCR_API_KEY = "K81881349288957"


# Function to sanitize filenames by removing special characters
def sanitize_filename(filename):
    # Replace special characters with underscores
    return re.sub(r'[^\w\s-]', '', filename).strip().replace(' ', '_')


# Function to extract text from images using OCR.space API
def extract_text_from_image_api(image):
    api_url = "https://api.ocr.space/parse/image"

    # Convert the image to bytes
    image_bytes = io.BytesIO()
    image.save(image_bytes, format='PNG')
    image_bytes = image_bytes.getvalue()

    # Send the image to OCR API
    payload = {
        'apikey': OCR_API_KEY,
        'language': 'ara,eng',  # Support both Arabic and English
    }
    files = {
        'file': ('image.png', image_bytes, 'image/png'),
    }

    response = requests.post(api_url, files=files, data=payload)

    # Check for errors in the API response
    if response.status_code != 200:
        st.error(f"Error: OCR API request failed with status {response.status_code}")
        return ""

    result = response.json()
    if "ParsedResults" not in result or not result["ParsedResults"]:
        st.error("No ParsedResults found in the OCR API response.")
        return ""

    # Extract the text from the API response
    extracted_text = result["ParsedResults"][0].get("ParsedText", "")
    return extracted_text


# Function to extract text from a PDF, handling both text and images
def extract_text_from_pdf(pdf_file):
    full_text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # Extract text from the page
            text = page.extract_text()
            if text:
                full_text += text + "\n"
            else:
                # If no text is found, attempt OCR on any images on the page
                for image in page.images:
                    img = page.to_image()
                    ocr_text = extract_text_from_image_api(img.original)
                    full_text += ocr_text + "\n"
    return full_text


# Function to extract text from a Word document, including OCR on images
def extract_text_from_word(docx_file):
    full_text = ""
    doc = Document(docx_file)

    # Extract text from paragraphs
    for para in doc.paragraphs:
        full_text += para.text + "\n"

    # Extract images and perform OCR
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            img_part = rel.target_part
            img = Image.open(io.BytesIO(img_part.blob))
            ocr_text = extract_text_from_image_api(img)
            full_text += ocr_text + "\n"

    return full_text


# Function to extract text from an image file directly
def extract_text_from_image_file(image_file):
    img = Image.open(image_file)
    ocr_text = extract_text_from_image_api(img)
    return ocr_text


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


# Function to process files (PDF, Word, or Image)
def process_file(uploaded_file, sanitized_filename):
    extracted_text = ""

    # Check file type and extract text accordingly
    if sanitized_filename.endswith('.pdf'):
        extracted_text = extract_text_from_pdf(uploaded_file)
    elif sanitized_filename.endswith('.docx'):
        extracted_text = extract_text_from_word(uploaded_file)
    elif sanitized_filename.endswith(('.png', '.jpg', '.jpeg')):
        extracted_text = extract_text_from_image_file(uploaded_file)

    # Extract relevant information (Name, Email, Phone)
    info = extract_info_from_text(extracted_text)

    # Debugging: Check if all fields are empty
    if not info['Name'] and not info['Email'] and not info['Phone']:
        st.warning(f"Failed to extract data from: {sanitized_filename}")

    return info


# Streamlit interface
st.title("CV Parser App")

# File uploader (accepting multiple PDFs, Word, and image files)
uploaded_files = st.file_uploader("Choose PDF, Word, or Image files", type=["pdf", "docx", "png", "jpg", "jpeg"],
                                  accept_multiple_files=True)

# Button to start the parsing process
if st.button("Upload Resumes"):
    if uploaded_files:
        data = []
        for uploaded_file in uploaded_files:
            # Sanitize the filename
            sanitized_filename = sanitize_filename(uploaded_file.name)

            # Log the sanitized filename for debugging
            st.write(f"Processing file: {sanitized_filename}")

            # Process the file (extracting text and info)
            try:
                info = process_file(uploaded_file, sanitized_filename)
                info['File Name'] = sanitized_filename
                data.append(info)
            except Exception as e:
                # Log the full error message for debugging
                st.error(f"Error processing file '{sanitized_filename}': {str(e)}")
                continue

        # If we have valid data, display and allow download
        if data:
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
            st.error("No valid resumes were processed.")
    else:
        st.error("Please upload at least one file.")

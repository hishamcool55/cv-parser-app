import requests
import streamlit as st
import re
import pdfplumber
from PIL import Image
import io
import pandas as pd


# Function to sanitize filenames by removing special characters
def sanitize_filename(filename):
    # Replace special characters with underscores
    return re.sub(r'[^\w\s-]', '', filename).strip().replace(' ', '_')


# Function to extract text from images using OCR.space API
def extract_text_from_image_api(image, filename, page_number):
    # Convert the image to bytes and send to OCR.space API
    api_url = "https://api.ocr.space/parse/image"
    api_key = "K81881349288957"  # Your OCR.space API key

    # Convert the image to bytes
    image_bytes = io.BytesIO()
    image.save(image_bytes, format='PNG')
    image_bytes = image_bytes.getvalue()

    # Send the image to OCR API
    payload = {
        'apikey': api_key,
        'language': 'ara,eng',  # Support for Arabic and English
    }
    files = {
        'file': ('image.png', image_bytes, 'image/png'),
    }

    response = requests.post(api_url, files=files, data=payload)

    # Check for errors in the API response
    if response.status_code != 200:
        st.error(f"Error: OCR API request failed with status {response.status_code}")
        return ""

    # Parse and return the OCR text from the API response
    result = response.json()
    extracted_text = result.get("ParsedResults")[0].get("ParsedText", "")

    # Print a message to indicate that OCR was used
    if extracted_text.strip():
        st.write(f"OCR was used on page {page_number} in file '{filename}'.")
    else:
        st.warning(f"OCR on page {page_number} in file '{filename}' found no readable text.")

    return extracted_text


# Function to extract text from a PDF, using OCR on image-only pages
def extract_text_from_pdf(pdf_path, filename):
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            # Extract text from the page
            text = page.extract_text()
            if text:
                # If there is text, extract and append it
                full_text += text + "\n"
            else:
                # If no text is found, attempt OCR on the image
                st.warning(f"Page {page.page_number} in file '{filename}' has no text. Attempting OCR on image.")

                # Get the image from the PDF page
                image = page.to_image().original

                # Perform OCR using the API (Arabic + English)
                ocr_text = extract_text_from_image_api(image, filename, page.page_number)

                # If OCR found any text, add it to the full text
                if ocr_text.strip():
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


# Function to process a single PDF file
def process_pdf_file(uploaded_file, sanitized_filename):
    extracted_text = extract_text_from_pdf(uploaded_file, sanitized_filename)
    info = extract_info_from_text(extracted_text)
    return info


# Streamlit interface
st.title("CV Parser App")

# File uploader (accepting multiple PDF files)
uploaded_files = st.file_uploader("Choose PDF files", type="pdf", accept_multiple_files=True)

# Button to start the parsing process
if st.button("Upload Resumes"):
    if uploaded_files:
        data = []
        for uploaded_file in uploaded_files:
            # Sanitize the filename
            sanitized_filename = sanitize_filename(uploaded_file.name)

            # Log the sanitized filename for debugging
            st.write(f"Processing file: {sanitized_filename}")

            # Check file size
            file_size = uploaded_file.size
            st.write(f"File size: {file_size} bytes")

            # If the file size exceeds a limit, skip processing (Example: 10MB)
            if file_size > 10 * 1024 * 1024:  # 10MB limit
                st.error(f"File '{sanitized_filename}' is too large. Maximum allowed size is 10MB.")
                continue

            # Process the file (extracting text and info)
            try:
                info = process_pdf_file(uploaded_file, sanitized_filename)
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
        st.error("Please upload at least one PDF file.")

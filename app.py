import streamlit as st
import win32com.client
import pythoncom
import os
import tempfile
from pdf2docx import Converter
from PIL import Image
import io
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF for PDF to JPG conversion
from PyPDF2 import PageObject

# Set up background and responsiveness using CSS
page_bg = '''
<style>
/* Digital Blue Background */
body {
    background-color: #1a2a6c;
    background-image: linear-gradient(120deg, #1a2a6c, #b21f1f, #fdbb2d);
    color: white;
    font-family: 'Arial', sans-serif;
}

/* Responsive heading styling */
h1 {
    color: #black;
    text-align: center;
    font-size: 2.5em;
    margin-bottom: 10px;
    text-shadow: 1px 1px #000000;
}

/* Button styling */
button {
    background-color: #1e3c72 !important;
    border: none !important;
    color: white !important;
    padding: 10px 20px !important;
    text-align: center !important;
    font-size: 16px !important;
    margin: 4px 2px !important;
    cursor: pointer !important;
    border-radius: 12px !important;
    transition: background-color 0.3s ease;
}

button:hover {
    background-color: #fdbb2d !important;
    color: #1a2a6c !important;
}

input[type="file"]:hover {
    background-color: #fdbb2d !important;
}

/* Responsive design */
@media (max-width: 768px) {
    h1 {
        font-size: 1.8em;
    }
}

@media (min-width: 769px) and (max-width: 1024px) {
    h1 {
        font-size: 2.2em;
    }
}

@media (min-width: 1025px) {
    h1 {
        font-size: 2.5em;
    }
}
</style>
'''

# Apply the CSS background
st.markdown(page_bg, unsafe_allow_html=True)

# Title and description for the web app
st.title("All-in-One File Converter & Compressor")
st.write("Convert and compress PDFs, Word docs, and images (JPG/PNG).")

# File conversion and compression functions
def pdf_to_word(pdf_file):
    word_file = pdf_file.name.replace(".pdf", ".docx")
    with open(pdf_file.name, "wb") as f:
        f.write(pdf_file.getbuffer())
    
    cv = Converter(pdf_file.name)
    cv.convert(word_file)
    cv.close()
    
    return word_file

def word_to_pdf(docx_file):
    temp_input_path = os.path.join(tempfile.gettempdir(), "temp_input.docx")
    output_pdf = os.path.join(tempfile.gettempdir(), "output.pdf")
    
    with open(temp_input_path, "wb") as f:
        f.write(docx_file.getbuffer())
    
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    
    try:
        doc = word.Documents.Open(temp_input_path)
        doc.SaveAs(output_pdf, FileFormat=17)
    except Exception as e:
        st.error(f"An error occurred: {e}")
    finally:
        if doc:
            doc.Close(False)
        if word:
            word.Quit()
        pythoncom.CoUninitialize()
    
    return output_pdf

def compress_image(image_file, quality, size_limit_kb):
    image = Image.open(image_file)
    img_format = image_file.type.split('/')[1]
    
    # Use a temporary file to save the compressed image
    temp_image = io.BytesIO()
    image.save(temp_image, format=img_format.upper(), quality=quality)
    temp_image.seek(0)
    
    # Check if the image size is within the limit
    if len(temp_image.getbuffer()) <= size_limit_kb * 1024:
        return temp_image
    
    # If not, progressively lower the quality until it meets the limit
    quality_step = 5
    while quality > 0:
        temp_image = io.BytesIO()
        image.save(temp_image, format=img_format.upper(), quality=quality)
        temp_image.seek(0)
        
        if len(temp_image.getbuffer()) <= size_limit_kb * 1024:
            return temp_image
        
        quality -= quality_step

    st.error(f"Unable to compress image to meet the size limit of {size_limit_kb} KB.")
    return None

def compress_pdf(pdf_file, size_limit_kb):
    temp_input_path = os.path.join(tempfile.gettempdir(), "temp_input.pdf")
    temp_output_path = os.path.join(tempfile.gettempdir(), "compressed_output.pdf")
    
    with open(temp_input_path, "wb") as f:
        f.write(pdf_file.getbuffer())
    
    # Attempt to compress PDF by iterating and adjusting quality
    # This is a simplified approach; for exact size control, more advanced methods might be needed
    quality_step = 10
    while True:
        pdf_writer = PdfWriter()
        pdf_reader = PdfReader(temp_input_path)
        
        for page_num in range(len(pdf_reader.pages)):
            pdf_writer.add_page(pdf_reader.pages[page_num])
        
        with open(temp_output_path, "wb") as f:
            pdf_writer.write(f)
        
        temp_output_file_size_kb = os.path.getsize(temp_output_path) / 1024
        
        if temp_output_file_size_kb <= size_limit_kb:
            break
        
        # Reduce the quality if necessary (this is a placeholder)
        # For more accurate control, consider using external libraries or tools
        
        if quality_step <= 0:
            st.error(f"Unable to compress PDF to meet the size limit of {size_limit_kb} KB.")
            return None
        
        quality_step -= 1

    with open(temp_output_path, "rb") as f:
        return io.BytesIO(f.read())

def jpg_to_pdf(image_file):
    image = Image.open(image_file)
    pdf_file = io.BytesIO()
    image.save(pdf_file, "PDF", resolution=100.0)
    pdf_file.seek(0)
    return pdf_file

def pdf_to_jpg(pdf_file):
    pdf_document = fitz.open(stream=pdf_file.read(), filetype="pdf")
    img_bytes_list = []
    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img_bytes = io.BytesIO()
        img.save(img_bytes, format="JPEG")
        img_bytes.seek(0)
        img_bytes_list.append(img_bytes)
    return img_bytes_list

def jpg_to_word(image_file):
    # This functionality requires OCR and is complex; placeholder implementation
    st.error("JPG to Word conversion is not implemented.")
    return None

# Sidebar options
st.sidebar.title("Choose the Action")
option = st.sidebar.selectbox(
    "What do you want to do?", 
    [
        "PDF to Word", 
        "Word to PDF", 
        "Compress Image", 
        "Compress PDF",
        "JPG to PDF",
        "PDF to JPG",
    ]
)

# PDF to Word Conversion
if option == "PDF to Word":
    st.header("Convert PDF to Word")
    pdf_file = st.file_uploader("Upload a PDF file", type="pdf")
    
    if pdf_file is not None:
        output_word_file = pdf_to_word(pdf_file)
        st.success(f"PDF file successfully converted to Word.")
        with open(output_word_file, "rb") as f:
            st.download_button(label="Download Word File", data=f.read(), file_name=output_word_file, mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

# Word to PDF Conversion
elif option == "Word to PDF":
    st.header("Convert Word to PDF")
    docx_file = st.file_uploader("Upload a Word file", type="docx")
    
    if docx_file is not None:
        output_pdf_file = word_to_pdf(docx_file)
        st.success(f"Word file successfully converted to PDF.")
        with open(output_pdf_file, "rb") as f:
            st.download_button(label="Download PDF File", data=f.read(), file_name=output_pdf_file, mime='application/pdf')

# Compress Image
elif option == "Compress Image":
    st.header("Compress an Image")
    image_file = st.file_uploader("Upload an image (JPEG, PNG)", type=["jpg", "jpeg", "png"])
    quality = st.slider("Select image quality (lower means more compression)", min_value=1, max_value=100, value=85)
    size_limit = st.selectbox("Select size limit", ["KB", "MB"])
    size_limit_value = st.number_input("Set size limit value", min_value=1, max_value=10000, value=500)
    
    if size_limit == "MB":
        size_limit_kb = size_limit_value * 1024
    else:
        size_limit_kb = size_limit_value
    
    if image_file is not None:
        compressed_image = compress_image(image_file, quality, size_limit_kb)
        if compressed_image:
            st.success("Image successfully compressed.")
            st.download_button(label="Download Compressed Image", data=compressed_image, file_name="compressed_image.jpg", mime="image/jpeg")

# Compress PDF
elif option == "Compress PDF":
    st.header("Compress a PDF")
    pdf_file = st.file_uploader("Upload a PDF file", type="pdf")
    size_limit = st.selectbox("Select size limit", ["KB", "MB"])
    size_limit_value = st.number_input("Set size limit value", min_value=1, max_value=10000, value=500)
    
    if size_limit == "MB":
        size_limit_kb = size_limit_value * 1024
    else:
        size_limit_kb = size_limit_value
    
    if pdf_file is not None:
        compressed_pdf = compress_pdf(pdf_file, size_limit_kb)
        if compressed_pdf:
            st.success("PDF successfully compressed.")
            st.download_button(label="Download Compressed PDF", data=compressed_pdf, file_name="compressed_pdf.pdf", mime="application/pdf")

# JPG to PDF Conversion
elif option == "JPG to PDF":
    st.header("Convert JPG to PDF")
    image_file = st.file_uploader("Upload an image (JPEG, PNG)", type=["jpg", "jpeg", "png"])
    
    if image_file is not None:
        pdf_file = jpg_to_pdf(image_file)
        st.success("Image successfully converted to PDF.")
        st.download_button(label="Download PDF File", data=pdf_file, file_name="image_to_pdf.pdf", mime="application/pdf")

# PDF to JPG Conversion
elif option == "PDF to JPG":
    st.header("Convert PDF to JPG")
    pdf_file = st.file_uploader("Upload a PDF file", type="pdf")
    
    if pdf_file is not None:
        images = pdf_to_jpg(pdf_file)
        for i, img_bytes in enumerate(images):
            st.image(img_bytes, caption=f"Page {i+1}")
            st.download_button(label=f"Download JPG for Page {i+1}", data=img_bytes, file_name=f"page_{i+1}.jpg", mime="image/jpeg")

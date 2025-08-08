import streamlit as st
import fitz  # PyMuPDF
from PIL import Image
import io
import os
import tempfile
from pdf2docx import Converter
from docx2pdf import convert
import pandas as pd
import img2pdf
from pdf2image import convert_from_path
import base64
from fpdf import FPDF
import csv

# Configuration de la page
st.set_page_config(page_title="PDF Master Tool", page_icon="📄", layout="wide")

# Sidebar avec logo et description
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/337/337946.png", width=100)
    st.title("PDF Master Tool")
    st.markdown("""
    **L'application complète pour manipuler vos fichiers PDF :**
    - Conversion entre formats
    - Compression et fusion
    - Protection des documents
    """)

# Fonctions principales
def compress_pdf(input_pdf, output_pdf, zoom_x=0.5, zoom_y=0.5, quality=50):
    pdf_document = fitz.open(input_pdf)
    new_pdf = fitz.open()
    
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        mat = fitz.Matrix(zoom_x, zoom_y)
        pix = page.get_pixmap(matrix=mat)
        img = Image.open(io.BytesIO(pix.tobytes()))
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format="JPEG", quality=quality)
        img_byte_arr = img_byte_arr.getvalue()
        new_page = new_pdf.new_page(width=pix.width, height=pix.height)
        new_page.insert_image(new_page.rect, stream=img_byte_arr)
    
    new_pdf.save(output_pdf)
    new_pdf.close()
    pdf_document.close()

def merge_pdfs(pdf_files, output_pdf):
    merged_pdf = fitz.open()
    
    for pdf_file in pdf_files:
        pdf_document = fitz.open(pdf_file)
        merged_pdf.insert_pdf(pdf_document)
        pdf_document.close()
    
    merged_pdf.save(output_pdf)
    merged_pdf.close()

def split_pdf(input_pdf, pages, output_pdf):
    pdf_document = fitz.open(input_pdf)
    new_pdf = fitz.open()
    
    for page_num in pages:
        new_pdf.insert_pdf(pdf_document, from_page=page_num-1, to_page=page_num-1)
    
    new_pdf.save(output_pdf)
    new_pdf.close()
    pdf_document.close()

def pdf_to_word(input_pdf, output_docx):
    cv = Converter(input_pdf)
    cv.convert(output_docx, start=0, end=None)
    cv.close()

def word_to_pdf(input_docx, output_pdf):
    convert(input_docx, output_pdf)

def pdf_to_excel(input_pdf, output_xlsx):
    pdf = fitz.open(input_pdf)
    text = ""
    for page in pdf:
        text += page.get_text()
    
    # Création d'un DataFrame simple
    data = []
    for line in text.split('\n'):
        if line.strip():
            row = [cell.strip() for cell in line.split('\t') if cell.strip()]
            if row:
                data.append(row)
    
    # Trouver le nombre maximum de colonnes
    max_cols = max(len(row) for row in data) if data else 0
    
    # Remplir les lignes avec moins de colonnes
    for row in data:
        while len(row) < max_cols:
            row.append('')
    
    df = pd.DataFrame(data)
    df.to_excel(output_xlsx, index=False, header=False)
    pdf.close()

def excel_to_pdf(input_xlsx, output_pdf):
    try:
        # Lire le fichier Excel
        df = pd.read_excel(input_xlsx)
        
        # Créer un PDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=8)
        
        # Ajouter les en-têtes
        for col in df.columns:
            pdf.cell(40, 10, str(col), border=1)
        pdf.ln()
        
        # Ajouter les données
        for _, row in df.iterrows():
            for item in row:
                # Gestion des caractères spéciaux
                try:
                    text = str(item).encode('latin1', 'replace').decode('latin1')
                except:
                    text = str(item)
                pdf.cell(40, 10, text, border=1)
            pdf.ln()
        
        pdf.output(output_pdf)
        return True
    except Exception as e:
        st.error(f"Erreur lors de la conversion : {str(e)}")
        return False

def jpg_to_pdf(input_images, output_pdf):
    with open(output_pdf, "wb") as f:
        f.write(img2pdf.convert(input_images))

def protect_pdf(input_pdf, output_pdf, password):
    pdf = fitz.open(input_pdf)
    pdf.save(output_pdf, encryption=fitz.PDF_ENCRYPT_AES_256, user_pw=password)
    pdf.close()

# Interface principale
st.title("📄 PDF Master Tool")

# Menu des fonctionnalités dans la sidebar
option = st.sidebar.selectbox(
    "Choisissez une fonctionnalité",
    [
        "Compresser PDF", "Fusionner PDF", "Diviser PDF",
        "PDF vers Word", "Word vers PDF",
        "PDF vers Excel", "Excel vers PDF",
        "JPG vers PDF", "Protéger PDF"
    ]
)

# Fonctionnalités
if option == "Compresser PDF":
    st.header("Compresser un PDF")
    uploaded_file = st.file_uploader("Importer un fichier PDF", type="pdf")
    
    if uploaded_file:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name
        
        col1, col2 = st.columns(2)
        with col1:
            quality = st.slider("Qualité de compression", 0, 100, 50)
        with col2:
            zoom = st.slider("Facteur de zoom", 0.1, 1.0, 0.5)
        
        if st.button("Compresser"):
            output_path = "compressed.pdf"
            compress_pdf(tmp_path, output_path, zoom, zoom, quality)
            
            with open(output_path, "rb") as f:
                st.download_button("Télécharger PDF compressé", f, "compressed.pdf")
            
            os.unlink(tmp_path)
            os.unlink(output_path)

elif option == "Fusionner PDF":
    st.header("Fusionner des PDF")
    uploaded_files = st.file_uploader("Importer plusieurs PDF", type="pdf", accept_multiple_files=True)
    
    if uploaded_files and len(uploaded_files) > 1:
        temp_files = []
        for i, file in enumerate(uploaded_files):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(file.getvalue())
                temp_files.append(tmp.name)
        
        if st.button("Fusionner"):
            output_path = "merged.pdf"
            merge_pdfs(temp_files, output_path)
            
            with open(output_path, "rb") as f:
                st.download_button("Télécharger PDF fusionné", f, "merged.pdf")
            
            for f in temp_files:
                os.unlink(f)
            os.unlink(output_path)

elif option == "Diviser PDF":
    st.header("Diviser un PDF")
    uploaded_file = st.file_uploader("Importer un PDF à diviser", type="pdf")
    
    if uploaded_file:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name
        
        pdf = fitz.open(tmp_path)
        pages = st.multiselect("Sélectionnez les pages à extraire", range(1, len(pdf)+1))
        pdf.close()
        
        if pages and st.button("Extraire les pages"):
            output_path = "extracted.pdf"
            split_pdf(tmp_path, pages, output_path)
            
            with open(output_path, "rb") as f:
                st.download_button("Télécharger PDF extrait", f, "extracted.pdf")
            
            os.unlink(tmp_path)
            os.unlink(output_path)

elif option == "PDF vers Word":
    st.header("Convertir PDF en Word")
    uploaded_file = st.file_uploader("Importer un PDF", type="pdf")
    
    if uploaded_file:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(uploaded_file.getvalue())
            pdf_path = tmp_pdf.name
        
        if st.button("Convertir en Word"):
            docx_path = "converted.docx"
            pdf_to_word(pdf_path, docx_path)
            
            with open(docx_path, "rb") as f:
                st.download_button("Télécharger Word", f, "converted.docx")
            
            os.unlink(pdf_path)
            os.unlink(docx_path)

elif option == "Word vers PDF":
    st.header("Convertir Word en PDF")
    uploaded_file = st.file_uploader("Importer un document Word", type=["docx"])
    
    if uploaded_file:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
            tmp_docx.write(uploaded_file.getvalue())
            docx_path = tmp_docx.name
        
        if st.button("Convertir en PDF"):
            pdf_path = "converted.pdf"
            word_to_pdf(docx_path, pdf_path)
            
            with open(pdf_path, "rb") as f:
                st.download_button("Télécharger PDF", f, "converted.pdf")
            
            os.unlink(docx_path)
            os.unlink(pdf_path)

elif option == "PDF vers Excel":
    st.header("Convertir PDF en Excel")
    uploaded_file = st.file_uploader("Importer un PDF", type="pdf")
    
    if uploaded_file:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(uploaded_file.getvalue())
            pdf_path = tmp_pdf.name
        
        if st.button("Convertir en Excel"):
            xlsx_path = "converted.xlsx"
            pdf_to_excel(pdf_path, xlsx_path)
            
            with open(xlsx_path, "rb") as f:
                st.download_button("Télécharger Excel", f, "converted.xlsx")
            
            os.unlink(pdf_path)
            os.unlink(xlsx_path)

elif option == "Excel vers PDF":
    st.header("Convertir Excel en PDF")
    uploaded_file = st.file_uploader("Importer un fichier Excel", type=["xlsx", "xls"])
    
    if uploaded_file:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
            tmp_xlsx.write(uploaded_file.getvalue())
            xlsx_path = tmp_xlsx.name
        
        if st.button("Convertir en PDF"):
            pdf_path = "converted.pdf"
            excel_to_pdf(xlsx_path, pdf_path)
            
            with open(pdf_path, "rb") as f:
                st.download_button("Télécharger PDF", f, "converted.pdf")
            
            os.unlink(xlsx_path)
            os.unlink(pdf_path)

elif option == "JPG vers PDF":
    st.header("Convertir JPG en PDF")
    uploaded_files = st.file_uploader("Importer des images JPG", type=["jpg", "jpeg"], accept_multiple_files=True)
    
    if uploaded_files:
        temp_files = []
        for i, file in enumerate(uploaded_files):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                tmp.write(file.getvalue())
                temp_files.append(tmp.name)
        
        if st.button("Convertir en PDF"):
            pdf_path = "converted.pdf"
            jpg_to_pdf(temp_files, pdf_path)
            
            with open(pdf_path, "rb") as f:
                st.download_button("Télécharger PDF", f, "converted.pdf")
            
            for f in temp_files:
                os.unlink(f)
            os.unlink(pdf_path)

elif option == "Protéger PDF":
    st.header("Protéger un PDF avec mot de passe")
    uploaded_file = st.file_uploader("Importer un PDF", type="pdf")
    password = st.text_input("Mot de passe", type="password")
    
    if uploaded_file and password:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name
        
        if st.button("Protéger le PDF"):
            output_path = "protected.pdf"
            protect_pdf(tmp_path, output_path, password)
            
            with open(output_path, "rb") as f:
                st.download_button("Télécharger PDF protégé", f, "protected.pdf")
            
            os.unlink(tmp_path)
            os.unlink(output_path)

# Pied de page
st.markdown("---")
st.markdown("© 2023 PDF Master Tool - Tous droits réservés")
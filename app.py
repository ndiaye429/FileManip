import streamlit as st
import fitz  # PyMuPDF
from PIL import Image  # Pour la compression d'image
import io
import os

# Fonction pour compresser un PDF
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

# Fonction pour fusionner des PDF
def merge_pdfs(pdf_files, output_pdf):
    merged_pdf = fitz.open()
    
    for pdf_file in pdf_files:
        pdf_document = fitz.open(pdf_file)
        merged_pdf.insert_pdf(pdf_document)
        pdf_document.close()
    
    merged_pdf.save(output_pdf)
    merged_pdf.close()

# Interface Streamlit
st.title("Application de Manipulation PDF")

# Onglets pour séparer les fonctionnalités
tab1, tab2 = st.tabs(["Compresser un PDF", "Fusionner des PDF"])

# Onglet de compression
with tab1:
    st.header("Compresser un PDF")
    st.write("Téléchargez un fichier PDF pour le compresser.")
    
    uploaded_file = st.file_uploader("Importer un fichier PDF", type="pdf", key="compress")
    
    if uploaded_file is not None:
        # Afficher les informations du fichier
        file_size = len(uploaded_file.getvalue()) / 1024  # Taille en Ko
        st.write(f"Nom du fichier : {uploaded_file.name}")
        st.write(f"Taille du fichier : {file_size:.2f} Ko")
        
        # Ouvrir le fichier pour obtenir le nombre de pages
        pdf_document = fitz.open(stream=uploaded_file.getvalue(), filetype="pdf")
        st.write(f"Nombre de pages : {len(pdf_document)}")
        pdf_document.close()
        
        # Paramètres de compression
        st.subheader("Paramètres de compression")
        quality = st.slider("Qualité de compression (0-100)", 0, 100, 50, key="quality")
        zoom_x = st.slider("Facteur de zoom horizontal", 0.1, 1.0, 0.5, key="zoom_x")
        zoom_y = st.slider("Facteur de zoom vertical", 0.1, 1.0, 0.5, key="zoom_y")
        
        # Bouton pour valider les paramètres
        if st.button("Valider et compresser le PDF"):
            # Sauvegarder le fichier temporairement
            with open("temp.pdf", "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            # Compresser le PDF avec les paramètres choisis
            output_pdf = "compressed.pdf"
            compress_pdf("temp.pdf", output_pdf, zoom_x=zoom_x, zoom_y=zoom_y, quality=quality)
            
            # Afficher un message de succès
            st.success("Le fichier PDF a été compressé avec succès !")
            
            # Télécharger le fichier compressé
            with open(output_pdf, "rb") as f:
                st.download_button(
                    label="Télécharger le fichier compressé",
                    data=f,
                    file_name="compressed.pdf",
                    mime="application/pdf"
                )
            
            # Supprimer les fichiers temporaires
            os.remove("temp.pdf")
            os.remove(output_pdf)

# Onglet de fusion
with tab2:
    st.header("Fusionner des PDF")
    st.write("Téléchargez plusieurs fichiers PDF pour les fusionner en un seul.")
    
    uploaded_files = st.file_uploader("Importer des fichiers PDF", type="pdf", accept_multiple_files=True, key="merge")
    
    if uploaded_files:
        if len(uploaded_files) < 2:
            st.warning("Veuillez importer au moins deux fichiers PDF pour les fusionner.")
        else:
            # Afficher les fichiers importés
            st.subheader("Fichiers importés")
            for i, uploaded_file in enumerate(uploaded_files):
                st.write(f"{i + 1}. {uploaded_file.name} ({(len(uploaded_file.getvalue()) / 1024):.2f} Ko)")
            
            # Ajustement de l'ordre des fichiers
            st.subheader("Ajuster l'ordre des fichiers")
            file_names = [file.name for file in uploaded_files]
            selected_order = st.multiselect("Sélectionnez l'ordre des fichiers", file_names, default=file_names)
            
            # Bouton pour valider la fusion
            if st.button("Valider et fusionner les PDF"):
                # Sauvegarder les fichiers temporairement dans l'ordre sélectionné
                temp_files = []
                for file_name in selected_order:
                    for uploaded_file in uploaded_files:
                        if uploaded_file.name == file_name:
                            temp_file = f"temp_{file_name}.pdf"
                            with open(temp_file, "wb") as f:
                                f.write(uploaded_file.getbuffer())
                            temp_files.append(temp_file)
                            break
                
                # Fusionner les fichiers
                output_pdf = "merged.pdf"
                merge_pdfs(temp_files, output_pdf)
                
                st.success("Les fichiers PDF ont été fusionnés avec succès !")
                
                # Télécharger le fichier fusionné
                with open(output_pdf, "rb") as f:
                    st.download_button(
                        label="Télécharger le fichier fusionné",
                        data=f,
                        file_name="merged.pdf",
                        mime="application/pdf"
                    )
                
                # Supprimer les fichiers temporaires
                for temp_file in temp_files:
                    os.remove(temp_file)
                os.remove(output_pdf)
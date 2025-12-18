import streamlit as st
import pandas as pd
from pdfrw import PdfReader, PdfWriter, PageMerge
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import requests
from io import BytesIO
from zipfile import ZipFile
import os

# ===== CONFIG SITE =====
st.set_page_config(
    page_title="Générateur de fiches PDF",
    page_icon="logo.png",
    layout="centered"
)

# --- Style CSS personnalisé ---
st.markdown(
    """
    <style>
    /* Dégradé de fond */
    .stApp {
        background: linear-gradient(135deg, #3daeeb, #b508fe);
        background-attachment: fixed;
    }

    /* Logo principal avec ombre */
    .logo-container img {
        box-shadow: 0px 4px 15px rgba(0,0,0,0.3);
        border-radius: 12px;
        max-width: 200px;
    }

    /* Barre PDF étudiant */
    .pdf-bar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background-color: #a221fb;
        color: white;
        padding: 12px 20px;
        border-radius: 12px;
        margin-bottom: 10px;
        font-size: 18px;
        font-weight: bold;
    }

    .stButton>button {
        background-color: #4CAF50;
        color: white;
        font-size: 16px;
        border-radius: 8px;
        padding: 8px 15px;
    }

    /* Messages info et succès uniquement */
    .stInfo, .stSuccess {
        color: black !important;
        font-weight: normal;
        text-shadow: none;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# --- Logo principal et titre ---
st.markdown('<div class="logo-container" style="text-align:center">', unsafe_allow_html=True)
st.image("logo 1.png")
st.markdown('</div>', unsafe_allow_html=True)
st.title("Générateur de fiches PDF")

# ===== CONFIG PDF =====
EXCEL_URL = "https://drive.google.com/uc?export=download&id=1ky704tr63-DNl0CW7m__4DyoN1zPTP9u"
PDF_TEMPLATE = "Fiche de renseignements.pdf"

# --- Initialisation de la session ---
if "pdf_buffers" not in st.session_state:
    st.session_state.pdf_buffers = None

# --- Bouton de génération ---
if st.button("Générer les Fiches de renseignements"):
    try:
        st.info("Téléchargement du fichier Excel et génération en cours...")
        resp = requests.get(EXCEL_URL)
        resp.raise_for_status()
        df = pd.read_excel(BytesIO(resp.content))

        mapping = {
            "Champ de texte 110": "A",
            "Champ de texte 111": "F",
            "Champ de texte 112": "J",
            "Champ de texte 113": "N",
            "Champ de texte 114": "H",
            "Champ de texte 115": "L",
            "Champ de texte 117": "G",
            "Champ de texte 118": "K",
            "Champ de texte 120": "I",
            "Champ de texte 121": "M",
            "Champ de texte 122": "E",
            "Champ de texte 123": "B",
            "Champ de texte 124": "C",
            "Champ de texte 125": "D",
            "Champ de texte 126": "O",
            "Champ de texte 127": "P",
            "Champ de texte 128": "Q",
            "Champ de texte 129": "R",
            "Champ de texte 130": "S",
            "Champ de texte 131": "T",
            "Champ de texte 132": "U"
        }

        def idx(c):
            return ord(c) - 65

        pdf_buffers = []

        for _, row in df.iterrows():
            base_pdf = PdfReader(PDF_TEMPLATE)
            overlay_path = "__overlay.pdf"

            c = canvas.Canvas(overlay_path, pagesize=A4)
            c.setFont("Helvetica", 10)

            for page in base_pdf.pages:
                if page.Annots:
                    for a in page.Annots:
                        if a.T and a.Rect:
                            name = a.T.to_unicode()
                            if name in mapping:
                                value = row.iloc[idx(mapping[name])]
                                if pd.notna(value):
                                    x1, y1, x2, y2 = map(float, a.Rect)
                                    c.drawString(x1 + 4, y1 + 6, str(value))
                c.showPage()

            c.save()

            overlay_pdf = PdfReader(overlay_path)
            writer = PdfWriter()

            for i in range(len(base_pdf.pages)):
                page = base_pdf.pages[i]
                PageMerge(page).add(overlay_pdf.pages[i]).render()
                writer.addpage(page)

            output_name = f"{row.iloc[0]}_{row.iloc[1]}.pdf"
            buffer = BytesIO()
            writer.write(buffer)
            buffer.seek(0)
            pdf_buffers.append((output_name, buffer))

            os.remove(overlay_path)

        st.session_state.pdf_buffers = pdf_buffers
        st.success("✅ Terminé")

    except Exception as e:
        st.error(f"Erreur : {e}")

# --- Affichage des PDFs avec barre personnalisée ---
if st.session_state.pdf_buffers:
    st.subheader("Téléchargement des PDFs")
    for pdf_name, pdf_buffer in st.session_state.pdf_buffers:
        # Crée une ligne avec deux colonnes
        col1, col2 = st.columns([4,1])  # 4:1 pour donner plus de place au nom
        with col1:
            st.markdown(f'<div class="pdf-bar">{pdf_name}</div>', unsafe_allow_html=True)
        with col2:
            st.download_button(
                label="Télécharger",
                data=pdf_buffer,
                file_name=pdf_name,
                mime="application/pdf",
                key=pdf_name
            )


    # Bouton ZIP
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, "w") as zip_file:
        for pdf_name, pdf_buffer in st.session_state.pdf_buffers:
            pdf_buffer.seek(0)
            zip_file.writestr(pdf_name, pdf_buffer.read())

    st.download_button(
        label="Télécharger tous les PDFs en un seul ZIP",
        data=zip_buffer.getvalue(),
        file_name="PDFs_etudiants.zip",
        mime="application/zip"
    )

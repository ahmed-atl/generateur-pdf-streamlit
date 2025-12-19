import streamlit as st
import pandas as pd
from pdfrw import PdfReader, PdfWriter, PageMerge
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import requests
from io import BytesIO
from zipfile import ZipFile
import os
from PyPDF2 import PdfReader as PdfReader2, PdfWriter as PdfWriter2
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# =========================
# ===== CONFIG SITE =====
# =========================
st.set_page_config(
    page_title="Générateur de fiches PDF",
    page_icon="logo.png",
    layout="centered"
)

st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #3daeeb, #b508fe);
    background-attachment: fixed;
}

.logo-container img {
    box-shadow: 0px 4px 15px rgba(0,0,0,0.3);
    border-radius: 12px;
    max-width: 200px;
}

.pdf-bar {
    background-color: #a221fb;
    color: white;
    padding: 12px 20px;
    border-radius: 12px;
    margin-bottom: 10px;
    font-size: 16px;
    font-weight: bold;
}

.stButton>button {
    background-color: #4CAF50;
    color: white;
    font-size: 16px;
    border-radius: 8px;
    padding: 8px 15px;
}

.stInfo, .stSuccess {
    color: black !important;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="logo-container" style="text-align:center">', unsafe_allow_html=True)
st.image("logo 1.png")
st.markdown('</div>', unsafe_allow_html=True)
st.title("Générateur de fiches PDF")


# =========================
# ===== CONFIG FICHIERS =====
# =========================
EXCEL_URL = "https://drive.google.com/uc?export=download&id=1ky704tr63-DNl0CW7m__4DyoN1zPTP9u"
FICHE_PDF = "Fiche de renseignements.pdf"
REGLEMENT_URL = "https://drive.google.com/uc?export=download&id=1ZvYzecItMFm1vK2bdBEoBIMbCk54Zhse"

FONT_PATH = "Roboto-VariableFont_wdth,wght.ttf"
PAGE_REGLEMENT = 23  # page 24 indexée à 0

pdfmetrics.registerFont(TTFont("Roboto", FONT_PATH))


# =========================
# ===== SESSION =====
# =========================
if "fiches" not in st.session_state:
    st.session_state.fiches = None

if "reglements" not in st.session_state:
    st.session_state.reglements = None


# =========================
# ===== FICHES DE RENSEIGNEMENTS
# =========================
st.header("📄 Fiches de renseignements")

if st.button("Générer les fiches de renseignements"):
    try:
        st.info("Génération des fiches en cours...")
        df = pd.read_excel(BytesIO(requests.get(EXCEL_URL).content))

        mapping = {
            "Champ de texte 110": "A", "Champ de texte 111": "F",
            "Champ de texte 112": "J", "Champ de texte 113": "N",
            "Champ de texte 114": "H", "Champ de texte 115": "L",
            "Champ de texte 117": "G", "Champ de texte 118": "K",
            "Champ de texte 120": "I", "Champ de texte 121": "M",
            "Champ de texte 122": "E", "Champ de texte 123": "B",
            "Champ de texte 124": "C", "Champ de texte 125": "D",
            "Champ de texte 126": "O", "Champ de texte 127": "P",
            "Champ de texte 128": "Q", "Champ de texte 129": "R",
            "Champ de texte 130": "S", "Champ de texte 131": "T",
            "Champ de texte 132": "U"
        }

        def idx(c): return ord(c) - 65
        results = []

        for _, row in df.iterrows():
            base_pdf = PdfReader(FICHE_PDF)
            overlay = "__overlay.pdf"
            c = canvas.Canvas(overlay, pagesize=A4)
            c.setFont("Helvetica", 10)

            for page in base_pdf.pages:
                if page.Annots:
                    for a in page.Annots:
                        if a.T and a.Rect and a.T.to_unicode() in mapping:
                            val = row.iloc[idx(mapping[a.T.to_unicode()])]
                            if pd.notna(val):
                                x1, y1, x2, y2 = map(float, a.Rect)
                                c.drawString(x1 + 4, y1 + 6, str(val))
                c.showPage()

            c.save()
            overlay_pdf = PdfReader(overlay)
            writer = PdfWriter()
            for i in range(len(base_pdf.pages)):
                PageMerge(base_pdf.pages[i]).add(overlay_pdf.pages[i]).render()
                writer.addpage(base_pdf.pages[i])

            buf = BytesIO()
            writer.write(buf)
            buf.seek(0)
            filename = f"{row.iloc[0]}_{row.iloc[1]}.pdf"
            results.append((filename, buf))
            os.remove(overlay)

        st.session_state.fiches = results
        st.success("✅ Fiches générées")

    except Exception as e:
        st.error(e)


# =========================
# ===== AFFICHAGE FICHES
# =========================
if st.session_state.fiches:
    for name, buf in st.session_state.fiches:
        col1, col2 = st.columns([4,1])
        col1.markdown(f'<div class="pdf-bar">{name}</div>', unsafe_allow_html=True)
        col2.download_button("Télécharger", buf, name, "application/pdf", key=name)

    zip_buf = BytesIO()
    with ZipFile(zip_buf, "w") as z:
        for name, buf in st.session_state.fiches:
            buf.seek(0)
            z.writestr(name, buf.read())
    st.download_button("📦 ZIP fiches", zip_buf.getvalue(), "Fiches.zip", "application/zip")


# =========================
# ===== REGLEMENT INTERIEUR
# =========================
st.header("📘 Règlement intérieur")

if st.button("Générer les règlements intérieurs"):
    try:
        st.info("Génération des règlements...")

        df = pd.read_excel(BytesIO(requests.get(EXCEL_URL).content))
        reglement_pdf = BytesIO(requests.get(REGLEMENT_URL).content)

        results = []

        for _, row in df.iterrows():
            nom = f"{row.iloc[0]} {row.iloc[1]}"
            date = row.iloc[21]
            date = date.strftime("%d/%m/%Y") if pd.notna(date) else ""

            # création overlay
            packet = BytesIO()
            c = canvas.Canvas(packet, pagesize=A4)
            c.setFont("Roboto", 11)
            c.drawString(115, 296, nom)
            c.setFont("Roboto", 10)
            c.drawString(70, 195, date)
            c.save()
            packet.seek(0)

            overlay = PdfReader2(packet)
            base = PdfReader2(reglement_pdf)
            writer = PdfWriter2()

            for i, page in enumerate(base.pages):
                if i == PAGE_REGLEMENT:
                    page.merge_page(overlay.pages[0])
                writer.add_page(page)

            buf = BytesIO()
            writer.write(buf)
            buf.seek(0)
            results.append((f"Reglement_{nom}.pdf", buf))

        st.session_state.reglements = results
        st.success("✅ Règlements générés")

    except Exception as e:
        st.error(e)


# =========================
# ===== AFFICHAGE REGLEMENTS
# =========================
if st.session_state.reglements:
    for name, buf in st.session_state.reglements:
        col1, col2 = st.columns([4,1])
        col1.markdown(f'<div class="pdf-bar">{name}</div>', unsafe_allow_html=True)
        col2.download_button("Télécharger", buf, name, "application/pdf", key=f"r_{name}")

    zip_buf = BytesIO()
    with ZipFile(zip_buf, "w") as z:
        for name, buf in st.session_state.reglements:
            buf.seek(0)
            z.writestr(name, buf.read())
    st.download_button("📦 ZIP règlements", zip_buf.getvalue(), "Reglements.zip", "application/zip")

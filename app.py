import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import qrcode
from barcode import Code39, Code128
from barcode.writer import ImageWriter
import io
import zipfile
from docx import Document
from docx.shared import Mm
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

def mm_to_px(mm): return int(mm * 3.78)
def mm_to_pt(mm): return mm * 2.8346
def mm_to_excel_width(mm): return mm * 0.14

def generate_image(data, code_type, font_size, width_px, height_px):
    try:
        if code_type == "QR Code":
            qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M)
            qr.add_data(data)
            qr.make(fit=True)
            code_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
        elif code_type == "Code 128":
            barcode = Code128(data, writer=ImageWriter())
            output = io.BytesIO()
            barcode.write(output, {"write_text": False})
            code_img = Image.open(output).convert("RGB")
        else:
            barcode = Code39(data, writer=ImageWriter(), add_checksum=False)
            output = io.BytesIO()
            barcode.write(output, {"write_text": False})
            code_img = Image.open(output).convert("RGB")

        try:
            font = ImageFont.truetype("arial.ttf", font_size)
        except:
            font = ImageFont.load_default()

        draw_tmp = ImageDraw.Draw(Image.new("RGB", (1, 1)))
        bbox = draw_tmp.textbbox((0, 0), data, font=font)
        text_width, text_height = bbox[2] - bbox[0], bbox[3] - bbox[1]

        text_img = Image.new("RGB", (width_px, text_height + 10), "white")
        draw = ImageDraw.Draw(text_img)
        draw.text(((width_px - text_width) // 2, 5), data, fill="black", font=font)

        code_height = height_px - text_img.height
        code_img = code_img.resize((width_px, code_height))

        final_img = Image.new("RGB", (width_px, height_px), "white")
        final_img.paste(code_img, (0, 0))
        final_img.paste(text_img, (0, code_height))
        return final_img
    except Exception as e:
        st.error(f"Erreur de génération d'image : {e}")
        return Image.new("RGB", (width_px, height_px), "red")

def generate_excel(codes, code_type, font_size, label_w, label_h):
    wb = Workbook()
    ws = wb.active
    width_px = mm_to_px(label_w)
    height_px = mm_to_px(label_h)
    col_letter = get_column_letter(1)
    ws.column_dimensions[col_letter].width = mm_to_excel_width(label_w)

    for i, code in enumerate(codes, start=1):
        img = generate_image(code, code_type, font_size, width_px, height_px)
        tmp = io.BytesIO()
        img.save(tmp, format='PNG')
        tmp.seek(0)
        xl_img = XLImage(tmp)
        ws.row_dimensions[i].height = mm_to_pt(label_h)
        ws.add_image(xl_img, f"A{i}")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def generate_docx(codes, code_type, font_size, label_w, label_h):
    doc = Document()
    width_px = mm_to_px(label_w)
    height_px = mm_to_px(label_h)
    for code in codes:
        img = generate_image(code, code_type, font_size, width_px, height_px)
        tmp = io.BytesIO()
        img.save(tmp, format='PNG')
        tmp.seek(0)
        doc.add_picture(tmp, width=Mm(label_w), height=Mm(label_h))
        doc.add_paragraph("")
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

def generate_pdf(codes, code_type, font_size, label_w, label_h):
    output = io.BytesIO()
    c = canvas.Canvas(output, pagesize=A4)
    width_px = mm_to_px(label_w)
    height_px = mm_to_px(label_h)
    page_width, page_height = A4
    x, y = 40, page_height - 100

    for code in codes:
        img = generate_image(code, code_type, font_size, width_px, height_px)
        tmp = io.BytesIO()
        img.save(tmp, format='PNG')
        tmp.seek(0)
        c.drawImage(tmp, x, y, width=mm_to_pt(label_w), height=mm_to_pt(label_h))
        y -= mm_to_pt(label_h) + 10
        if y < 100:
            c.showPage()
            y = page_height - 100

    c.save()
    output.seek(0)
    return output

# Interface Streamlit
st.set_page_config(page_title="Générateur de codes-barres et QR codes", layout="centered")
st.image("Logo_Pelichet.svg", width=150)
st.markdown("<h1 style='text-align: center; color: #f26522; font-family: Trebuchet MS;'>Générateur de codes-barres et QR codes</h1>", unsafe_allow_html=True)
st.markdown("---")

with st.form("formulaire"):
    uploaded_file = st.file_uploader("📄 Fichier Excel", type=["xlsx"])
    format = st.selectbox("📂 Format de sortie", ["PDF", "Word", "Excel"])
    code_type = st.selectbox("🔢 Type de code", ["QR Code", "Code 39", "Code 128"])
    font_size = st.slider("✏️ Taille du texte", 6, 36, 12)
    label_w = st.number_input("📏 Largeur étiquette (mm)", value=50.0)
    label_h = st.number_input("📏 Hauteur étiquette (mm)", value=25.0)
    submitted = st.form_submit_button("Générer")

    if submitted:
        if uploaded_file is not None:
            df = pd.read_excel(uploaded_file)
            codes = df.iloc[:, 0].dropna().astype(str).tolist()
            if format == "Excel":
                file = generate_excel(codes, code_type, font_size, label_w, label_h)
                st.download_button("📥 Télécharger Excel", file, "etiquettes.xlsx")
            elif format == "Word":
                file = generate_docx(codes, code_type, font_size, label_w, label_h)
                st.download_button("📥 Télécharger Word", file, "etiquettes.docx")
            elif format == "PDF":
                file = generate_pdf(codes, code_type, font_size, label_w, label_h)
                st.download_button("📥 Télécharger PDF", file, "etiquettes.pdf")
        else:
            st.error("Veuillez déposer un fichier Excel.")

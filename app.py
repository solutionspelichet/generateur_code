
import streamlit as st
import traceback
import pandas as pd
import qrcode
from barcode import Code128, Code39
from barcode.writer import ImageWriter
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Mm
import io

def mm_to_px(mm): return int(mm * 3.78)

def generate_image(data, code_type, font_size, width_px, height_px):
    if code_type == "QR Code":
        qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M)
        qr.add_data(data)
        qr.make(fit=True)
        code_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    elif code_type == "Code 128":
        barcode = Code128(data, writer=ImageWriter())
        output = io.BytesIO()
        barcode.write(output, {"write_text": False})
        output.seek(0)
        code_img = Image.open(output).convert("RGB")
    else:
        barcode = Code39(data, writer=ImageWriter(), add_checksum=False)
        output = io.BytesIO()
        barcode.write(output, {"write_text": False})
        output.seek(0)
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

def main():
    st.set_page_config(page_title="Générateur Word", layout="centered")
    st.title("🧾 Générateur de codes (Word uniquement)")

    try:
        uploaded_file = st.file_uploader("📄 Fichier Excel", type=["xlsx"])
        code_type = st.selectbox("🔢 Type de code", ["QR Code", "Code 39", "Code 128"])
        font_size = st.slider("✏️ Taille du texte", 6, 36, 12)
        label_w = st.number_input("📏 Largeur étiquette (mm)", value=50.0)
        label_h = st.number_input("📏 Hauteur étiquette (mm)", value=25.0)
        submitted = st.button("Générer Word")

        if submitted and uploaded_file:
            df = pd.read_excel(uploaded_file)
            codes = df.iloc[:, 0].dropna().astype(str).tolist()
            width_px = mm_to_px(label_w)
            height_px = mm_to_px(label_h)

            doc = Document()
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
            st.download_button("📥 Télécharger Word", output, "etiquettes.docx")

        elif submitted:
            st.error("❌ Merci de fournir un fichier Excel valide.")

    except Exception as e:
        st.error("💥 Une erreur est survenue :")
        st.code(traceback.format_exc())

if __name__ == "__main__":
    main()

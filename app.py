
import streamlit as st
import traceback
import pandas as pd
import qrcode
from barcode import Code128, Code39
from barcode.writer import ImageWriter
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Mm
from docx.enum.section import WD_ORIENT
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
    st.set_page_config(page_title="Générateur Word complet", layout="centered")
    st.title("🧾 Générateur Word avec grille")

    try:
        uploaded_file = st.file_uploader("📄 Fichier Excel", type=["xlsx"])
        code_type = st.selectbox("🔢 Type de code", ["QR Code", "Code 39", "Code 128"])
        font_size = st.slider("✏️ Taille du texte", 6, 36, 12)
        label_w = st.number_input("📏 Largeur étiquette (mm)", value=50.0)
        label_h = st.number_input("📏 Hauteur étiquette (mm)", value=25.0)
        cols = st.number_input("🧱 Colonnes", min_value=1, max_value=10, value=2)
        rows = st.number_input("🧱 Lignes", min_value=1, max_value=30, value=6)
        spacing_x = st.number_input("↔️ Espacement horizontal (mm)", value=5.0)
        spacing_y = st.number_input("↕️ Espacement vertical (mm)", value=5.0)
        margin_top = st.number_input("⬆️ Marge haute (mm)", value=10.0)
        margin_right = st.number_input("➡️ Marge droite (mm)", value=10.0)
        submitted = st.button("Générer Word")

        if submitted and uploaded_file:
            df = pd.read_excel(uploaded_file)
            codes = df.iloc[:, 0].dropna().astype(str).tolist()
            width_px = mm_to_px(label_w)
            height_px = mm_to_px(label_h)

            doc = Document()
            section = doc.sections[-1]
            section.top_margin = Mm(margin_top)
            section.right_margin = Mm(margin_right)
            section.left_margin = Mm(margin_right)  # symmetry

            table = doc.add_table(rows=rows, cols=cols)
            table.autofit = False

            i = 0
            for r in range(rows):
                row_cells = table.rows[r].cells
                for c in range(cols):
                    if i < len(codes):
                        code = codes[i]
                        img = generate_image(code, code_type, font_size, width_px, height_px)
                        tmp = io.BytesIO()
                        img.save(tmp, format='PNG')
                        tmp.seek(0)
                        paragraph = row_cells[c].paragraphs[0]
                        run = paragraph.add_run()
                        run.add_picture(tmp, width=Mm(label_w), height=Mm(label_h))
                        i += 1

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


import streamlit as st
import traceback
import pandas as pd
import qrcode
from barcode import Code128, Code39
from barcode.writer import ImageWriter
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import io

def mm_to_px(mm): return int(mm * 3.78)
def mm_to_pt(mm): return mm * 2.8346

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
    st.set_page_config(page_title="Générateur PDF sécurisé", layout="centered")
    st.title("🧾 Générateur PDF (sécurisé)")

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
        submitted = st.button("Générer PDF")

        if submitted and uploaded_file:
            try:
                df = pd.read_excel(uploaded_file)
                codes = df.iloc[:, 0].dropna().astype(str).tolist()
                width_px = mm_to_px(label_w)
                height_px = mm_to_px(label_h)

                output = io.BytesIO()
                c = canvas.Canvas(output, pagesize=A4)
                page_width, page_height = A4

                label_w_pt = mm_to_pt(label_w)
                label_h_pt = mm_to_pt(label_h)
                spacing_x_pt = mm_to_pt(spacing_x)
                spacing_y_pt = mm_to_pt(spacing_y)
                margin_top_pt = mm_to_pt(margin_top)
                margin_right_pt = mm_to_pt(margin_right)
                margin_left_pt = margin_right_pt

                x0 = margin_left_pt
                y0 = page_height - margin_top_pt

                for i, code in enumerate(codes):
                    img = generate_image(code, code_type, font_size, width_px, height_px)
                    tmp = io.BytesIO()
                    img.save(tmp, format='PNG')
                    tmp.seek(0)

                    col = i % cols
                    row = (i // cols) % rows

                    if i > 0 and i % (cols * rows) == 0:
                        c.showPage()

                    x = x0 + col * (label_w_pt + spacing_x_pt)
                    y = y0 - row * (label_h_pt + spacing_y_pt)

                    c.drawImage(tmp, x, y - label_h_pt, width=label_w_pt, height=label_h_pt)

                c.save()
                output.seek(0)
                st.download_button("📥 Télécharger PDF", output, "etiquettes.pdf")

            except Exception as pdf_error:
                st.error("💥 Erreur lors de la génération PDF :")
                st.code(traceback.format_exc())

        elif submitted:
            st.error("❌ Merci de fournir un fichier Excel valide.")

    except Exception as e:
        st.error("💥 Une erreur générale est survenue :")
        st.code(traceback.format_exc())

if __name__ == "__main__":
    main()

from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
import os, io, tempfile
import pandas as pd
import qrcode
from barcode import Code39, Code128
from barcode.writer import ImageWriter
from PIL import Image, ImageDraw, ImageFont
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def mm_to_px(mm): return int(mm * 3.78)
def mm_to_pt(mm): return mm * 2.8346
def mm_to_excel_width(mm): return mm * 0.14

def generate_code_image_fit(data, code_type, font_size, width_px, height_px):
    try:
        # Génère le code seul (sans texte intégré)
        if code_type == "qr":
            qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M)
            qr.add_data(data)
            qr.make(fit=True)
            code_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
        elif code_type == "barcode128":
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

        # Prépare le texte
        temp_draw = ImageDraw.Draw(Image.new("RGB", (1, 1)))
        bbox = temp_draw.textbbox((0, 0), data, font=font)
        text_width, text_height = bbox[2] - bbox[0], bbox[3] - bbox[1]

        # Calcul des hauteurs
        available_height = height_px
        text_img = Image.new("RGB", (width_px, text_height + 10), "white")
        text_draw = ImageDraw.Draw(text_img)
        text_draw.text(((width_px - text_width) // 2, 5), data, fill="black", font=font)

        # Redimensionner le code pour qu'il tienne avec le texte
        code_height = available_height - text_img.height
        if code_height <= 0:
            code_height = 1
        code_img = code_img.resize((width_px, code_height))

        # Assemble image finale
        final_img = Image.new("RGB", (width_px, height_px), "white")
        final_img.paste(code_img, (0, 0))
        final_img.paste(text_img, (0, code_height))
        return final_img
    except Exception as e:
        print(f"[ERREUR CODE] {data} -> {e}")
        font = ImageFont.load_default()
        msg = f"Erreur: {data}"
        img = Image.new("RGB", (max(200, len(msg) * 6), 30), "white")
        draw = ImageDraw.Draw(img)
        draw.text((10, 10), msg, fill="red", font=font)
        return img

def detect_codes(excel_path):
    df = pd.read_excel(excel_path)
    for col in df.columns:
        if df[col].dropna().astype(str).str.len().gt(0).any():
            return df[col].dropna().astype(str).tolist()
    return []

def generate_excel(codes, output_path, code_type, font_size, label_w, label_h):
    wb = Workbook()
    ws = wb.active
    width_px = mm_to_px(label_w)
    height_px = mm_to_px(label_h)
    col_width = mm_to_excel_width(label_w)
    col_letter = get_column_letter(1)
    ws.column_dimensions[col_letter].width = col_width

    for i, code in enumerate(codes, start=1):
        img = generate_code_image_fit(code, code_type, font_size, width_px, height_px)
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        img.save(tmp.name)
        xl_img = XLImage(tmp.name)
        ws.row_dimensions[i].height = mm_to_pt(label_h)
        ws.add_image(xl_img, f"A{i}")
    wb.save(output_path)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['excel']
        output_format = request.form['format']
        code_type = request.form['code_type']
        font_size = int(request.form.get('font_size', 12))
        cols = int(request.form['cols'])
        rows = int(request.form['rows'])
        label_w = float(request.form['label_w'])
        label_h = float(request.form['label_h'])
        spacing_x = float(request.form['spacing_x'])
        spacing_y = float(request.form['spacing_y'])
        margin_top = float(request.form['margin_top'])
        margin_left = float(request.form['margin_left'])

        filename = secure_filename(file.filename)
        excel_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(excel_path)
        codes = detect_codes(excel_path)
        output_file = os.path.join(UPLOAD_FOLDER, f"output.{output_format}")

        if output_format == 'xlsx':
            generate_excel(codes, output_file, code_type, font_size, label_w, label_h)

        return send_file(output_file, as_attachment=True)

    return render_template('index.html')

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"Flask démarre sur le port {port}...")
    app.run(host="0.0.0.0", port=port)

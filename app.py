from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
import os, io, tempfile
import pandas as pd
import qrcode
from barcode import Code39, Code128
from barcode.writer import ImageWriter
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Mm
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def mm_to_px(mm): return int(mm * 3.78)
def mm_to_pt(mm): return mm * 2.8346
def mm_to_excel_width(mm): return mm * 0.14

def generate_code_image(data, code_type, font_size):
    if code_type == "qr":
        qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M)
        qr.add_data(data)
        qr.make(fit=True)
        base_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    elif code_type == "barcode128":
        barcode = Code128(data, writer=ImageWriter())
        output = io.BytesIO()
        barcode.write(output, {"write_text": False})
        output.seek(0)
        base_img = Image.open(output).convert("RGB")
    else:
        barcode = Code39(data, writer=ImageWriter(), add_checksum=False)
        output = io.BytesIO()
        barcode.write(output, {"write_text": False})
        output.seek(0)
        base_img = Image.open(output).convert("RGB")

    try:
        font = ImageFont.truetype("arial.ttf", font_size)
    except:
        font = ImageFont.load_default()

    draw_tmp = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    bbox = draw_tmp.textbbox((0, 0), data, font=font)
    text_width, text_height = bbox[2] - bbox[0], bbox[3] - bbox[1]

    img_with_text = Image.new("RGB", (base_img.width, base_img.height + text_height + 10), "white")
    img_with_text.paste(base_img, (0, 0))
    draw = ImageDraw.Draw(img_with_text)
    draw.text(((base_img.width - text_width) // 2, base_img.height + 5), data, fill="black", font=font)
    return img_with_text

def generate_code_image_stretched(data, code_type, font_size, width_px, height_px):
    img_with_text = generate_code_image(data, code_type, font_size)
    return img_with_text.resize((width_px, height_px))

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
        img = generate_code_image_stretched(code, code_type, font_size, width_px, height_px)
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        img.save(tmp.name)
        xl_img = XLImage(tmp.name)
        ws.row_dimensions[i].height = mm_to_pt(label_h)
        ws.add_image(xl_img, f"A{i}")
    wb.save(output_path)

def generate_pdf(codes, output_path, code_type, font_size, cols, rows, label_w, label_h, spacing_x, spacing_y, margin_top, margin_left):
    c = canvas.Canvas(output_path, pagesize=A4)
    page_width, page_height = A4
    label_w_pt = mm_to_pt(label_w)
    label_h_pt = mm_to_pt(label_h)
    spacing_x_pt = mm_to_pt(spacing_x)
    spacing_y_pt = mm_to_pt(spacing_y)
    margin_top_pt = mm_to_pt(margin_top)
    margin_left_pt = mm_to_pt(margin_left)

    index = 0
    for _ in range(9999):
        for r in range(rows):
            for col in range(cols):
                if index >= len(codes):
                    c.save()
                    return
                x = margin_left_pt + col * (label_w_pt + spacing_x_pt)
                y = page_height - margin_top_pt - (r + 1) * (label_h_pt + spacing_y_pt) + spacing_y_pt
                img = generate_code_image(codes[index], code_type, font_size)
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                img.save(tmp.name)
                c.drawImage(tmp.name, x, y, width=label_w_pt, height=label_h_pt)
                index += 1
        c.showPage()
    c.save()

def generate_word(codes, output_path, code_type, font_size, cols, rows, label_w, label_h, spacing_x, spacing_y, margin_top, margin_left):
    doc = Document()
    full_cols = cols * 2 - 1
    full_rows = rows * 2 - 1
    index = 0

    while index < len(codes):
        section = doc.sections[-1]
        section.top_margin = Mm(margin_top)
        section.left_margin = Mm(margin_left)
        section.right_margin = Mm(0)
        section.bottom_margin = Mm(0)
        table = doc.add_table(rows=full_rows, cols=full_cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        for r in range(full_rows):
            row = table.rows[r]
            row.height = Mm(label_h if r % 2 == 0 else spacing_y)
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            for c_idx in range(full_cols):
                cell = table.cell(r, c_idx)
                if r % 2 == 0 and c_idx % 2 == 0 and index < len(codes):
                    img = generate_code_image(codes[index], code_type, font_size)
                    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                    img.save(tmp.name)
                    run = cell.paragraphs[0].add_run()
                    run.add_picture(tmp.name, width=Mm(label_w), height=Mm(label_h))
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    cell.paragraphs[0].alignment = 1
                    index += 1
        if index < len(codes):
            doc.add_page_break()
    doc.save(output_path)

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
        elif output_format == 'pdf':
            generate_pdf(codes, output_file, code_type, font_size, cols, rows, label_w, label_h, spacing_x, spacing_y, margin_top, margin_left)
        elif output_format == 'docx':
            generate_word(codes, output_file, code_type, font_size, cols, rows, label_w, label_h, spacing_x, spacing_y, margin_top, margin_left)

        return send_file(output_file, as_attachment=True)

    return render_template('index.html')

print("Flask démarre...")
app.run(debug=True)

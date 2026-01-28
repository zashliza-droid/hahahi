from flask import Flask, render_template, request, send_file
import pandas as pd
import os

from docx import Document
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# ===============================
# PATH
# ===============================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads", "excel")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "uploads", "hasil_excel")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ===============================
# GLOBAL DATA
# ===============================
df_global = None
kolom_kode = None

# ===============================
# FORMAT NOMINAL (1.000.000)
# ===============================
def format_nominal(val):
    try:
        if pd.isna(val):
            return ""
        num = float(val)
        return f"{int(num):,}".replace(",", ".")
    except:
        return str(val)

# ===============================
# INDEX
# ===============================
@app.route("/")
def index():
    return render_template(
        "index.html",
        projects=[],
        message="Silakan upload file Excel"
    )

# ===============================
# UPLOAD EXCEL
# ===============================
@app.route("/upload", methods=["POST"])
def upload():
    global df_global, kolom_kode

    file = request.files.get("file")
    if not file:
        return "File tidak ditemukan"

    path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(path)

    # Cari header otomatis
    df_raw = pd.read_excel(path, header=None)
    header_row = None
    for i, row in df_raw.iterrows():
        if row.notna().sum() >= 2:
            header_row = i
            break

    if header_row is None:
        return "Header tidak ditemukan"

    df = pd.read_excel(path, header=header_row)

    # Bersihkan kolom
    df = df.replace(r'^\s*$', pd.NA, regex=True)
    df = df.dropna(axis=1, how="all")
    df = df.loc[:, ~df.columns.str.contains("^Unnamed", na=False)]
    df = df.reset_index(drop=True)

    df_global = df

    # Cari kolom kode kegiatan
    kolom_kode = next(
        (c for c in df.columns
         if "kode kegiatan" in str(c).lower()
         or "kode keg" in str(c).lower()),
        None
    )

    if kolom_kode is None:
        return "Kolom Kode Kegiatan tidak ditemukan"

    kode_list = (
        df[kolom_kode]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    return render_template(
        "index.html",
        projects=kode_list,
        message=None
    )

# ===============================
# EXPORT WORD
# ===============================
def buat_word(data, kode):
    path = os.path.join(OUTPUT_FOLDER, f"{kode}.docx")
    doc = Document()
    doc.add_heading(f"Data Kode Kegiatan: {kode}", 1)

    table = doc.add_table(rows=1, cols=len(data.columns))
    table.style = "Table Grid"

    for i, col in enumerate(data.columns):
        table.rows[0].cells[i].text = str(col)

    for _, row in data.iterrows():
        cells = table.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = format_nominal(val)

    doc.save(path)
    return path

# ===============================
# EXPORT PDF (RAPI)
# ===============================
def buat_pdf(data, kode):
    path = os.path.join(OUTPUT_FOLDER, f"{kode}.pdf")

    pdf = SimpleDocTemplate(
        path,
        pagesize=landscape(A4),
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20
    )

    styles = getSampleStyleSheet()
    style = styles["Normal"]
    style.fontSize = 8
    style.leading = 9

    # Hitung lebar kolom otomatis
    col_widths = []
    for col in data.columns:
        max_len = max(
            data[col].fillna("").astype(str).map(len).max(),
            len(str(col))
        )
        col_widths.append(min(max_len * 6, 120))

    table_data = [[Paragraph(str(c), style) for c in data.columns]]

    for _, row in data.iterrows():
        table_data.append(
            [Paragraph(format_nominal(v), style) for v in row]
        )

    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))

    title = Paragraph(f"Data Kode Kegiatan: {kode}", styles["Heading2"])
    pdf.build([title, table])

    return path

# ===============================
# DETAIL
# ===============================
@app.route("/detail/<kode>")
def detail(kode):
    global df_global, kolom_kode

    if df_global is None:
        return "Data belum diupload"

    data = df_global[df_global[kolom_kode].astype(str) == kode]

    # Bersihkan lagi (AMAN)
    data = data.loc[:, ~data.columns.str.contains("^Unnamed", na=False)]
    data = data.dropna(axis=1, how="all")

    if data.empty:
        return "Data tidak ditemukan"

    # Excel
    excel_path = os.path.join(OUTPUT_FOLDER, f"{kode}.xlsx")
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        data.to_excel(writer, index=False)
        ws = writer.sheets["Sheet1"]

        for i, col in enumerate(data.columns, 1):
            max_len = max(
                data[col].fillna("").astype(str).map(len).max(),
                len(str(col))
            )
            ws.column_dimensions[get_column_letter(i)].width = max_len + 4

    word_path = buat_word(data, kode)
    pdf_path = buat_pdf(data, kode)

    return render_template(
        "detail.html",
        kode=kode,
        table=data.apply(lambda c: c.map(format_nominal)).to_html(index=False),
        excel=os.path.basename(excel_path),
        word=os.path.basename(word_path),
        pdf=os.path.basename(pdf_path)
    )

# ===============================
# DOWNLOAD
# ===============================
@app.route("/download/<filename>")
def download(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    return "File tidak ditemukan"

# ===============================
# RUN (RAILWAY)
# ===============================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)

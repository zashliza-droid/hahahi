from flask import Flask, render_template, request, send_file
import pandas as pd
import os

from docx import Document
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers

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
# FORMAT
# ===============================
def format_nominal(val):
    try:
        if pd.isna(val):
            return ""
        if isinstance(val, (int, float)):
            return f"{int(val):,}".replace(",", ".")
        return str(val)
    except:
        return str(val)

def format_datetime(val):
    if isinstance(val, pd.Timestamp):
        return val.strftime("%d-%m-%Y")
    return val

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
# UPLOAD
# ===============================
@app.route("/upload", methods=["POST"])
def upload():
    global df_global, kolom_kode

    file = request.files.get("file")
    if not file:
        return "File tidak ditemukan"

    path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(path)

    df_raw = pd.read_excel(path, header=None)
    header_row = None

    for i, row in df_raw.iterrows():
        if row.notna().sum() >= 2:
            header_row = i
            break

    if header_row is None:
        return "Header tidak ditemukan"

    df = pd.read_excel(path, header=header_row)
    df = df.replace(r'^\s*$', pd.NA, regex=True)
    df = df.dropna(axis=1, how="all")
    df = df.loc[:, ~df.columns.str.contains("^Unnamed", na=False)]
    df = df.reset_index(drop=True)

    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].apply(format_datetime)

    df_global = df

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
# WORD
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
# PDF (AUTO FIT)
# ===============================
def buat_pdf(data, kode, kolom_fleksibel=None):
    path = os.path.join(OUTPUT_FOLDER, f"{kode}.pdf")

    if kolom_fleksibel is None:
        kolom_fleksibel = []

    doc = SimpleDocTemplate(
        path,
        pagesize=landscape(A4),
        leftMargin=15,
        rightMargin=15,
        topMargin=15,
        bottomMargin=15
    )

    styles = getSampleStyleSheet()
    style = styles["Normal"]
    style.fontSize = 7
    style.leading = 9

    # ===============================
    # LEBAR HALAMAN YANG REAL
    # ===============================
    page_width, _ = landscape(A4)
    usable_width = page_width - doc.leftMargin - doc.rightMargin

    col_widths = []
    fixed_total = 0

    # ===============================
    # HITUNG KOLOM
    # ===============================
    for col in data.columns:
        if col in kolom_fleksibel:
            col_widths.append(None)
        else:
            max_len = max(
                data[col].astype(str).map(len).max(),
                len(str(col))
            )
            w = min(max(max_len * 5, 45), 85)
            col_widths.append(w)
            fixed_total += w

    fleksibel_count = col_widths.count(None)

    # ===============================
    # SISA LEBAR
    # ===============================
    sisa = usable_width - fixed_total

    if fleksibel_count > 0:
        per_fleksibel = sisa / fleksibel_count
    else:
        per_fleksibel = sisa

    col_widths = [
        per_fleksibel if w is None else w
        for w in col_widths
    ]

    # ===============================
    # 🔴 PAKSA TOTAL WIDTH = PAGE WIDTH
    # ===============================
    total_col_width = sum(col_widths)
    scale = usable_width / total_col_width
    col_widths = [w * scale for w in col_widths]

    # ===============================
    # DATA TABLE
    # ===============================
    table_data = [[Paragraph(str(c), style) for c in data.columns]]

    for _, row in data.iterrows():
        table_data.append(
            [Paragraph(format_nominal(v), style) for v in row]
        )

    table = Table(
        table_data,
        colWidths=col_widths,
        repeatRows=1,
        hAlign="LEFT"   # 🔥 INI KUNCI
    )

    table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
    ]))

    title = Paragraph(
        f"<b>Data Kode Kegiatan: {kode}</b>",
        styles["Heading3"]
    )

    doc.build([title, table])
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
    if data.empty:
        return "Data tidak ditemukan"

    # ===== EXCEL =====
    excel_path = os.path.join(OUTPUT_FOLDER, f"{kode}.xlsx")
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        data.to_excel(writer, index=False)
        ws = writer.sheets["Sheet1"]

        for col_idx, col in enumerate(data.columns, 1):
            max_len = len(str(col))

            for row_idx, val in enumerate(data[col], 2):
                cell = ws.cell(row=row_idx, column=col_idx)
                if isinstance(val, (int, float)):
                    cell.number_format = '#,##0'
                max_len = max(max_len, len(str(val)))

            ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 4

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
# FILE PUBLIC
# ===============================
@app.route("/preview-excel/<kode>")
def preview_excel(kode):
    global df_global, kolom_kode

    if df_global is None:
        return "Data belum diupload"

    data = df_global[df_global[kolom_kode].astype(str) == kode]
    if data.empty:
        return "Data tidak ditemukan"

    # ===============================
    # PASTIKAN FILE EXCEL ADA
    # ===============================
    excel_path = os.path.join(OUTPUT_FOLDER, f"{kode}.xlsx")

    if not os.path.exists(excel_path):
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            data.to_excel(writer, index=False)
            ws = writer.sheets["Sheet1"]

            for col_idx, col in enumerate(data.columns, 1):
                max_len = len(str(col))
                for row_idx, val in enumerate(data[col], 2):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if isinstance(val, (int, float)):
                        cell.number_format = '#,##0'
                    max_len = max(max_len, len(str(val)))
                ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 4

    # ===============================
    # GOOGLE VIEWER
    # ===============================
    excel_url = request.host_url.rstrip("/") + "/files/" + kode + ".xlsx"
    google_viewer = f"https://docs.google.com/gview?url={excel_url}&embedded=true"

    return render_template(
        "preview_excel.html",
        excel_url=google_viewer,
        kode=kode
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
# RUN
# ===============================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)



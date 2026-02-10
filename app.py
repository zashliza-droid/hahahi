from flask import Flask, render_template, request, send_file, abort
import pandas as pd
import os
import uuid

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
# CACHE
# ===============================
DATA_CACHE = {}

# ===============================
# FORMAT
# ===============================
def format_nominal(val):
    if pd.isna(val):
        return ""
    if isinstance(val, (int, float)):
        return f"{int(val):,}".replace(",", ".")
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
    return render_template("index.html", projects=[], message="Silakan upload file Excel")

# ===============================
# UPLOAD
# ===============================
@app.route("/upload", methods=["POST"])
def upload():
    file = request.files.get("file")
    if not file:
        return "File tidak ditemukan", 400

    session_id = str(uuid.uuid4())
    path = os.path.join(UPLOAD_FOLDER, f"{session_id}_{file.filename}")
    file.save(path)

    df_raw = pd.read_excel(path, header=None)
    header_row = next((i for i, r in df_raw.iterrows() if r.notna().sum() >= 2), None)
    if header_row is None:
        return "Header tidak ditemukan", 400

    df = pd.read_excel(path, header=header_row)
    df = df.dropna(axis=1, how="all")
    df = df.loc[:, ~df.columns.str.contains("^Unnamed", na=False)]
    df = df.reset_index(drop=True)

    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].apply(format_datetime)

    kolom_kode = next(
        (c for c in df.columns if "kode kegiatan" in c.lower() or "kode keg" in c.lower()),
        None
    )
    if not kolom_kode:
        return "Kolom Kode Kegiatan tidak ditemukan", 400

    DATA_CACHE[session_id] = {"df": df, "kode_col": kolom_kode}

    kode_list = df[kolom_kode].dropna().astype(str).unique().tolist()

    return render_template(
        "index.html",
        projects=kode_list,
        session_id=session_id,
        message=None
    )

# ===============================
# WORD
# ===============================
def buat_word(data, filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    doc = Document()
    table = doc.add_table(rows=1, cols=len(data.columns))
    table.style = "Table Grid"

    for i, col in enumerate(data.columns):
        table.rows[0].cells[i].text = str(col)

    for _, row in data.iterrows():
        cells = table.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = format_nominal(val)

    doc.save(path)

# ===============================
# PDF
# ===============================
def buat_pdf(data, filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    style = getSampleStyleSheet()["Normal"]
    style.fontSize = 7

    table_data = [[Paragraph(str(c), style) for c in data.columns]]
    for _, row in data.iterrows():
        table_data.append([Paragraph(format_nominal(v), style) for v in row])

    table = Table(table_data, repeatRows=1)
    table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
    ]))

    doc = SimpleDocTemplate(path, pagesize=landscape(A4))
    doc.build([table])

# ===============================
# DETAIL
# ===============================
@app.route("/detail/<session_id>/<kode>")
def detail(session_id, kode):
    data_pack = DATA_CACHE.get(session_id)
    if not data_pack:
        return "Session tidak valid", 400

    df = data_pack["df"]
    kolom_kode = data_pack["kode_col"]
    data = df[df[kolom_kode].astype(str) == kode]

    excel = f"{session_id}_{kode}.xlsx"
    word = f"{session_id}_{kode}.docx"
    pdf = f"{session_id}_{kode}.pdf"

    excel_path = os.path.join(OUTPUT_FOLDER, excel)
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        data.to_excel(writer, index=False)
        ws = writer.sheets["Sheet1"]
        for i, col in enumerate(data.columns, 1):
            ws.column_dimensions[get_column_letter(i)].width = 20

    buat_word(data, word)
    buat_pdf(data, pdf)

    return render_template(
        "detail.html",
        kode=kode,
        session_id=session_id,
        excel=excel,
        word=word,
        pdf=pdf,
        table=data.apply(lambda c: c.map(format_nominal)).to_html(index=False)
    )

# ===============================
# PREVIEW EXCEL (CHROME)
# ===============================
@app.route("/preview-excel/<filename>")
def preview_excel(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    if not os.path.exists(path):
        abort(404)

    df = pd.read_excel(path)
    table = df.apply(lambda c: c.map(format_nominal)).to_html(index=False)

    return render_template("preview.html", table=table)


@app.route("/excel-online/<filename>")
def excel_online(filename):
    file_url = request.host_url.rstrip("/") + "/open-excel/" + filename

    excel_url = (
        "https://excel.officeapps.live.com/x/_layouts/xlviewerinternal.aspx?"
        "WOPISrc=" + file_url
    )

    return redirect(excel_url)

@app.route("/view-excel/<filename>")
def view_excel(filename):
    file_url = request.host_url.rstrip("/") + "/open-excel/" + filename
    return redirect("https://docs.google.com/gview?url=" + file_url)
# ===============================
# OPEN & DOWNLOAD
# ===============================
@app.route("/open-excel/<filename>")
def open_excel(filename):
    return send_file(os.path.join(OUTPUT_FOLDER, filename), as_attachment=False)

@app.route("/download/<filename>")
def download(filename):
    return send_file(os.path.join(OUTPUT_FOLDER, filename), as_attachment=True)

# ===============================
# RUN
# ===============================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))



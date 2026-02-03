from flask import Flask, render_template, request, redirect, send_file
import pandas as pd
import os

from docx import Document
from reportlab.platypus import SimpleDocTemplate, Table
from reportlab.lib.pagesizes import A4

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ===============================
# FUNGSI BANTU (WAJIB)
# ===============================
def load_data():
    data_path = os.path.join(UPLOAD_FOLDER, "data.pkl")
    kolom_path = os.path.join(UPLOAD_FOLDER, "kolom.txt")

    if not os.path.exists(data_path):
        return None, None

    df = pd.read_pickle(data_path)
    with open(kolom_path) as f:
        kolom_kode = f.read().strip()

    return df, kolom_kode


# ===============================
# HALAMAN UTAMA
# ===============================
@app.route("/")
def index():
    projects = []

    df, kolom_kode = load_data()
    if df is not None:
        projects = sorted(df[kolom_kode].astype(str).unique())

    return render_template("index.html", projects=projects)


# ===============================
# UPLOAD EXCEL
# ===============================
@app.route("/upload", methods=["POST"])
def upload():
    file = request.files.get("file")
    if not file:
        return redirect("/")

    df = pd.read_excel(file)

    # cari kolom kode kegiatan otomatis
    kolom_kode = None
    for col in df.columns:
        if "kode" in col.lower():
            kolom_kode = col
            break

    if not kolom_kode:
        return "Kolom kode kegiatan tidak ditemukan", 400

    df.to_pickle(os.path.join(UPLOAD_FOLDER, "data.pkl"))
    with open(os.path.join(UPLOAD_FOLDER, "kolom.txt"), "w") as f:
        f.write(kolom_kode)

    return redirect("/")


# ===============================
# DETAIL KODE KEGIATAN
# ===============================
@app.route("/detail/<kode>")
def detail(kode):
    df, kolom_kode = load_data()
    if df is None:
        return "Data belum diupload", 400

    data = df[df[kolom_kode].astype(str) == kode]
    if data.empty:
        return "Data tidak ditemukan", 404

    return render_template(
        "detail.html",
        kode=kode,
        kolom=data.columns.tolist(),
        tabel=data.to_dict(orient="records")
    )


# ===============================
# EXCEL (BROWSER / APK)
# ===============================
@app.route("/open-excel/<kode>")
def open_excel(kode):
    df, kolom_kode = load_data()
    if df is None:
        return "Data belum diupload", 400

    data = df[df[kolom_kode].astype(str) == kode]
    if data.empty:
        return "Data tidak ditemukan", 404

    path = os.path.join(OUTPUT_FOLDER, f"{kode}.xlsx")
    data.to_excel(path, index=False)

    return send_file(path, as_attachment=True)


# ===============================
# WORD
# ===============================
@app.route("/download/word/<kode>")
def download_word(kode):
    df, kolom_kode = load_data()
    data = df[df[kolom_kode].astype(str) == kode]

    path = os.path.join(OUTPUT_FOLDER, f"{kode}.docx")

    doc = Document()
    doc.add_heading(f"Kode Kegiatan {kode}", level=1)

    table = doc.add_table(rows=1, cols=len(data.columns))
    for i, col in enumerate(data.columns):
        table.rows[0].cells[i].text = col

    for _, row in data.iterrows():
        cells = table.add_row().cells
        for i, col in enumerate(data.columns):
            cells[i].text = str(row[col])

    doc.save(path)
    return send_file(path, as_attachment=True)


# ===============================
# PDF
# ===============================
@app.route("/download/pdf/<kode>")
def download_pdf(kode):
    df, kolom_kode = load_data()
    data = df[df[kolom_kode].astype(str) == kode]

    path = os.path.join(OUTPUT_FOLDER, f"{kode}.pdf")

    pdf = SimpleDocTemplate(path, pagesize=A4)
    table_data = [data.columns.tolist()] + data.values.tolist()
    table = Table(table_data)

    pdf.build([table])
    return send_file(path, as_attachment=True)


# ===============================
# JALANKAN SERVER (PORT SERVER)
# ===============================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

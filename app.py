from flask import Flask, render_template, request, redirect, send_file
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ===============================
# HALAMAN UTAMA
# ===============================
@app.route("/")
def index():
    projects = []

    data_path = os.path.join(UPLOAD_FOLDER, "data.pkl")
    kolom_path = os.path.join(UPLOAD_FOLDER, "kolom.txt")

    if os.path.exists(data_path) and os.path.exists(kolom_path):
        df = pd.read_pickle(data_path)
        with open(kolom_path) as f:
            kolom_kode = f.read().strip()

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

    # otomatis cari kolom kode kegiatan
    kolom_kode = None
    for col in df.columns:
        if "kode" in col.lower():
            kolom_kode = col
            break

    if not kolom_kode:
        return "Kolom kode kegiatan tidak ditemukan", 400

    # SIMPAN KE FILE (AMAN PUBLIK)
    df.to_pickle(os.path.join(UPLOAD_FOLDER, "data.pkl"))
    with open(os.path.join(UPLOAD_FOLDER, "kolom.txt"), "w") as f:
        f.write(kolom_kode)

    return redirect("/")


# ===============================
# DETAIL KODE KEGIATAN
# ===============================
@app.route("/detail/<kode>")
def detail(kode):

    data_path = os.path.join(UPLOAD_FOLDER, "data.pkl")
    kolom_path = os.path.join(UPLOAD_FOLDER, "kolom.txt")

    if not os.path.exists(data_path):
        return "Data belum diupload", 400

    df = pd.read_pickle(data_path)

    with open(kolom_path) as f:
        kolom_kode = f.read().strip()

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
# DOWNLOAD / BUKA EXCEL
# ===============================
@app.route("/open-excel/<kode>")
def open_excel(kode):

    data_path = os.path.join(UPLOAD_FOLDER, "data.pkl")
    kolom_path = os.path.join(UPLOAD_FOLDER, "kolom.txt")

    if not os.path.exists(data_path):
        return "Data belum diupload", 400

    df = pd.read_pickle(data_path)

    with open(kolom_path) as f:
        kolom_kode = f.read().strip()

    data = df[df[kolom_kode].astype(str) == kode]
    if data.empty:
        return "Data tidak ditemukan", 404

    excel_path = os.path.join(OUTPUT_FOLDER, f"{kode}.xlsx")

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        data.to_excel(writer, index=False)

    return send_file(excel_path, as_attachment=True)


# ===============================
# JALANKAN APP
# ===============================
if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)


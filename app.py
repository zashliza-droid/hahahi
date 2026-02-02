import os
import pandas as pd
from flask import (
    Flask, render_template, request,
    redirect, url_for, send_from_directory
)
from werkzeug.utils import secure_filename
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

# ===============================
# CONFIG
# ===============================
app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# ===============================
# GLOBAL DATA
# ===============================
df_global = None
kolom_kode = None

# ===============================
# ROUTES
# ===============================

@app.route("/", methods=["GET", "POST"])
def index():
    global df_global, kolom_kode

    if request.method == "POST":
        file = request.files.get("file")
        kolom_kode = request.form.get("kolom_kode")

        if not file or not kolom_kode:
            return "File atau kolom kode belum dipilih"

        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)

        df_global = pd.read_excel(filepath)
        df_global.columns = df_global.columns.astype(str)

        return redirect(url_for("index"))

    kolom_list = list(df_global.columns) if df_global is not None else []
    return render_template("index.html", kolom_list=kolom_list)


@app.route("/detail/<kode>")
def detail(kode):
    global df_global, kolom_kode

    if df_global is None:
        return "Data belum diupload"

    data = df_global[df_global[kolom_kode].astype(str) == kode]
    if data.empty:
        return "Data tidak ditemukan"

    return render_template(
        "detail.html",
        kode=kode,
        data=data.to_dict(orient="records"),
        columns=data.columns
    )


# ===============================
# PREVIEW EXCEL (ONLINE)
# ===============================
@app.route("/preview-excel/<kode>", endpoint="preview_excel_page")
def preview_excel_page(kode):
    global df_global, kolom_kode

    if df_global is None:
        return "Data belum tersedia"

    data = df_global[df_global[kolom_kode].astype(str) == kode]
    if data.empty:
        return "Data tidak ditemukan"

    excel_path = os.path.join(OUTPUT_FOLDER, f"{kode}.xlsx")

    # ===== CREATE EXCEL =====
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name="Data")
        ws = writer.sheets["Data"]

        for col_idx, col_name in enumerate(data.columns, 1):
            max_len = len(str(col_name))
            is_text_column = False

            for row_idx, val in enumerate(data[col_name], 2):
                cell = ws.cell(row=row_idx, column=col_idx)

                # format angka
                if isinstance(val, (int, float)):
                    cell.number_format = '#,##0'
                else:
                    is_text_column = True

                max_len = max(max_len, len(str(val)))

            # lebar kolom
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 50)

            # hanya kolom teks yang wrap
            if is_text_column:
                for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.alignment = Alignment(wrap_text=True, vertical="top")

    # ===== GOOGLE VIEWER =====
    excel_url = request.host_url.rstrip("/") + url_for("download_file", filename=f"{kode}.xlsx")
    google_viewer = f"https://docs.google.com/gview?url={excel_url}&embedded=true"

    return render_template(
        "preview_excel.html",
        excel_url=google_viewer,
        kode=kode
    )


# ===============================
# FILE SERVING (PUBLIC)
# ===============================
@app.route("/files/<filename>")
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=False)


# ===============================
# RUN LOCAL
# ===============================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)

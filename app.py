import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, send_from_directory
from werkzeug.utils import secure_filename
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

df_global = None
kolom_kode = None


@app.route("/", methods=["GET", "POST"])
def index():
    global df_global, kolom_kode

    if request.method == "POST":
        file = request.files.get("file")
        kolom_kode = request.form.get("kolom_kode")

        if not file or not kolom_kode:
            return "File atau kolom kode belum dipilih"

        filename = secure_filename(file.filename)
        path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(path)

        df_global = pd.read_excel(path)
        df_global.columns = df_global.columns.astype(str)

        return redirect(url_for("index"))

    kolom_list = list(df_global.columns) if df_global is not None else []
    return render_template("index.html", kolom_list=kolom_list)


@app.route("/detail/<kode>")
def detail(kode):
    global df_global, kolom_kode

    if df_global is None:
        return "Data belum tersedia"

    data = df_global[df_global[kolom_kode].astype(str) == kode]
    if data.empty:
        return "Data tidak ditemukan"

    return render_template(
        "detail.html",
        kode=kode,
        data=data.to_dict(orient="records"),
        columns=data.columns
    )


@app.route("/preview-excel/<kode>", endpoint="preview_excel_page")
def preview_excel_page(kode):
    global df_global, kolom_kode

    data = df_global[df_global[kolom_kode].astype(str) == kode]
    if data.empty:
        return "Data tidak ditemukan"

    excel_path = os.path.join(OUTPUT_FOLDER, f"{kode}.xlsx")

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name="Data")
        ws = writer.sheets["Data"]

        for col_idx, col in enumerate(data.columns, 1):
            max_len = len(col)
            is_text = False

            for row_idx, val in enumerate(data[col], 2):
                cell = ws.cell(row=row_idx, column=col_idx)
                if isinstance(val, (int, float)):
                    cell.number_format = '#,##0'
                else:
                    is_text = True
                max_len = max(max_len, len(str(val)))

            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 50)

            if is_text:
                for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.alignment = Alignment(wrap_text=True, vertical="top")

    excel_url = request.host_url.rstrip("/") + url_for(
        "download_file", filename=f"{kode}.xlsx"
    )
    viewer = f"https://docs.google.com/gview?url={excel_url}&embedded=true"

    return render_template(
        "preview_excel.html",
        excel_url=viewer,
        kode=kode
    )


@app.route("/files/<filename>")
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename)


if __name__ == "__main__":
    app.run(debug=True)

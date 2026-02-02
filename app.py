from flask import Flask, render_template, request, redirect, url_for, send_from_directory
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

df_global = None
kolom_kode = None


@app.route("/")
def index():
    projects = []
    if df_global is not None:
        projects = df_global[kolom_kode].astype(str).unique().tolist()
    return render_template("index.html", projects=projects, message="")


@app.route("/upload", methods=["POST"])
def upload():
    global df_global, kolom_kode

    file = request.files["file"]
    path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(path)

    df_global = pd.read_excel(path)
    kolom_kode = df_global.columns[0]

    return redirect("/")


@app.route("/detail/<kode>")
def detail(kode):
    data = df_global[df_global[kolom_kode].astype(str) == kode]

    return render_template(
        "detail.html",
        kode=kode,
        table=data.to_html(index=False),
        excel=f"{kode}.xlsx",
        word=f"{kode}.docx",
        pdf=f"{kode}.pdf"
    )

# =========================
# AUTO OPEN EXCEL
# =========================
@app.route("/open-excel/<kode>")
def open_excel(kode):
    ua = request.headers.get("User-Agent", "").lower()
    excel_url = request.host_url.rstrip("/") + f"/files/{kode}.xlsx"

    if "android" in ua or "iphone" in ua:
        return redirect(
            f"intent://{excel_url.replace('https://','')}#Intent;"
            "scheme=https;"
            "package=com.microsoft.office.excel;"
            "end;"
        )
    else:
        return redirect(
            f"https://view.officeapps.live.com/op/view.aspx?src={excel_url}"
        )


@app.route("/files/<filename>")
def files(filename):
    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=False)


@app.route("/download/<filename>")
def download(filename):
    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)

@app.route("/preview-html/<kode>")
def preview_html(kode):
    data = df_global[df_global[kolom_kode].astype(str) == kode]

    return render_template(
        "preview_html.html",
        data=data.to_dict(orient="records"),
        columns=data.columns,
        kode=kode
    )


if __name__ == "__main__":
    app.run(debug=True)



from flask import Flask, render_template, request, redirect, url_for, send_file
from werkzeug.utils import secure_filename
import os
import tabula
import pandas as pd

app = Flask(__name__)

# Defina as extensões permitidas
ALLOWED_EXTENSIONS = {"pdf"}


# Função para verificar se a extensão do arquivo é permitida
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# Rota principal para carregar a página inicial
@app.route("/")
def index():
    return render_template("index.html")


# Rota para a página de sucesso após a conversão
@app.route("/success")
def success():
    filenames = request.args.getlist("filename")
    return render_template("success.html", filenames=filenames)


# Rota para baixar o arquivo convertido
@app.route("/download/<filename>")
def download(filename):
    return send_file(filename, as_attachment=True)


# Função para converter o PDF em Excel
def convert_pdf_to_excel(file):
    df = tabula.read_pdf(file, pages="all", multiple_tables=True)
    concatenated_df = pd.concat(df)
    return concatenated_df


# Rota para a conversão do PDF para Excel
@app.route("/converter", methods=["POST"])
def converter():
    if "files" not in request.files:
        return redirect(request.url)

    files = request.files.getlist("files")
    conversion_option = request.form.get("conversion")

    if len(files) == 0:
        return redirect(request.url)

    try:
        if conversion_option == "single":
            # Converter para um único arquivo Excel
            concatenated_dfs = [
                convert_pdf_to_excel(file)
                for file in files
                if file.filename and allowed_file(file.filename)
            ]
            concatenated_df = pd.concat(concatenated_dfs)
            excel_filename = "converted_file.xlsx"
            excel_path = os.path.join(app.root_path, excel_filename)
            concatenated_df.to_excel(excel_path, index=False)
            return redirect(url_for("success", filename=excel_filename))
        elif conversion_option == "multiple":
            # Converter e baixar cada arquivo individualmente
            filenames = []
            for file in files:
                if file.filename == "" or not allowed_file(file.filename):
                    continue
                filename = secure_filename(file.filename)
                excel_filename = os.path.splitext(filename)[0] + ".xlsx"
                excel_path = os.path.join(app.root_path, excel_filename)
                df = convert_pdf_to_excel(file)
                df.to_excel(excel_path, index=False)
                filenames.append(excel_filename)
            return redirect(url_for("success", filename=filenames))
    except Exception as e:
        print(f"Erro ao converter PDFs para Excel: {e}")
        return "Erro ao converter PDFs para Excel"


if __name__ == "__main__":
    app.run(host="192.168.0.105", debug=True, port=80)

from flask import Flask, request, send_file, jsonify
import camelot
import pandas as pd
import uuid
import os

app = Flask(__name__)

@app.route("/convert", methods=["POST"])
def convert_pdf():
    if "pdf" not in request.files:
        return jsonify({"error": "Nenhum arquivo PDF enviado"}), 400

    pdf_file = request.files["pdf"]
    pdf_path = f"/tmp/{uuid.uuid4()}.pdf"
    pdf_file.save(pdf_path)

    try:
        tables = camelot.read_pdf(pdf_path, pages="all", flavor="stream")
        if not tables:
            return jsonify({"error": "Não encontrei tabelas"}), 422

        xlsx_path = f"/tmp/{uuid.uuid4()}.xlsx"
        with pd.ExcelWriter(xlsx_path) as writer:
            for i, t in enumerate(tables):
                t.df.to_excel(writer, sheet_name=f"Tabela_{i+1}", index=False)

        return send_file(
            xlsx_path,
            as_attachment=True,
            download_name="planilha.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    finally:
        os.remove(pdf_path)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)

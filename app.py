from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import camelot
import pandas as pd
import uuid
import os
import logging
from werkzeug.utils import secure_filename
import pdf2image
import pytesseract
from PIL import Image
import io
import cv2
import numpy as np

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

CORS(app, origins=[
    "https://upload-agent.vercel.app",
    "https://upload-agent-git-main-ghalba-vieiras-projects-f2e9b128.vercel.app",
    "http://localhost:3000",
    "http://127.0.0.1:3000"
])

def extract_tables_with_ocr(pdf_path):
    """Extrai tabelas usando OCR quando Camelot falha"""
    
    try:
        # Converter PDF para imagens
        images = pdf2image.convert_from_path(pdf_path, dpi=300)
        all_data = []
        
        for page_num, image in enumerate(images):
            logging.info(f"Processando página {page_num + 1} com OCR")
            
            # Converter PIL para OpenCV
            opencv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            
            # Pré-processamento da imagem
            gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
            
            # Threshold para melhorar contraste
            _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            # OCR com configuração para tabelas
            custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.,;:!?()[]{}\"\'\ -'
            
            # Extrair texto
            text = pytesseract.image_to_string(thresh, config=custom_config, lang='por')
            
            if text.strip():
                # Tentar identificar estrutura de tabela
                lines = [line.strip() for line in text.split('\n') if line.strip()]
                
                if lines:
                    page_data = {
                        'page': page_num + 1,
                        'text': text,
                        'lines': lines
                    }
                    all_data.append(page_data)
        
        return all_data
        
    except Exception as e:
        logging.error(f"Erro no OCR: {e}")
        return []

def convert_ocr_to_dataframe(ocr_data):
    """Converte dados OCR em DataFrame estruturado"""
    
    all_tables = []
    
    for page_data in ocr_data:
        lines = page_data['lines']
        
        # Tentar identificar padrões de tabela
        table_lines = []
        for line in lines:
            # Simples heurística: linhas com múltiplas palavras/números
            if len(line.split()) > 1:
                table_lines.append(line.split())
        
        if table_lines:
            # Criar DataFrame
            max_cols = max(len(row) for row in table_lines) if table_lines else 0
            
            # Padronizar número de colunas
            standardized_rows = []
            for row in table_lines:
                while len(row) < max_cols:
                    row.append('')
                standardized_rows.append(row[:max_cols])
            
            if standardized_rows:
                df = pd.DataFrame(standardized_rows[1:], columns=standardized_rows[0] if standardized_rows else None)
                all_tables.append(df)
    
    return all_tables

@app.route("/convert", methods=["POST"])
def convert_pdf():
    pdf_path = None
    xlsx_path = None
    
    try:
        if "pdf" not in request.files:
            return jsonify({"error": "Nenhum arquivo PDF enviado"}), 400
        
        pdf_file = request.files["pdf"]
        
        if pdf_file.filename == '':
            return jsonify({"error": "Nenhum arquivo selecionado"}), 400
        
        # Salvar arquivo temporário
        filename = secure_filename(pdf_file.filename)
        pdf_path = f"/tmp/{uuid.uuid4()}_{filename}"
        pdf_file.save(pdf_path)
        
        logging.info(f"Arquivo salvo: {pdf_path}")
        
        # Método 1: Tentar Camelot primeiro
        tables = []
        camelot_success = False
        
        try:
            tables = camelot.read_pdf(pdf_path, pages="all", flavor="stream")
            if len(tables) > 0:
                # Verificar se as tabelas têm conteúdo útil
                valid_tables = []
                for table in tables:
                    if not table.df.empty:
                        has_content = table.df.astype(str).apply(lambda x: x.str.strip().str.len() > 0).any().any()
                        if has_content:
                            valid_tables.append(table)
                
                if valid_tables:
                    tables = valid_tables
                    camelot_success = True
                    logging.info(f"Camelot encontrou {len(tables)} tabelas válidas")
        except Exception as e:
            logging.warning(f"Camelot falhou: {e}")
        
        # Método 2: Se Camelot falhou, tentar Lattice
        if not camelot_success:
            try:
                tables = camelot.read_pdf(pdf_path, pages="all", flavor="lattice")
                if len(tables) > 0:
                    valid_tables = []
                    for table in tables:
                        if not table.df.empty:
                            has_content = table.df.astype(str).apply(lambda x: x.str.strip().str.len() > 0).any().any()
                            if has_content:
                                valid_tables.append(table)
                    
                    if valid_tables:
                        tables = valid_tables
                        camelot_success = True
                        logging.info(f"Camelot lattice encontrou {len(tables)} tabelas válidas")
            except Exception as e:
                logging.warning(f"Camelot lattice falhou: {e}")
        
        # Método 3: Se Camelot falhou completamente, usar OCR
        if not camelot_success:
            logging.info("Tentando extração com OCR...")
            
            try:
                ocr_data = extract_tables_with_ocr(pdf_path)
                
                if ocr_data:
                    ocr_tables = convert_ocr_to_dataframe(ocr_data)
                    
                    # Converter para formato compatível com Camelot
                    tables = []
                    for df in ocr_tables:
                        if not df.empty:
                            # Criar objeto similar ao Camelot
                            class OCRTable:
                                def __init__(self, dataframe):
                                    self.df = dataframe
                            
                            tables.append(OCRTable(df))
                    
                    logging.info(f"OCR encontrou {len(tables)} tabelas")
                else:
                    return jsonify({
                        "error": "Nenhuma tabela encontrada",
                        "details": "Tentamos métodos de extração automática e OCR, mas não conseguimos encontrar tabelas válidas",
                        "suggestions": [
                            "Verifique se o PDF contém tabelas reais",
                            "Certifique-se de que a qualidade do PDF é boa",
                            "Tente usar um PDF com tabelas mais simples"
                        ]
                    }), 422
            
            except Exception as e:
                logging.error(f"OCR falhou: {e}")
                return jsonify({
                    "error": "Falha na extração de tabelas",
                    "details": f"Tanto Camelot quanto OCR falharam: {str(e)}"
                }), 422
        
        # Verificar se encontrou tabelas
        if not tables:
            return jsonify({
                "error": "Nenhuma tabela encontrada",
                "details": "O PDF não contém tabelas extraíveis"
            }), 422
        
        # Criar arquivo Excel
        xlsx_path = f"/tmp/{uuid.uuid4()}.xlsx"
        
        with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
            for i, table in enumerate(tables):
                df = table.df.copy()
                
                # Limpeza dos dados
                df = df.dropna(how='all')
                df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
                
                # Nome da aba
                sheet_name = f"Tabela_{i+1}"[:31]
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                logging.info(f"Tabela {i+1} salva: {len(df)} linhas x {len(df.columns)} colunas")
        
        return send_file(
            xlsx_path,
            as_attachment=True,
            download_name="tabelas_extraidas.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        logging.error(f"Erro geral: {str(e)}")
        return jsonify({
            "error": "Erro interno do servidor",
            "details": str(e)
        }), 500
    
    finally:
        # Limpeza
        try:
            if pdf_path and os.path.exists(pdf_path):
                os.remove(pdf_path)
            if xlsx_path and os.path.exists(xlsx_path):
                os.remove(xlsx_path)
        except Exception as e:
            logging.warning(f"Erro na limpeza: {e}")

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 10000))
    app.run(host="0.0.0.0", port=port, debug=False)
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
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
from pptx.util import Inches as PptxInches
from datetime import datetime
import PyPDF2
import fitz  # PyMuPDF

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

CORS(app, origins=[
    "https://upload-agent.vercel.app",
    "https://upload-agent-git-main-ghalba-vieiras-projects-f2e9b128.vercel.app",
    "http://localhost:3000",
    "http://127.0.0.1:3000"
])

def extract_text_from_pdf(pdf_path):
    """Extrai texto completo do PDF para resumo"""
    try:
        text = ""
        
        try:
            doc = fitz.open(pdf_path)
            for page_num in range(doc.page_count):
                page = doc.load_page(page_num)
                text += page.get_text() + "\n"
            doc.close()
            
            if text.strip():
                return text
        except Exception as e:
            logging.warning(f"PyMuPDF falhou: {e}")
        
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page in pdf_reader.pages:
                    text += page.extract_text() + "\n"
            
            if text.strip():
                return text
        except Exception as e:
            logging.warning(f"PyPDF2 falhou: {e}")
        
        try:
            images = pdf2image.convert_from_path(pdf_path, dpi=200)
            for image in images:
                opencv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
                gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
                text += pytesseract.image_to_string(gray, lang='por') + "\n"
            
            return text
        except Exception as e:
            logging.error(f"OCR falhou: {e}")
            return ""
    
    except Exception as e:
        logging.error(f"Erro na extração de texto: {e}")
        return ""

def extract_tables_with_ocr(pdf_path):
    """Extrai tabelas usando OCR quando Camelot falha"""
    try:
        images = pdf2image.convert_from_path(pdf_path, dpi=300)
        all_data = []
        
        for page_num, image in enumerate(images):
            logging.info(f"Processando página {page_num + 1} com OCR")
            
            opencv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
            _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.,;:!?()[]{}\"\'\ -'
            text = pytesseract.image_to_string(thresh, config=custom_config, lang='por')
            
            if text.strip():
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
        
        table_lines = []
        for line in lines:
            if len(line.split()) > 1:
                table_lines.append(line.split())
        
        if table_lines:
            max_cols = max(len(row) for row in table_lines) if table_lines else 0
            
            standardized_rows = []
            for row in table_lines:
                while len(row) < max_cols:
                    row.append('')
                standardized_rows.append(row[:max_cols])
            
            if standardized_rows:
                df = pd.DataFrame(standardized_rows[1:], columns=standardized_rows[0] if standardized_rows else None)
                all_tables.append(df)
    
    return all_tables

def extract_tables_from_pdf(pdf_path):
    """Extrai tabelas do PDF usando múltiplos métodos"""
    tables = []
    camelot_success = False
    
    # Método 1: Camelot Stream
    try:
        tables = camelot.read_pdf(pdf_path, pages="all", flavor="stream")
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
                logging.info(f"Camelot stream encontrou {len(tables)} tabelas válidas")
    except Exception as e:
        logging.warning(f"Camelot stream falhou: {e}")
    
    # Método 2: Camelot Lattice
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
    
    # Método 3: OCR
    if not camelot_success:
        logging.info("Tentando extração com OCR...")
        try:
            ocr_data = extract_tables_with_ocr(pdf_path)
            if ocr_data:
                ocr_tables = convert_ocr_to_dataframe(ocr_data)
                
                tables = []
                for df in ocr_tables:
                    if not df.empty:
                        class OCRTable:
                            def __init__(self, dataframe):
                                self.df = dataframe
                        
                        tables.append(OCRTable(df))
                
                logging.info(f"OCR encontrou {len(tables)} tabelas")
        except Exception as e:
            logging.error(f"OCR falhou: {e}")
    
    return tables



@app.route("/convert/excel", methods=["POST"])
def convert_to_excel():
    """Conversão para Excel (mantém a funcionalidade original)"""
    pdf_path = None
    xlsx_path = None
    
    try:
        if "file" not in request.files:
            return jsonify({"error": "Nenhum arquivo PDF enviado"}), 400
        
        pdf_file = request.files["file"]
        
        if pdf_file.filename == '':
            return jsonify({"error": "Nenhum arquivo selecionado"}), 400
        
        filename = secure_filename(pdf_file.filename)
        pdf_path = f"/tmp/{uuid.uuid4()}_{filename}"
        pdf_file.save(pdf_path)
        
        logging.info(f"Arquivo salvo: {pdf_path}")
        
        tables = extract_tables_from_pdf(pdf_path)
        
        if not tables:
            return jsonify({
                "error": "Nenhuma tabela encontrada",
                "details": "O PDF não contém tabelas extraíveis"
            }), 422
        
        xlsx_path = f"/tmp/{uuid.uuid4()}.xlsx"
        
        with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
            for i, table in enumerate(tables):
                df = table.df.copy()
                df = df.dropna(how='all')
                df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
                
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
        try:
            if pdf_path and os.path.exists(pdf_path):
                os.remove(pdf_path)
            if xlsx_path and os.path.exists(xlsx_path):
                os.remove(xlsx_path)
        except Exception as e:
            logging.warning(f"Erro na limpeza: {e}")

@app.route("/convert/word", methods=["POST"])
def convert_to_word():
    """Conversão para Word"""
    pdf_path = None
    docx_path = None
    
    try:
        if "file" not in request.files:
            return jsonify({"error": "Nenhum arquivo PDF enviado"}), 400
        
        pdf_file = request.files["file"]
        
        if pdf_file.filename == '':
            return jsonify({"error": "Nenhum arquivo selecionado"}), 400
        
        filename = secure_filename(pdf_file.filename)
        pdf_path = f"/tmp/{uuid.uuid4()}_{filename}"
        pdf_file.save(pdf_path)
        
        logging.info(f"Convertendo para Word: {pdf_path}")
        
        # Extrair texto do PDF
        text = extract_text_from_pdf(pdf_path)
        
        if not text.strip():
            return jsonify({
                "error": "Não foi possível extrair texto do PDF",
                "details": "O PDF pode estar protegido ou não conter texto extraível"
            }), 422
        
        # Extrair tabelas
        tables = extract_tables_from_pdf(pdf_path)
        
        # Criar documento Word
        doc = Document()
        
        title = doc.add_heading('Documento Convertido do PDF', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph(f"Arquivo original: {pdf_file.filename}")
        doc.add_paragraph(f"Convertido em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        doc.add_paragraph("")
        
        if text.strip():
            doc.add_heading('Conteúdo do Documento', level=1)
            paragraphs = text.split('\n\n')
            for paragraph in paragraphs:
                if paragraph.strip():
                    doc.add_paragraph(paragraph.strip())
        
        # Adicionar tabelas
        if tables:
            doc.add_heading('Tabelas Extraídas', level=1)
            
            for i, table in enumerate(tables):
                doc.add_heading(f'Tabela {i+1}', level=2)
                
                df = table.df.copy()
                df = df.dropna(how='all')
                df = df.fillna('')
                
                # Criar tabela no Word
                if not df.empty:
                    word_table = doc.add_table(rows=len(df) + 1, cols=len(df.columns))
                    word_table.style = 'Table Grid'
                    
                    # Adicionar cabeçalhos
                    for j, column in enumerate(df.columns):
                        cell = word_table.cell(0, j)
                        cell.text = str(column) if column else f"Coluna {j+1}"
                        cell.paragraphs[0].runs[0].bold = True
                    
                    # Adicionar dados
                    for i, row in df.iterrows():
                        for j, value in enumerate(row):
                            cell = word_table.cell(i + 1, j)
                            cell.text = str(value) if pd.notna(value) else ""
                
                doc.add_paragraph("")
        
        # Salvar documento
        docx_path = f"/tmp/{uuid.uuid4()}.docx"
        doc.save(docx_path)
        
        return send_file(
            docx_path,
            as_attachment=True,
            download_name="documento_convertido.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    except Exception as e:
        logging.error(f"Erro na conversão para Word: {str(e)}")
        return jsonify({
            "error": "Erro interno do servidor",
            "details": str(e)
        }), 500
    
    finally:
        try:
            if pdf_path and os.path.exists(pdf_path):
                os.remove(pdf_path)
            if docx_path and os.path.exists(docx_path):
                os.remove(docx_path)
        except Exception as e:
            logging.warning(f"Erro na limpeza: {e}")

@app.route("/convert/powerpoint", methods=["POST"])
def convert_to_powerpoint():
    """Conversão para PowerPoint"""
    pdf_path = None
    pptx_path = None
    
    try:
        if "file" not in request.files:
            return jsonify({"error": "Nenhum arquivo PDF enviado"}), 400
        
        pdf_file = request.files["file"]
        
        if pdf_file.filename == '':
            return jsonify({"error": "Nenhum arquivo selecionado"}), 400
        
        filename = secure_filename(pdf_file.filename)
        pdf_path = f"/tmp/{uuid.uuid4()}_{filename}"
        pdf_file.save(pdf_path)
        
        logging.info(f"Convertendo para PowerPoint: {pdf_path}")
        
        # Extrair texto do PDF
        text = extract_text_from_pdf(pdf_path)
        
        if not text.strip():
            return jsonify({
                "error": "Não foi possível extrair texto do PDF",
                "details": "O PDF pode estar protegido ou não conter texto extraível"
            }), 422
        
        # Extrair tabelas
        tables = extract_tables_from_pdf(pdf_path)
        
        # Criar apresentação PowerPoint
        prs = Presentation()
        
        # Slide de título
        slide_layout = prs.slide_layouts[0]  # Title slide
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        
        title.text = "Apresentação Convertida do PDF"
        subtitle.text = f"Arquivo: {pdf_file.filename}\nConvertido em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
        
        # Dividir texto em slides
        if text.strip():
            paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]
            
            # Agrupar parágrafos em slides (máximo 5 por slide)
            slide_groups = []
            current_group = []
            
            for paragraph in paragraphs[:20]:  # Limitar a 20 parágrafos
                if len(current_group) < 5:
                    current_group.append(paragraph)
                else:
                    slide_groups.append(current_group)
                    current_group = [paragraph]
            
            if current_group:
                slide_groups.append(current_group)
            
            # Criar slides de conteúdo
            for i, group in enumerate(slide_groups):
                slide_layout = prs.slide_layouts[1]  # Title and Content
                slide = prs.slides.add_slide(slide_layout)
                
                title = slide.shapes.title
                content = slide.placeholders[1]
                
                title.text = f"Conteúdo - Página {i+1}"
                
                # Adicionar texto ao slide
                text_frame = content.text_frame
                text_frame.clear()
                
                for j, paragraph in enumerate(group):
                    if j == 0:
                        p = text_frame.paragraphs[0]
                    else:
                        p = text_frame.add_paragraph()
                    
                    p.text = paragraph[:200] + "..." if len(paragraph) > 200 else paragraph
                    p.level = 0
        
        # Adicionar tabelas em slides separados
        if tables:
            for i, table in enumerate(tables):
                slide_layout = prs.slide_layouts[1]  # Title and Content
                slide = prs.slides.add_slide(slide_layout)
                
                title = slide.shapes.title
                title.text = f"Tabela {i+1}"
                
                df = table.df.copy()
                df = df.dropna(how='all')
                df = df.fillna('')
                
                if not df.empty:
                    # Limitar tamanho da tabela para caber no slide
                    max_rows = min(10, len(df))
                    max_cols = min(6, len(df.columns))
                    
                    df_limited = df.iloc[:max_rows, :max_cols]
                    
                    # Adicionar tabela ao slide
                    left = PptxInches(1)
                    top = PptxInches(2)
                    width = PptxInches(8)
                    height = PptxInches(4)
                    
                    table_shape = slide.shapes.add_table(
                        rows=len(df_limited) + 1,
                        cols=len(df_limited.columns),
                        left=left,
                        top=top,
                        width=width,
                        height=height
                    )
                    
                    table_obj = table_shape.table
                    
                    # Adicionar cabeçalhos
                    for j, column in enumerate(df_limited.columns):
                        cell = table_obj.cell(0, j)
                        cell.text = str(column) if column else f"Col {j+1}"
                    
                    # Adicionar dados
                    for i, row in df_limited.iterrows():
                        for j, value in enumerate(row):
                            cell = table_obj.cell(i + 1, j)
                            cell.text = str(value)[:50] if pd.notna(value) else ""
        
        # Salvar apresentação
        pptx_path = f"/tmp/{uuid.uuid4()}.pptx"
        prs.save(pptx_path)
        
        return send_file(
            pptx_path,
            as_attachment=True,
            download_name="apresentacao_convertida.pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
        
    except Exception as e:
        logging.error(f"Erro na conversão para PowerPoint: {str(e)}")
        return jsonify({
            "error": "Erro interno do servidor",
            "details": str(e)
        }), 500
    
    finally:
        try:
            if pdf_path and os.path.exists(pdf_path):
                os.remove(pdf_path)
            if pptx_path and os.path.exists(pptx_path):
                os.remove(pptx_path)
        except Exception as e:
            logging.warning(f"Erro na limpeza: {e}")

@app.route("/convert/summary", methods=["POST"])
def generate_summary():
    """Extrai texto do PDF para ser usado no frontend"""
    pdf_path = None
    
    try:
        if "file" not in request.files:
            return jsonify({"error": "Nenhum arquivo PDF enviado"}), 400
        
        pdf_file = request.files["file"]
        
        if pdf_file.filename == '':
            return jsonify({"error": "Nenhum arquivo selecionado"}), 400
        
        filename = secure_filename(pdf_file.filename)
        pdf_path = f"/tmp/{uuid.uuid4()}_{filename}"
        pdf_file.save(pdf_path)
        
        logging.info(f"Extraindo texto para resumo: {pdf_path}")
        
        # Extrair texto do PDF
        text = extract_text_from_pdf(pdf_path)
        
        if not text.strip():
            return jsonify({
                "error": "Não foi possível extrair texto do PDF",
                "details": "O PDF pode estar protegido ou não conter texto extraível"
            }), 422
        
        # Retornar texto extraído para o frontend processar
        return jsonify({
            "filename": secure_filename(pdf_file.filename),
            "summary": summary or "Resumo não disponível.",
            "extracted_at": datetime.now().isoformat()
        })
        
    except Exception as e:
        logging.error(f"Erro na extração de texto: {str(e)}")
        return jsonify({
            "error": "Erro interno do servidor",
            "details": str(e)
        }), 500
    
    finally:
        try:
            if pdf_path and os.path.exists(pdf_path):
                os.remove(pdf_path)
        except Exception as e:
            logging.warning(f"Erro na limpeza: {e}")

@app.route("/health", methods=["GET"])
def health_check():
    """Health check endpoint"""
    return jsonify({
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "services": {
            "pdf_processing": "active",
            "text_extraction": "active",
            "table_extraction": "active"
        }
    })

# Manter o endpoint original para compatibilidade
@app.route("/convert", methods=["POST"])
def convert_pdf():
    """Endpoint original mantido para compatibilidade"""
    return convert_to_excel()

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 10000))
    app.run(host="0.0.0.0", port=port, debug=False)
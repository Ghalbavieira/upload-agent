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
    "http://127.0.0.1:3000",
    "https://ghalba.app.n8n.cloud"
])

def extract_text_from_pdf(pdf_path):
    """Extrai texto completo do PDF com múltiplos métodos"""
    try:
        text = ""
        
        # Método 1: PyMuPDF (melhor para texto)
        try:
            doc = fitz.open(pdf_path)
            for page_num in range(doc.page_count):
                page = doc.load_page(page_num)
                text += page.get_text() + "\n\n"
            doc.close()
            
            if text.strip():
                logging.info(f"PyMuPDF extraiu {len(text)} caracteres")
                return text
        except Exception as e:
            logging.warning(f"PyMuPDF falhou: {e}")
        
        # Método 2: PyPDF2
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page in pdf_reader.pages:
                    text += page.extract_text() + "\n\n"
            
            if text.strip():
                logging.info(f"PyPDF2 extraiu {len(text)} caracteres")
                return text
        except Exception as e:
            logging.warning(f"PyPDF2 falhou: {e}")
        
        # Método 3: OCR (para PDFs escaneados)
        try:
            logging.info("Tentando OCR para extração de texto...")
            images = pdf2image.convert_from_path(pdf_path, dpi=300)
            for i, image in enumerate(images):
                logging.info(f"Processando página {i+1} com OCR")
                
                # Converter PIL para OpenCV
                opencv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
                gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
                
                # Aplicar filtros para melhorar OCR
                clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
                gray = clahe.apply(gray)
                
                # Threshold adaptativo
                thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                             cv2.THRESH_BINARY, 11, 2)
                
                # OCR com configurações otimizadas
                custom_config = r'--oem 3 --psm 6 -l por'
                page_text = pytesseract.image_to_string(thresh, config=custom_config)
                text += page_text + "\n\n"
            
            logging.info(f"OCR extraiu {len(text)} caracteres")
            return text
        except Exception as e:
            logging.error(f"OCR falhou: {e}")
            return ""
    
    except Exception as e:
        logging.error(f"Erro na extração de texto: {e}")
        return ""

def extract_images_from_pdf(pdf_path):
    """Extrai e processa imagens do PDF"""
    try:
        doc = fitz.open(pdf_path)
        images_text = ""
        
        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            image_list = page.get_images()
            
            for img_index, img in enumerate(image_list):
                try:
                    # Extrair imagem
                    xref = img[0]
                    pix = fitz.Pixmap(doc, xref)
                    
                    if pix.n - pix.alpha < 4:  # GRAY ou RGB
                        # Converter para PIL Image
                        img_data = pix.tobytes("ppm")
                        pil_image = Image.open(io.BytesIO(img_data))
                        
                        # Converter para OpenCV
                        opencv_image = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
                        gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
                        
                        # Aplicar OCR na imagem
                        custom_config = r'--oem 3 --psm 6 -l por'
                        img_text = pytesseract.image_to_string(gray, config=custom_config)
                        
                        if img_text.strip():
                            images_text += f"\n[IMAGEM {page_num+1}-{img_index+1}]\n{img_text}\n"
                    
                    pix = None
                except Exception as e:
                    logging.warning(f"Erro ao processar imagem {img_index}: {e}")
                    continue
        
        doc.close()
        return images_text
    except Exception as e:
        logging.error(f"Erro na extração de imagens: {e}")
        return ""

def extract_tables_with_ocr(pdf_path):
    """Extrai tabelas usando OCR melhorado"""
    try:
        images = pdf2image.convert_from_path(pdf_path, dpi=300)
        all_data = []
        
        for page_num, image in enumerate(images):
            logging.info(f"Processando página {page_num + 1} com OCR para tabelas")
            
            # Converter para OpenCV
            opencv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
            
            # Detectar linhas horizontais e verticais (estrutura de tabela)
            horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
            vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
            
            # Aplicar filtros
            horizontal_lines = cv2.morphologyEx(gray, cv2.MORPH_OPEN, horizontal_kernel)
            vertical_lines = cv2.morphologyEx(gray, cv2.MORPH_OPEN, vertical_kernel)
            
            # Combinar linhas
            table_mask = cv2.addWeighted(horizontal_lines, 0.5, vertical_lines, 0.5, 0.0)
            
            # Se detectar estrutura de tabela, aplicar OCR otimizado
            if cv2.countNonZero(table_mask) > 100:
                # Configuração específica para tabelas
                custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïñòóôõöùúûüý.,;:!?()[]{}\"\'\ -'
                text = pytesseract.image_to_string(gray, config=custom_config, lang='por')
                
                if text.strip():
                    lines = [line.strip() for line in text.split('\n') if line.strip()]
                    if lines:
                        page_data = {
                            'page': page_num + 1,
                            'text': text,
                            'lines': lines,
                            'has_table_structure': True
                        }
                        all_data.append(page_data)
        
        return all_data
    except Exception as e:
        logging.error(f"Erro no OCR de tabelas: {e}")
        return []

def convert_ocr_to_dataframe(ocr_data):
    """Converte dados OCR em DataFrame melhorado"""
    all_tables = []
    
    for page_data in ocr_data:
        lines = page_data['lines']
        
        # Identificar possíveis linhas de tabela
        table_lines = []
        for line in lines:
            # Separar por múltiplos espaços ou tabs
            parts = [part.strip() for part in line.split() if part.strip()]
            if len(parts) >= 2:  # Pelo menos 2 colunas
                table_lines.append(parts)
        
        if table_lines:
            # Determinar número de colunas baseado na linha com mais elementos
            max_cols = max(len(row) for row in table_lines)
            
            # Padronizar todas as linhas
            standardized_rows = []
            for row in table_lines:
                # Completar com strings vazias se necessário
                while len(row) < max_cols:
                    row.append('')
                standardized_rows.append(row[:max_cols])
            
            if standardized_rows:
                # Primeira linha como cabeçalho se parecer com cabeçalho
                if len(standardized_rows) > 1:
                    headers = standardized_rows[0]
                    data = standardized_rows[1:]
                else:
                    headers = [f'Coluna_{i+1}' for i in range(max_cols)]
                    data = standardized_rows
                
                df = pd.DataFrame(data, columns=headers)
                all_tables.append(df)
    
    return all_tables

def extract_tables_from_pdf(pdf_path):
    """Extrai tabelas do PDF com múltiplos métodos melhorados"""
    tables = []
    camelot_success = False
    
    # Método 1: Camelot Stream (melhor para tabelas sem bordas)
    try:
        tables_stream = camelot.read_pdf(pdf_path, pages="all", flavor="stream")
        if len(tables_stream) > 0:
            valid_tables = []
            for table in tables_stream:
                if not table.df.empty:
                    # Verificar se tem conteúdo significativo
                    non_empty_cells = table.df.astype(str).apply(
                        lambda x: x.str.strip().str.len() > 0
                    ).sum().sum()
                    
                    if non_empty_cells > 3:  # Pelo menos 3 células com conteúdo
                        valid_tables.append(table)
            
            if valid_tables:
                tables = valid_tables
                camelot_success = True
                logging.info(f"Camelot stream encontrou {len(tables)} tabelas válidas")
    except Exception as e:
        logging.warning(f"Camelot stream falhou: {e}")
    
    # Método 2: Camelot Lattice (melhor para tabelas com bordas)
    if not camelot_success:
        try:
            tables_lattice = camelot.read_pdf(pdf_path, pages="all", flavor="lattice")
            if len(tables_lattice) > 0:
                valid_tables = []
                for table in tables_lattice:
                    if not table.df.empty:
                        non_empty_cells = table.df.astype(str).apply(
                            lambda x: x.str.strip().str.len() > 0
                        ).sum().sum()
                        
                        if non_empty_cells > 3:
                            valid_tables.append(table)
                
                if valid_tables:
                    tables = valid_tables
                    camelot_success = True
                    logging.info(f"Camelot lattice encontrou {len(tables)} tabelas válidas")
        except Exception as e:
            logging.warning(f"Camelot lattice falhou: {e}")
    
    # Método 3: OCR melhorado para tabelas
    if not camelot_success:
        logging.info("Tentando extração de tabelas com OCR...")
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
            logging.error(f"OCR de tabelas falhou: {e}")
    
    return tables

@app.route("/convert/summary", methods=["POST"])
def generate_summary():
    """Extrai texto completo do PDF incluindo imagens"""
    pdf_path = None
    
    try:
        # Verificar se o arquivo foi enviado
        if "file" not in request.files:
            return jsonify({"error": "Nenhum arquivo PDF enviado"}), 400
        
        pdf_file = request.files["file"]
        
        if pdf_file.filename == '':
            return jsonify({"error": "Nenhum arquivo selecionado"}), 400
        
        # Salvar arquivo
        filename = secure_filename(pdf_file.filename)
        pdf_path = f"/tmp/{uuid.uuid4()}_{filename}"
        pdf_file.save(pdf_path)
        
        logging.info(f"Extraindo texto para resumo: {pdf_path}")
        
        # Extrair texto principal
        text = extract_text_from_pdf(pdf_path)
        
        # Extrair texto de imagens
        images_text = extract_images_from_pdf(pdf_path)
        
        # Combinar textos
        full_text = text + "\n\n" + images_text if images_text else text
        
        if not full_text.strip():
            return jsonify({
                "error": "Não foi possível extrair texto do PDF",
                "details": "O PDF pode estar protegido, corrompido ou não conter texto extraível"
            }), 422
        
        # Retornar texto extraído
        return jsonify({
            "text": full_text,
            "filename": filename,
            "extracted_at": datetime.now().isoformat(),
            "char_count": len(full_text),
            "has_images": len(images_text) > 0
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

@app.route("/convert/excel", methods=["POST"])
def convert_to_excel():
    """Conversão melhorada para Excel"""
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
        
        logging.info(f"Convertendo para Excel: {pdf_path}")
        
        # Extrair tabelas
        tables = extract_tables_from_pdf(pdf_path)
        
        # Extrair texto das imagens para incluir no Excel
        images_text = extract_images_from_pdf(pdf_path)
        
        if not tables and not images_text:
            return jsonify({
                "error": "Nenhuma tabela ou conteúdo extraível encontrado",
                "details": "O PDF não contém tabelas ou imagens com texto extraível"
            }), 422
        
        xlsx_path = f"/tmp/{uuid.uuid4()}.xlsx"
        
        with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
            sheet_count = 0
            
            # Adicionar tabelas
            for i, table in enumerate(tables):
                df = table.df.copy()
                df = df.dropna(how='all', axis=0)  # Remover linhas completamente vazias
                df = df.dropna(how='all', axis=1)  # Remover colunas completamente vazias
                
                # Limpar dados
                df = df.applymap(lambda x: str(x).strip() if pd.notna(x) else '')
                
                # Remover linhas que são apenas espaços
                df = df[df.apply(lambda x: x.str.strip().str.len().sum() > 0, axis=1)]
                
                if not df.empty:
                    sheet_name = f"Tabela_{i+1}"[:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    sheet_count += 1
                    logging.info(f"Tabela {i+1} salva: {len(df)} linhas x {len(df.columns)} colunas")
            
            # Adicionar texto de imagens se houver
            if images_text:
                lines = images_text.split('\n')
                image_data = []
                
                for line in lines:
                    if line.strip():
                        image_data.append([line.strip()])
                
                if image_data:
                    df_images = pd.DataFrame(image_data, columns=['Texto_Imagens'])
                    df_images.to_excel(writer, sheet_name='Texto_Imagens', index=False)
                    sheet_count += 1
            
            # Se não houver dados, criar uma planilha vazia com mensagem
            if sheet_count == 0:
                df_empty = pd.DataFrame([['Nenhuma tabela ou texto extraível encontrado']], 
                                      columns=['Resultado'])
                df_empty.to_excel(writer, sheet_name='Resultado', index=False)
        
        return send_file(
            xlsx_path,
            as_attachment=True,
            download_name=f"tabelas_{filename.replace('.pdf', '')}.xlsx",
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
    """Conversão melhorada para Word"""
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
        
        # Extrair texto principal
        text = extract_text_from_pdf(pdf_path)
        
        # Extrair texto de imagens
        images_text = extract_images_from_pdf(pdf_path)
        
        # Extrair tabelas
        tables = extract_tables_from_pdf(pdf_path)
        
        if not text.strip() and not tables:
            return jsonify({
                "error": "Nenhum conteúdo extraível encontrado",
                "details": "O PDF pode estar protegido, corrompido ou não conter texto extraível"
            }), 422
        
        doc = Document()
        doc.add_heading('Documento convertido de PDF', level=1)
        
        if text.strip():
            p = doc.add_paragraph(text)
        
        if images_text:
            doc.add_page_break()
            doc.add_heading("Texto extraído de imagens", level=2)
            p_img = doc.add_paragraph(images_text)
        
        # Adicionar tabelas ao Word
        if tables:
            doc.add_page_break()
            doc.add_heading("Tabelas extraídas", level=2)
            
            for table in tables:
                df = table.df.copy()
                df = df.dropna(how='all', axis=0)
                df = df.dropna(how='all', axis=1)
                
                if df.empty:
                    continue
                
                table_word = doc.add_table(rows=1, cols=len(df.columns))
                hdr_cells = table_word.rows[0].cells
                for i, col_name in enumerate(df.columns):
                    hdr_cells[i].text = str(col_name)
                
                for _, row in df.iterrows():
                    row_cells = table_word.add_row().cells
                    for i, cell_val in enumerate(row):
                        row_cells[i].text = str(cell_val)
                
                doc.add_paragraph()
        
        docx_path = f"/tmp/{uuid.uuid4()}.docx"
        doc.save(docx_path)
        
        return send_file(
            docx_path,
            as_attachment=True,
            download_name=f"{filename.replace('.pdf', '')}.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    
    except Exception as e:
        logging.error(f"Erro na conversão Word: {str(e)}")
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

@app.route("/convert/pptx", methods=["POST"])
def convert_to_pptx():
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
        
        # Extrair texto
        text = extract_text_from_pdf(pdf_path)
        
        # Extrair tabelas
        tables = extract_tables_from_pdf(pdf_path)
        
        prs = Presentation()
        blank_slide_layout = prs.slide_layouts[6]
        
        # Slide texto
        if text.strip():
            slide = prs.slides.add_slide(blank_slide_layout)
            shapes = slide.shapes
            txBox = shapes.add_textbox(PptxInches(1), PptxInches(1), PptxInches(8), PptxInches(5))
            tf = txBox.text_frame
            tf.word_wrap = True
            tf.text = text[:3000]  # Limite de caracteres por slide
            
            # Se texto longo, dividir em slides adicionais
            remaining_text = text[3000:]
            while remaining_text:
                slide = prs.slides.add_slide(blank_slide_layout)
                shapes = slide.shapes
                txBox = shapes.add_textbox(PptxInches(1), PptxInches(1), PptxInches(8), PptxInches(5))
                tf = txBox.text_frame
                tf.word_wrap = True
                tf.text = remaining_text[:3000]
                remaining_text = remaining_text[3000:]
        
        # Slides tabelas
        for table in tables:
            df = table.df.copy()
            df = df.dropna(how='all', axis=0)
            df = df.dropna(how='all', axis=1)
            
            if df.empty:
                continue
            
            slide = prs.slides.add_slide(blank_slide_layout)
            shapes = slide.shapes
            
            rows, cols = df.shape
            table_shape = shapes.add_table(rows+1, cols, PptxInches(0.5), PptxInches(0.5), PptxInches(9), PptxInches(5)).table
            
            # Cabeçalhos
            for col_idx, col_name in enumerate(df.columns):
                table_shape.cell(0, col_idx).text = str(col_name)
            
            # Dados
            for row_idx in range(rows):
                for col_idx in range(cols):
                    val = df.iat[row_idx, col_idx]
                    table_shape.cell(row_idx+1, col_idx).text = str(val)
        
        pptx_path = f"/tmp/{uuid.uuid4()}.pptx"
        prs.save(pptx_path)
        
        return send_file(
            pptx_path,
            as_attachment=True,
            download_name=f"{filename.replace('.pdf', '')}.pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    
    except Exception as e:
        logging.error(f"Erro na conversão PPTX: {str(e)}")
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

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)

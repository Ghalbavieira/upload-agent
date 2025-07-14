from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os

app = Flask(__name__)

# Configurar CORS para permitir Vercel
CORS(app, origins=[
    "https://upload-agent.vercel.app",
    "https://upload-agent-git-main-ghalba-vieiras-projects-f2e9b128.vercel.app",
    "http://localhost:3000",  # Para desenvolvimento
    "http://127.0.0.1:3000"   # Para desenvolvimento
])

@app.route('/convert', methods=['POST'])
def convert_pdf():
    if 'pdf' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400
    
    file = request.files['pdf']
    if file.filename == '':
        return jsonify({'error': 'Nenhum arquivo selecionado'}), 400
    
    if file and file.filename.lower().endswith('.pdf'):
        # Sua lógica de conversão aqui
        # Por enquanto, retorna um arquivo de exemplo
        return jsonify({'message': 'PDF processado com sucesso!'})
    
    return jsonify({'error': 'Arquivo deve ser PDF'}), 400

@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({'status': 'OK', 'message': 'Servidor funcionando!'})

if __name__ == '__main__':
    app.run(debug=True, port=int(os.environ.get('PORT', 5000)))
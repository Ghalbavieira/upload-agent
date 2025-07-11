 const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const uploadBtn = document.getElementById('uploadBtn');
const loading = document.getElementById('loading');
const progressBar = document.getElementById('progressBar');
const progressFill = document.getElementById('progressFill');
const result = document.getElementById('result');

const N8N_WEBHOOK_URL = 'https://ghalba.app.n8n.cloud/webhook/pdf-to-excel';

let selectedFile = null;


uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFileSelect(files[0]);
    }
});

// Evento de clique na área de upload
uploadArea.addEventListener('click', () => {
    fileInput.click();
});

// Evento de seleção de arquivo
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFileSelect(e.target.files[0]);
    }
});

function handleFileSelect(file) {
    if (file.type !== 'application/pdf') {
        showResult('Apenas arquivos PDF são aceitos.', true);
        return;
    }

    selectedFile = file;
    fileName.textContent = file.name;
    fileInfo.classList.add('show');
    uploadBtn.disabled = false;
    hideResult();
}

uploadBtn.addEventListener('click', async () => {
    if (!selectedFile) return;

    try {
        showLoading();
        simulateProgress();

        const formData = new FormData();
        formData.append('pdf', selectedFile);

        const response = await fetch(N8N_WEBHOOK_URL, {
            method: 'POST',
            body: formData
        });

        hideLoading();

        if (!response.ok) {
            throw new Error(`Erro na requisição: ${response.status}`);
        }

        const responseData = await response.json();
        
        if (responseData.downloadUrl) {
            showResult(
                `✅ PDF processado com sucesso!<br>
                <a href="${responseData.downloadUrl}" class="download-btn" download>
                    📥 Baixar Planilha
                </a>`,
                false
            );
        } else {
            showResult('✅ PDF processado com sucesso!', false);
        }

    } catch (error) {
        hideLoading();
        console.error('Erro ao processar PDF:', error);
        showResult(`❌ Erro ao processar o PDF: ${error.message}`, true);
    }
});

function showLoading() {
    loading.classList.add('show');
    progressBar.classList.add('show');
    uploadBtn.disabled = true;
}

function hideLoading() {
    loading.classList.remove('show');
    progressBar.classList.remove('show');
    uploadBtn.disabled = false;
    progressFill.style.width = '0%';
}

function showResult(message, isError) {
    result.innerHTML = message;
    result.classList.add('show');
    if (isError) {
        result.classList.add('error');
    } else {
        result.classList.remove('error');
    }
}

function hideResult() {
    result.classList.remove('show');
}

function simulateProgress() {
    let progress = 0;
    const interval = setInterval(() => {
        progress += Math.random() * 15;
        if (progress > 90) progress = 90;
        progressFill.style.width = `${progress}%`;
        
        if (progress >= 90) {
            clearInterval(interval);
        }
    }, 200);
}
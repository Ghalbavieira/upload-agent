let selectedFile = null;
        
        document.addEventListener('DOMContentLoaded', function() {
            const uploadArea = document.getElementById('uploadArea');
            const fileInput = document.getElementById('fileInput');
            const fileInfo = document.getElementById('fileInfo');
            const fileName = document.getElementById('fileName');
            const uploadBtn = document.getElementById('uploadBtn');
            const loading = document.getElementById('loading');
            const progressBar = document.getElementById('progressBar');
            const progressFill = document.getElementById('progressFill');
            const result = document.getElementById('result');

            // Função para mostrar loading
            function showLoading() {
                loading.classList.add('show');
                progressBar.classList.add('show');
                result.classList.remove('show');
            }

            // Função para esconder loading
            function hideLoading() {
                loading.classList.remove('show');
                progressBar.classList.remove('show');
            }

            // Função para simular progresso
            function simulateProgress() {
                let progress = 0;
                const interval = setInterval(() => {
                    progress += Math.random() * 15;
                    if (progress >= 90) {
                        clearInterval(interval);
                        progress = 90;
                    }
                    progressFill.style.width = progress + '%';
                }, 200);
            }

            // Função para mostrar resultado
            function showResult(message, isError = false) {
                result.innerHTML = message;
                result.classList.add('show');
                result.className = 'result show ' + (isError ? 'error' : 'success');
                progressFill.style.width = '100%';
            }

            // Clique na área de upload
            uploadArea.addEventListener('click', function() {
                fileInput.click();
            });

            // Quando um arquivo é selecionado
            fileInput.addEventListener('change', function(e) {
                const file = e.target.files[0];
                if (file && file.type === 'application/pdf') {
                    selectedFile = file;
                    fileName.textContent = file.name;
                    fileInfo.classList.add('show');
                    uploadBtn.disabled = false;
                } else if (file) {
                    alert('Por favor, selecione apenas arquivos PDF.');
                    fileInput.value = '';
                }
            });

            // Drag and drop events
            uploadArea.addEventListener('dragover', function(e) {
                e.preventDefault();
                uploadArea.classList.add('drag-over');
            });

            uploadArea.addEventListener('dragleave', function(e) {
                e.preventDefault();
                uploadArea.classList.remove('drag-over');
            });

            uploadArea.addEventListener('drop', function(e) {
                e.preventDefault();
                uploadArea.classList.remove('drag-over');
                
                const files = e.dataTransfer.files;
                if (files.length > 0) {
                    const file = files[0];
                    if (file.type === 'application/pdf') {
                        selectedFile = file;
                        fileName.textContent = file.name;
                        fileInfo.classList.add('show');
                        uploadBtn.disabled = false;
                        
                        // Simular seleção no input
                        const dataTransfer = new DataTransfer();
                        dataTransfer.items.add(file);
                        fileInput.files = dataTransfer.files;
                    } else {
                        alert('Por favor, selecione apenas arquivos PDF.');
                    }
                }
            });

            // Processar upload
            uploadBtn.addEventListener('click', async () => {
                if (!selectedFile) return;
                
                showLoading();
                simulateProgress();
                showResult('Enviando PDF para o servidor…');
                
                try {
                    const fd = new FormData();
                    fd.append('pdf', selectedFile);
                    const BACKEND_URL = 'https://upload-agent.onrender.com/convert';
                    
                    const res = await fetch(BACKEND_URL, {
                        method: 'POST',
                        body: fd
                    });
                    
                    if (!res.ok) throw new Error('Falha na conversão');
                    
                    const blob = await res.blob();
                    const downloadUrl = URL.createObjectURL(blob);
                    
                    hideLoading();
                    showResult(`
                        ✅ PDF convertido!<br>
                        <a href="${downloadUrl}" download="planilha.xlsx" class="download-btn" target="_blank">
                            📥 Baixar Planilha
                        </a>
                    `, false);
                    
                } catch (err) {
                    hideLoading();
                    showResult(`❌ ${err.message}`, true);
                }
            });
        });
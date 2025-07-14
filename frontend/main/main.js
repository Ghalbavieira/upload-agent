uploadBtn.addEventListener('click', async () => {
    if (!selectedFile) return;

    showLoading();
    simulateProgress();
    showResult('Enviando PDF para o servidor…');

    try {
        const fd = new FormData();
        fd.append('pdf', selectedFile);

        // Troque esta URL pela sua (ex: render, fly.io ou localhost)
        const BACKEND_URL = 'https://seu-backend.onrender.com/convert';

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

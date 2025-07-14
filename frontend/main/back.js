
const uploadArea  = document.getElementById('uploadArea');
const fileInput   = document.getElementById('fileInput');
const fileInfo    = document.getElementById('fileInfo');
const fileName    = document.getElementById('fileName');
const uploadBtn   = document.getElementById('uploadBtn');
const loading     = document.getElementById('loading');
const progressBar = document.getElementById('progressBar');
const progressFill= document.getElementById('progressFill');
const result      = document.getElementById('result');


const N8N_WEBHOOK_URL      = 'https://ghalba.app.n8n.cloud/webhook-test/pdf-to-excel';
const N8N_POLLING_ENDPOINT = N8N_WEBHOOK_URL;  // o mesmo endpoint (GET)


let selectedFile = null;


uploadArea.addEventListener('dragover', e => { e.preventDefault(); uploadArea.classList.add('dragover'); });
uploadArea.addEventListener('dragleave',   () => uploadArea.classList.remove('dragover'));
uploadArea.addEventListener('drop', e => {
  e.preventDefault(); uploadArea.classList.remove('dragover');
  if (e.dataTransfer.files.length) handleFileSelect(e.dataTransfer.files[0]);
});


uploadArea.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', e => { if (e.target.files.length) handleFileSelect(e.target.files[0]); });

function handleFileSelect(file) {
  if (file.type !== 'application/pdf') return showResult('Apenas PDF é aceito.', true);
  selectedFile = file;
  fileName.textContent = file.name;
  fileInfo.classList.add('show');
  uploadBtn.disabled = false;
  hideResult();
}

uploadBtn.addEventListener('click', async () => {
  if (!selectedFile) return;
  try {
    showLoading(); simulateProgress();

    const fd = new FormData();
    fd.append('pdf', selectedFile);

    const res  = await fetch(N8N_WEBHOOK_URL, { method: 'POST', body: fd });
    const json = await res.json();
    hideLoading();

    if (json.downloadUrl) {
      return showDownload(json.downloadUrl);
    }
    if (json.jobId) {
      showResult('📄 PDF recebido! Processando…', false);
      await pollForResult(json.jobId);
    } else {
      throw new Error('Resposta inesperada do servidor');
    }
  } catch (err) {
    hideLoading();
    showResult(`❌ ${err.message}`, true);
  }
});

async function pollForResult(jobId) {
  const maxTries = 60;          // 60 × 5 s = 5 min
  for (let n = 0; n < maxTries; n++) {
    await sleep(5000);
    const r   = await fetch(`${N8N_POLLING_ENDPOINT}?jobId=${encodeURIComponent(jobId)}`);
    const j   = await r.json();
    if (j.downloadUrl) return showDownload(j.downloadUrl);
    console.log(`Ainda processando… (${n + 1}/${maxTries})`);
  }
  showResult('⚠️ Tempo limite atingido. Tente novamente mais tarde.', true);
}

function showDownload(url) {
  showResult(`
    ✅ PDF convertido!<br>
    <a href="${url}" class="download-btn" download target="_blank" rel="noopener">
      📥 Baixar Planilha
    </a>
  `, false);
}
function showLoading()  { loading.classList.add('show'); progressBar.classList.add('show'); uploadBtn.disabled = true; }
function hideLoading()  { loading.classList.remove('show'); progressBar.classList.remove('show'); uploadBtn.disabled = false; progressFill.style.width='0%'; }
function showResult(msg, isErr){ result.innerHTML=msg; result.classList.add('show'); result.classList.toggle('error',isErr); }
function hideResult()   { result.classList.remove('show'); }
function sleep(ms)      { return new Promise(r => setTimeout(r, ms)); }
function simulateProgress(){
  let p=0; const int=setInterval(()=>{ p+=Math.random()*15; if(p>90)p=90; progressFill.style.width=`${p}%`; if(p>=90)clearInterval(int);},200);
}


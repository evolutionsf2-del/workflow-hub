// WorkFlow Hub Front-end Application

// Initialize Lucide icons
document.addEventListener('DOMContentLoaded', () => {
  if (window.lucide) {
    window.lucide.createIcons();
  }
  setupPdfHandlers();
  setupExcelHandlers();
});

// Format byte sizes into readable text
function formatBytes(bytes, decimals = 1) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

// Tab Switching
function switchTab(tab) {
  const tabPdf = document.getElementById('tab-pdf');
  const tabExcel = document.getElementById('tab-excel');
  const panelPdf = document.getElementById('panel-pdf');
  const panelExcel = document.getElementById('panel-excel');

  if (tab === 'pdf') {
    tabPdf.className = 'tab-btn flex-1 py-2.5 px-4 rounded-xl text-sm font-semibold flex items-center justify-center gap-2 transition-all duration-200 bg-white text-blue-600 shadow-sm';
    tabExcel.className = 'tab-btn flex-1 py-2.5 px-4 rounded-xl text-sm font-semibold flex items-center justify-center gap-2 transition-all duration-200 text-slate-600 hover:text-slate-900';
    panelPdf.classList.remove('hidden');
    panelExcel.classList.add('hidden');
  } else {
    tabExcel.className = 'tab-btn flex-1 py-2.5 px-4 rounded-xl text-sm font-semibold flex items-center justify-center gap-2 transition-all duration-200 bg-white text-emerald-600 shadow-sm';
    tabPdf.className = 'tab-btn flex-1 py-2.5 px-4 rounded-xl text-sm font-semibold flex items-center justify-center gap-2 transition-all duration-200 text-slate-600 hover:text-slate-900';
    panelExcel.classList.remove('hidden');
    panelPdf.classList.add('hidden');
  }
  if (window.lucide) window.lucide.createIcons();
}

// -----------------------------------------------------
// PDF to Word Logic
// -----------------------------------------------------
let selectedPdfFile = null;

function setupPdfHandlers() {
  const dropzone = document.getElementById('pdf-dropzone');
  const input = document.getElementById('pdf-input');

  dropzone.addEventListener('click', () => input.click());

  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handlePdfSelected(e.dataTransfer.files[0]);
    }
  });

  input.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handlePdfSelected(e.target.files[0]);
    }
  });
}

function handlePdfSelected(file) {
  if (!file.name.toLowerCase().endsWith('.pdf')) {
    showPdfError('กรุณาเลือกไฟล์ PDF เท่านั้น (.pdf)');
    return;
  }
  hidePdfError();
  selectedPdfFile = file;

  document.getElementById('pdf-filename').textContent = file.name;
  document.getElementById('pdf-filesize').textContent = formatBytes(file.size);
  document.getElementById('pdf-file-card').classList.remove('hidden');
  document.getElementById('pdf-file-card').classList.add('flex');
  document.getElementById('pdf-convert-btn').disabled = false;
  document.getElementById('pdf-download-container').classList.add('hidden');
}

function clearPdfFile() {
  selectedPdfFile = null;
  document.getElementById('pdf-input').value = '';
  document.getElementById('pdf-file-card').classList.add('hidden');
  document.getElementById('pdf-file-card').classList.remove('flex');
  document.getElementById('pdf-convert-btn').disabled = true;
  document.getElementById('pdf-download-container').classList.add('hidden');
  hidePdfError();
  hidePdfStatus();
}

function showPdfStatus(msg) {
  const statusEl = document.getElementById('pdf-status');
  document.getElementById('pdf-status-text').textContent = msg;
  statusEl.classList.remove('hidden');
}

function hidePdfStatus() {
  document.getElementById('pdf-status').classList.add('hidden');
}

function showPdfError(msg) {
  const errEl = document.getElementById('pdf-error');
  document.getElementById('pdf-error-text').textContent = msg;
  errEl.classList.remove('hidden');
}

function hidePdfError() {
  document.getElementById('pdf-error').classList.add('hidden');
}

async function startPdfConversion() {
  if (!selectedPdfFile) return;

  hidePdfError();
  showPdfStatus('กำลังประมวลผลและแปลงไฟล์เป็น Word (.docx)... อาจใช้เวลาสักครู่ตามขนาดเอกสาร');
  const btn = document.getElementById('pdf-convert-btn');
  btn.disabled = true;

  const formData = new FormData();
  formData.append('file', selectedPdfFile);

  try {
    const response = await fetch('/api/convert/pdf-to-word', {
      method: 'POST',
      body: formData,
    });

    if (!response.ok) {
      let errData = {};
      try { errData = await response.json(); } catch(e) {}
      throw new Error(errData.detail || 'เกิดข้อผิดพลาดในการแปลงไฟล์จากเซิร์ฟเวอร์');
    }

    const blob = await response.blob();
    const downloadUrl = window.URL.createObjectURL(blob);
    const originalName = selectedPdfFile.name.replace(/\.[^/.]+$/, "");
    const downloadName = `${originalName}.docx`;

    // Setup download button
    const downloadLink = document.getElementById('pdf-download-link');
    downloadLink.href = downloadUrl;
    downloadLink.download = downloadName;
    document.getElementById('pdf-download-container').classList.remove('hidden');

    hidePdfStatus();

    // Trigger auto-download
    const tempAnchor = document.createElement('a');
    tempAnchor.href = downloadUrl;
    tempAnchor.download = downloadName;
    document.body.appendChild(tempAnchor);
    tempAnchor.click();
    document.body.removeChild(tempAnchor);

  } catch (err) {
    hidePdfStatus();
    showPdfError(err.message || 'เกิดข้อผิดพลาดในการเชื่อมต่อเซิร์ฟเวอร์');
  } finally {
    btn.disabled = false;
    if (window.lucide) window.lucide.createIcons();
  }
}

// -----------------------------------------------------
// Excel to Word Logic
// -----------------------------------------------------
let selectedExcelFile = null;
let availableSheets = [];
let chosenSheets = new Set();

function setupExcelHandlers() {
  const dropzone = document.getElementById('excel-dropzone');
  const input = document.getElementById('excel-input');

  dropzone.addEventListener('click', () => input.click());

  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleExcelSelected(e.dataTransfer.files[0]);
    }
  });

  input.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleExcelSelected(e.target.files[0]);
    }
  });
}

async function handleExcelSelected(file) {
  const name = file.name.toLowerCase();
  if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) {
    showExcelError('กรุณาเลือกไฟล์ Excel เท่านั้น (.xlsx, .xls)');
    return;
  }
  hideExcelError();
  selectedExcelFile = file;

  document.getElementById('excel-filename').textContent = file.name;
  document.getElementById('excel-filesize').textContent = formatBytes(file.size);
  document.getElementById('excel-file-card').classList.remove('hidden');
  document.getElementById('excel-file-card').classList.add('flex');
  document.getElementById('excel-convert-btn').disabled = false;
  document.getElementById('excel-download-container').classList.add('hidden');

  // Inspect sheets
  await inspectExcelSheets(file);
}

async function inspectExcelSheets(file) {
  const sheetsContainer = document.getElementById('excel-sheets-container');
  const sheetsList = document.getElementById('excel-sheets-list');
  sheetsList.innerHTML = '<span class="text-xs text-slate-400">กำลังตรวจหาแผ่นงาน...</span>';
  sheetsContainer.classList.remove('hidden');

  const formData = new FormData();
  formData.append('file', file);

  try {
    const res = await fetch('/api/inspect/excel-sheets', {
      method: 'POST',
      body: formData,
    });
    if (res.ok) {
      const data = await res.json();
      availableSheets = data.sheets || [];
      chosenSheets = new Set(availableSheets); // select all by default
      renderSheetBadges();
    } else {
      sheetsContainer.classList.add('hidden');
    }
  } catch (e) {
    sheetsContainer.classList.add('hidden');
  }
}

function renderSheetBadges() {
  const sheetsList = document.getElementById('excel-sheets-list');
  sheetsList.innerHTML = '';

  availableSheets.forEach((sheet) => {
    const badge = document.createElement('button');
    badge.type = 'button';
    const isSelected = chosenSheets.has(sheet);
    badge.className = `sheet-badge px-3 py-1.5 rounded-xl text-xs font-medium border flex items-center gap-1.5 ${
      isSelected 
        ? 'bg-emerald-600 text-white border-emerald-600 shadow-sm' 
        : 'bg-white text-slate-600 border-slate-200 hover:border-slate-300'
    }`;
    badge.innerHTML = `<span class="w-1.5 h-1.5 rounded-full ${isSelected ? 'bg-white' : 'bg-slate-300'}"></span>${sheet}`;

    badge.onclick = () => {
      if (chosenSheets.has(sheet)) {
        if (chosenSheets.size > 1) {
          chosenSheets.delete(sheet);
        }
      } else {
        chosenSheets.add(sheet);
      }
      renderSheetBadges();
    };

    sheetsList.appendChild(badge);
  });
}

function clearExcelFile() {
  selectedExcelFile = null;
  availableSheets = [];
  chosenSheets.clear();
  document.getElementById('excel-input').value = '';
  document.getElementById('excel-file-card').classList.add('hidden');
  document.getElementById('excel-file-card').classList.remove('flex');
  document.getElementById('excel-sheets-container').classList.add('hidden');
  document.getElementById('excel-convert-btn').disabled = true;
  document.getElementById('excel-download-container').classList.add('hidden');
  hideExcelError();
  hideExcelStatus();
}

function showExcelStatus(msg) {
  const statusEl = document.getElementById('excel-status');
  document.getElementById('excel-status-text').textContent = msg;
  statusEl.classList.remove('hidden');
}

function hideExcelStatus() {
  document.getElementById('excel-status').classList.add('hidden');
}

function showExcelError(msg) {
  const errEl = document.getElementById('excel-error');
  document.getElementById('excel-error-text').textContent = msg;
  errEl.classList.remove('hidden');
}

function hideExcelError() {
  document.getElementById('excel-error').classList.add('hidden');
}

async function startExcelConversion() {
  if (!selectedExcelFile) return;

  hideExcelError();
  showExcelStatus('กำลังแปลงข้อมูลตาราง Excel เป็นเอกสาร Word (.docx)...');
  const btn = document.getElementById('excel-convert-btn');
  btn.disabled = true;

  const formData = new FormData();
  formData.append('file', selectedExcelFile);
  if (chosenSheets.size > 0) {
    formData.append('sheets', Array.from(chosenSheets).join(','));
  }

  try {
    const response = await fetch('/api/convert/excel-to-word', {
      method: 'POST',
      body: formData,
    });

    if (!response.ok) {
      let errData = {};
      try { errData = await response.json(); } catch(e) {}
      throw new Error(errData.detail || 'เกิดข้อผิดพลาดในการแปลงไฟล์จากเซิร์ฟเวอร์');
    }

    const blob = await response.blob();
    const downloadUrl = window.URL.createObjectURL(blob);
    const originalName = selectedExcelFile.name.replace(/\.[^/.]+$/, "");
    const downloadName = `${originalName}.docx`;

    const downloadLink = document.getElementById('excel-download-link');
    downloadLink.href = downloadUrl;
    downloadLink.download = downloadName;
    document.getElementById('excel-download-container').classList.remove('hidden');

    hideExcelStatus();

    // Trigger auto-download
    const tempAnchor = document.createElement('a');
    tempAnchor.href = downloadUrl;
    tempAnchor.download = downloadName;
    document.body.appendChild(tempAnchor);
    tempAnchor.click();
    document.body.removeChild(tempAnchor);

  } catch (err) {
    hideExcelStatus();
    showExcelError(err.message || 'เกิดข้อผิดพลาดในการเชื่อมต่อเซิร์ฟเวอร์');
  } finally {
    btn.disabled = false;
    if (window.lucide) window.lucide.createIcons();
  }
}

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
let pdfTotalPages = 0;
const MAX_SAFE_PAGES = 15;

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

async function handlePdfSelected(file) {
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
  document.getElementById('pdf-convert-btn').disabled = true; // Wait for inspection
  document.getElementById('pdf-download-container').classList.add('hidden');

  const pagesBadge = document.getElementById('pdf-pages-badge');
  pagesBadge.classList.remove('hidden');
  pagesBadge.textContent = 'กำลังตรวจนับหน้า...';

  // Inspect PDF to get total pages safely
  try {
    const formData = new FormData();
    formData.append('file', file);

    const resp = await fetch('/api/inspect/pdf-info', {
      method: 'POST',
      body: formData
    });

    if (!resp.ok) {
      let err = {};
      try { err = await resp.json(); } catch(e) {}
      throw new Error(err.detail || 'ไม่สามารถตรวจสอบข้อมูลไฟล์ PDF ได้');
    }

    const data = await resp.json();
    pdfTotalPages = data.total_pages || 1;
    pagesBadge.textContent = `${pdfTotalPages} หน้า`;
    document.getElementById('pdf-total-hint').textContent = `/ ทั้งหมด ${pdfTotalPages} หน้า`;

    const startInput = document.getElementById('pdf-start-page');
    const endInput = document.getElementById('pdf-end-page');
    startInput.max = pdfTotalPages;
    endInput.max = pdfTotalPages;

    const optionsCard = document.getElementById('pdf-options-card');
    optionsCard.classList.remove('hidden');

    const limitNotice = document.getElementById('pdf-limit-notice');
    const limitNoticeText = document.getElementById('pdf-limit-notice-text');
    const allDesc = document.getElementById('pdf-all-desc');

    if (pdfTotalPages > MAX_SAFE_PAGES) {
      // Over limit: recommend page range automatically
      document.getElementById('pdf-mode-range').checked = true;
      startInput.value = 1;
      endInput.value = Math.min(MAX_SAFE_PAGES, pdfTotalPages);
      togglePdfPageMode();

      limitNoticeText.textContent = `เอกสารนี้มีทั้งหมด ${pdfTotalPages} หน้า ซึ่งเกิน ${MAX_SAFE_PAGES} หน้า เพื่อป้องกันหน่วยความจำของเซิร์ฟเวอร์เต็ม ระบบจึงเลือกโหมด "กำหนดช่วงหน้า" ให้โดยอัตโนมัติ (แนะนำแปลงรอบแรกหน้า 1-${Math.min(MAX_SAFE_PAGES, pdfTotalPages)})`;
      limitNotice.classList.remove('hidden');
      allDesc.textContent = `เอกสารนี้มี ${pdfTotalPages} หน้า (เกินโควตา 15 หน้า หากเลือกแปลงทั้งหมด เซิร์ฟเวอร์จะปฏิเสธเพื่อป้องกันเว็บดับ)`;
      allDesc.classList.add('text-amber-600');
    } else {
      // Within safe limit: default to all pages
      document.getElementById('pdf-mode-all').checked = true;
      startInput.value = 1;
      endInput.value = pdfTotalPages;
      togglePdfPageMode();

      limitNotice.classList.add('hidden');
      allDesc.textContent = `แปลงเอกสารทั้งหมดตั้งแต่หน้าแรกจนจบ (เหมาะสำหรับเอกสารไม่เกิน ${MAX_SAFE_PAGES} หน้า)`;
      allDesc.classList.remove('text-amber-600');
    }

    document.getElementById('pdf-convert-btn').disabled = false;
  } catch (err) {
    pagesBadge.textContent = 'ไม่ทราบจำนวนหน้า';
    showPdfError(`ไม่สามารถตรวจนับหน้าเอกสารได้: ${err.message}`);
    document.getElementById('pdf-convert-btn').disabled = false;
  } finally {
    if (window.lucide) window.lucide.createIcons();
  }
}

function togglePdfPageMode() {
  const isRange = document.getElementById('pdf-mode-range').checked;
  const rangeInputs = document.getElementById('pdf-range-inputs');
  const limitNotice = document.getElementById('pdf-limit-notice');
  const limitNoticeText = document.getElementById('pdf-limit-notice-text');

  if (isRange) {
    rangeInputs.classList.remove('hidden');
    rangeInputs.classList.add('flex');
    if (pdfTotalPages > MAX_SAFE_PAGES) {
      limitNoticeText.textContent = `เอกสารมี ${pdfTotalPages} หน้า กรุณาเลือกช่วงหน้าไม่เกิน ${MAX_SAFE_PAGES} หน้าต่อครั้ง (เช่น หน้า 1-${MAX_SAFE_PAGES} และรอบถัดไปหน้า ${MAX_SAFE_PAGES + 1}-${Math.min(pdfTotalPages, 2 * MAX_SAFE_PAGES)})`;
      limitNotice.classList.remove('hidden');
    } else {
      limitNotice.classList.add('hidden');
    }
  } else {
    rangeInputs.classList.add('hidden');
    rangeInputs.classList.remove('flex');
    if (pdfTotalPages > MAX_SAFE_PAGES) {
      limitNoticeText.textContent = `⚠️ คำเตือน: เซิร์ฟเวอร์ฟรีจำกัดไม่เกิน ${MAX_SAFE_PAGES} หน้าต่อการแปลง หากท่านกดเริ่มแปลง ระบบจะปฏิเสธเพื่อป้องกันเซิร์ฟเวอร์รีสตาร์ท กรุณาเลือกโหมด "กำหนดช่วงหน้า"`;
      limitNotice.classList.remove('hidden');
    } else {
      limitNotice.classList.add('hidden');
    }
  }
  if (window.lucide) window.lucide.createIcons();
}

function clearPdfFile() {
  selectedPdfFile = null;
  pdfTotalPages = 0;
  document.getElementById('pdf-input').value = '';
  document.getElementById('pdf-file-card').classList.add('hidden');
  document.getElementById('pdf-file-card').classList.remove('flex');
  document.getElementById('pdf-options-card').classList.add('hidden');
  document.getElementById('pdf-convert-btn').disabled = true;
  document.getElementById('pdf-download-container').classList.add('hidden');
  document.getElementById('pdf-mode-all').checked = true;
  document.getElementById('pdf-range-inputs').classList.add('hidden');
  document.getElementById('pdf-limit-notice').classList.add('hidden');
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
  const isRange = document.getElementById('pdf-mode-range').checked;
  const formData = new FormData();
  formData.append('file', selectedPdfFile);

  if (!isRange) {
    if (pdfTotalPages > MAX_SAFE_PAGES) {
      showPdfError(
        `ไฟล์มีจำนวนหน้าทั้งหมด ${pdfTotalPages} หน้า ซึ่งเกินขีดจำกัดความปลอดภัยของเซิร์ฟเวอร์ฟรี (ไม่เกิน ${MAX_SAFE_PAGES} หน้า) กรุณาเลือกโหมด "กำหนดช่วงหน้า" เพื่อแบ่งแปลงทีละส่วน (เช่น หน้า 1-${MAX_SAFE_PAGES}) หรือเปิดรันโปรแกรมบนคอมพิวเตอร์ของคุณ (Localhost)`
      );
      document.getElementById('pdf-mode-range').checked = true;
      togglePdfPageMode();
      return;
    }
  } else {
    const startVal = parseInt(document.getElementById('pdf-start-page').value, 10);
    const endVal = parseInt(document.getElementById('pdf-end-page').value, 10);

    if (isNaN(startVal) || isNaN(endVal)) {
      showPdfError('กรุณาระบุเลขหน้าที่ต้องการแปลงให้ถูกต้อง');
      return;
    }
    if (startVal < 1) {
      showPdfError('หน้าเริ่มต้นต้องไม่น้อยกว่า 1');
      return;
    }
    if (pdfTotalPages > 0 && endVal > pdfTotalPages) {
      showPdfError(`หน้าสิ้นสุดต้องไม่เกินจำนวนหน้าทั้งหมด (${pdfTotalPages} หน้า)`);
      return;
    }
    if (startVal > endVal) {
      showPdfError(`หน้าเริ่มต้น (${startVal}) ต้องไม่มากกว่าหน้าสิ้นสุด (${endVal})`);
      return;
    }
    const pageCount = endVal - startVal + 1;
    if (pageCount > MAX_SAFE_PAGES) {
      showPdfError(
        `ช่วงหน้าที่เลือกมีจำนวน ${pageCount} หน้า ซึ่งเกินขีดจำกัดความปลอดภัยของเซิร์ฟเวอร์ฟรี (ไม่เกิน ${MAX_SAFE_PAGES} หน้าต่อครั้ง) กรุณาระบุช่วงหน้าไม่เกิน ${MAX_SAFE_PAGES} หน้า`
      );
      return;
    }

    formData.append('start_page', startVal);
    formData.append('end_page', endVal);
  }

  showPdfStatus('กำลังประมวลผลและแปลงไฟล์เป็น Word (.docx)... อาจใช้เวลาสักครู่ตามขนาดเอกสาร');
  const btn = document.getElementById('pdf-convert-btn');
  btn.disabled = true;

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

    // Get filename from response header or compute
    let downloadName = '';
    const disposition = response.headers.get('content-disposition');
    if (disposition && disposition.includes('filename=')) {
      const match = disposition.match(/filename\*?=['"]?(?:UTF-\d['"]*)?([^;\r\n"']*)['"]?/i);
      if (match && match[1]) {
        downloadName = decodeURIComponent(match[1]);
      }
    }
    if (!downloadName) {
      const originalName = selectedPdfFile.name.replace(/\.[^/.]+$/, "");
      if (isRange) {
        const startVal = document.getElementById('pdf-start-page').value;
        const endVal = document.getElementById('pdf-end-page').value;
        downloadName = `${originalName}_หน้า${startVal}-${endVal}.docx`;
      } else {
        downloadName = `${originalName}.docx`;
      }
    }

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

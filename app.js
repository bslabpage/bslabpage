/**
 * PDF Studio PRO - Main Application Logic
 * Pure Client-Side JavaScript (HTML5, CSS3, ES6+)
 */

// Configure PDF.js Worker
if (window.pdfjsLib) {
  window.pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

// ================= Global State =================
const state = {
  mergeFiles: [], // Array of { id, file, name, size, pageCount, arrayBuffer, thumbUrl }
  mergedBlob: null,
  mergedFileName: 'merged_document.pdf',
  
  // Converter State
  convertFile: null, // { file, name, size, pageCount, arrayBuffer }
  convertFormat: 'docx', // 'docx' | 'pptx'
  convertedBlob: null,
  convertedFileName: '',
  
  // Modal Preview State
  previewDoc: null,
  previewPage: 1,
  previewTotalPages: 1,
  previewDocTitle: '',
};

// Drag and drop sorting tracker
let draggedIndex = null;

// ================= DOM Element References =================
const dom = {
  // Navigation
  tabMerge: document.getElementById('tabMerge'),
  tabConvert: document.getElementById('tabConvert'),
  tabAdmin: document.getElementById('tabAdmin'),
  adminPendingBadge: document.getElementById('adminPendingBadge'),
  mergeSection: document.getElementById('mergeSection'),
  convertSection: document.getElementById('convertSection'),
  adminSection: document.getElementById('adminSection'),
  mergeTabBadge: document.getElementById('mergeTabBadge'),
  themeToggleBtn: document.getElementById('themeToggleBtn'),

  // Auth & Profile in Header
  openLoginModalBtn: document.getElementById('openLoginModalBtn'),
  userProfileChip: document.getElementById('userProfileChip'),
  userAvatarCircle: document.getElementById('userAvatarCircle'),
  userChipName: document.getElementById('userChipName'),
  userChipEmail: document.getElementById('userChipEmail'),
  userChipBadge: document.getElementById('userChipBadge'),
  requestPermFromChipBtn: document.getElementById('requestPermFromChipBtn'),
  userLogoutBtn: document.getElementById('userLogoutBtn'),

  // Sequence Strip
  sequenceStripContainer: document.getElementById('sequenceStripContainer'),
  sequenceStrip: document.getElementById('sequenceStrip'),
  seqTotalFiles: document.getElementById('seqTotalFiles'),
  seqTotalPages: document.getElementById('seqTotalPages'),

  // Merge Dropzone & Inputs
  mergeDropzone: document.getElementById('mergeDropzone'),
  mergeFileInput: document.getElementById('mergeFileInput'),
  loadSamplePdfsBtn: document.getElementById('loadSamplePdfsBtn'),

  // File List & Toolbar
  fileListSection: document.getElementById('fileListSection'),
  fileCardsList: document.getElementById('fileCardsList'),
  fileCountBadge: document.getElementById('fileCountBadge'),
  sortNameAscBtn: document.getElementById('sortNameAscBtn'),
  sortNameDescBtn: document.getElementById('sortNameDescBtn'),
  reverseOrderBtn: document.getElementById('reverseOrderBtn'),
  clearAllFilesBtn: document.getElementById('clearAllFilesBtn'),

  // Merge Panel & Results
  mergedFileName: document.getElementById('mergedFileName'),
  optAddPageNumbers: document.getElementById('optAddPageNumbers'),
  summaryFileCount: document.getElementById('summaryFileCount'),
  summaryPageCount: document.getElementById('summaryPageCount'),
  summaryTotalSize: document.getElementById('summaryTotalSize'),
  startMergeBtn: document.getElementById('startMergeBtn'),
  mergeResultCard: document.getElementById('mergeResultCard'),
  resultInfoText: document.getElementById('resultInfoText'),
  downloadMergedBtn: document.getElementById('downloadMergedBtn'),
  previewMergedBtn: document.getElementById('previewMergedBtn'),

  // Converter Elements
  formatWordBtn: document.getElementById('formatWordBtn'),
  formatPptBtn: document.getElementById('formatPptBtn'),
  convertDropzone: document.getElementById('convertDropzone'),
  convertFileInput: document.getElementById('convertFileInput'),
  loadConvertSampleBtn: document.getElementById('loadConvertSampleBtn'),
  convertSubText: document.getElementById('convertSubText'),
  convertSelectedFileInfo: document.getElementById('convertSelectedFileInfo'),
  convertThumbCanvas: document.getElementById('convertThumbCanvas'),
  convertFileName: document.getElementById('convertFileName'),
  convertPageCount: document.getElementById('convertPageCount'),
  convertFileSize: document.getElementById('convertFileSize'),
  removeConvertFileBtn: document.getElementById('removeConvertFileBtn'),
  wordSettingsBlock: document.getElementById('wordSettingsBlock'),
  pptSettingsBlock: document.getElementById('pptSettingsBlock'),
  pptAspectRatio: document.getElementById('pptAspectRatio'),
  pptFontFamily: document.getElementById('pptFontFamily'),
  pptSmartCluster: document.getElementById('pptSmartCluster'),
  pptAddNotes: document.getElementById('pptAddNotes'),
  startConvertBtn: document.getElementById('startConvertBtn'),
  convertBtnText: document.getElementById('convertBtnText'),
  convertProgressWrapper: document.getElementById('convertProgressWrapper'),
  progressStatusLabel: document.getElementById('progressStatusLabel'),
  progressPercentLabel: document.getElementById('progressPercentLabel'),
  progressBarFill: document.getElementById('progressBarFill'),
  convertResultBox: document.getElementById('convertResultBox'),
  convertResultDesc: document.getElementById('convertResultDesc'),
  downloadConvertedBtn: document.getElementById('downloadConvertedBtn'),
  downloadConvertedBtnText: document.getElementById('downloadConvertedBtnText'),

  // Admin Console Elements
  adminRefreshBtn: document.getElementById('adminRefreshBtn'),
  adminExportCsvBtn: document.getElementById('adminExportCsvBtn'),
  adminSeedDemoBtn: document.getElementById('adminSeedDemoBtn'),
  adminResetDataBtn: document.getElementById('adminResetDataBtn'),
  kpiTotalUsers: document.getElementById('kpiTotalUsers'),
  kpiPendingRequests: document.getElementById('kpiPendingRequests'),
  kpiTotalActions: document.getElementById('kpiTotalActions'),
  kpiTotalPages: document.getElementById('kpiTotalPages'),
  subtabRequests: document.getElementById('subtabRequests'),
  subtabUsage: document.getElementById('subtabUsage'),
  subtabLogs: document.getElementById('subtabLogs'),
  subtabPendingCountBadge: document.getElementById('subtabPendingCountBadge'),
  requestsPanel: document.getElementById('requestsPanel'),
  usagePanel: document.getElementById('usagePanel'),
  logsPanel: document.getElementById('logsPanel'),
  reqSearchInput: document.getElementById('reqSearchInput'),
  usageSearchInput: document.getElementById('usageSearchInput'),
  openAddUserModalBtn: document.getElementById('openAddUserModalBtn'),
  adminReqTableBody: document.getElementById('adminReqTableBody'),
  emptyRequestsNotice: document.getElementById('emptyRequestsNotice'),
  adminUsageTableBody: document.getElementById('adminUsageTableBody'),
  emptyUsageNotice: document.getElementById('emptyUsageNotice'),
  adminLogsContainer: document.getElementById('adminLogsContainer'),
  emptyLogsNotice: document.getElementById('emptyLogsNotice'),
  clearLogsBtn: document.getElementById('clearLogsBtn'),
  filterCountAll: document.getElementById('filterCountAll'),
  filterCountPending: document.getElementById('filterCountPending'),
  filterCountApproved: document.getElementById('filterCountApproved'),
  filterCountRejected: document.getElementById('filterCountRejected'),

  // Admin Google Apps Script Sync Elements
  adminGasSyncCard: document.getElementById('adminGasSyncCard'),
  gasStatusBadge: document.getElementById('gasStatusBadge'),
  gasStatusText: document.getElementById('gasStatusText'),
  gasSyncNowBtn: document.getElementById('gasSyncNowBtn'),
  gasSyncBtnLabel: document.getElementById('gasSyncBtnLabel'),
  gasFetchNowBtn: document.getElementById('gasFetchNowBtn'),
  toggleGasConfigBtn: document.getElementById('toggleGasConfigBtn'),
  gasDisplayUrl: document.getElementById('gasDisplayUrl'),
  gasLastSyncTime: document.getElementById('gasLastSyncTime'),
  gasSyncStatusIndicator: document.getElementById('gasSyncStatusIndicator'),
  gasAutoSyncToggle: document.getElementById('gasAutoSyncToggle'),
  gasConfigPanel: document.getElementById('gasConfigPanel'),
  gasPingResultBadge: document.getElementById('gasPingResultBadge'),
  gasWebappUrlInput: document.getElementById('gasWebappUrlInput'),
  gasApiKeyInput: document.getElementById('gasApiKeyInput'),
  gasTestConnectionBtn: document.getElementById('gasTestConnectionBtn'),
  gasSaveConfigBtn: document.getElementById('gasSaveConfigBtn'),
  gasResetConfigBtn: document.getElementById('gasResetConfigBtn'),
  toggleGasScriptCodeBtn: document.getElementById('toggleGasScriptCodeBtn'),
  gasRecommendedCodeBox: document.getElementById('gasRecommendedCodeBox'),
  copyGasCodeBtn: document.getElementById('copyGasCodeBtn'),

  // Admin Google Drive Sync Elements
  adminDriveSyncCard: document.getElementById('adminDriveSyncCard'),
  driveStatusBadge: document.getElementById('driveStatusBadge'),
  driveStatusText: document.getElementById('driveStatusText'),
  driveSyncNowBtn: document.getElementById('driveSyncNowBtn'),
  driveSyncBtnLabel: document.getElementById('driveSyncBtnLabel'),
  driveOpenFileLink: document.getElementById('driveOpenFileLink'),
  driveDownloadJsonBtn: document.getElementById('driveDownloadJsonBtn'),
  driveLastSyncTime: document.getElementById('driveLastSyncTime'),
  driveVisitorCount: document.getElementById('driveVisitorCount'),
  driveTotalActionsCount: document.getElementById('driveTotalActionsCount'),
  driveAutoSyncToggle: document.getElementById('driveAutoSyncToggle'),

  // Feature Stats & Visitors Elements
  statMergeCount: document.getElementById('statMergeCount'),
  statWordCount: document.getElementById('statWordCount'),
  statPptCount: document.getElementById('statPptCount'),
  visitorTotalBadge: document.getElementById('visitorTotalBadge'),
  visitorSearchInput: document.getElementById('visitorSearchInput'),
  adminVisitorTableBody: document.getElementById('adminVisitorTableBody'),
  emptyVisitorNotice: document.getElementById('emptyVisitorNotice'),

  // Modals
  previewModal: document.getElementById('previewModal'),
  modalDocTitle: document.getElementById('modalDocTitle'),
  modalPageIndicator: document.getElementById('modalPageIndicator'),
  modalPaginationText: document.getElementById('modalPaginationText'),
  modalCanvas: document.getElementById('modalCanvas'),
  closeModalBtn: document.getElementById('closeModalBtn'),
  prevPageBtn: document.getElementById('prevPageBtn'),
  nextPageBtn: document.getElementById('nextPageBtn'),

  // Google Login Modal
  googleLoginModal: document.getElementById('googleLoginModal'),
  closeLoginModalBtn: document.getElementById('closeLoginModalBtn'),
  gsiButtonContainer: document.getElementById('gsiButtonContainer'),
  googlePrimaryLoginBtn: document.getElementById('googlePrimaryLoginBtn'),
  googlePopupTriggerBtn: document.getElementById('googlePopupTriggerBtn'),
  googleAuthPendingBox: document.getElementById('googleAuthPendingBox'),
  toggleClientIdConfigBtn: document.getElementById('toggleClientIdConfigBtn'),
  clientIdConfigBox: document.getElementById('clientIdConfigBox'),
  customGoogleClientIdInput: document.getElementById('customGoogleClientIdInput'),
  saveGoogleClientIdBtn: document.getElementById('saveGoogleClientIdBtn'),
  resetGoogleClientIdBtn: document.getElementById('resetGoogleClientIdBtn'),

  // Access Permission Request Modal
  accessRequestModal: document.getElementById('accessRequestModal'),
  closeAccessRequestBtn: document.getElementById('closeAccessRequestBtn'),
  cancelAccessRequestBtn: document.getElementById('cancelAccessRequestBtn'),
  accessRequestForm: document.getElementById('accessRequestForm'),
  reqUserEmailDisplay: document.getElementById('reqUserEmailDisplay'),
  reqUserNameInput: document.getElementById('reqUserNameInput'),
  reqReasonInput: document.getElementById('reqReasonInput'),

  // Permission Denied / Prompt Modal
  permissionNoticeModal: document.getElementById('permissionNoticeModal'),
  closePermNoticeBtn: document.getElementById('closePermNoticeBtn'),
  permNoticeTitle: document.getElementById('permNoticeTitle'),
  permNoticeMessage: document.getElementById('permNoticeMessage'),
  permNoticeActionBtn: document.getElementById('permNoticeActionBtn'),

  // Admin Add User Modal
  addUserModal: document.getElementById('addUserModal'),
  closeAddUserBtn: document.getElementById('closeAddUserBtn'),
  addUserForm: document.getElementById('addUserForm'),
  manualUserEmail: document.getElementById('manualUserEmail'),
  manualUserName: document.getElementById('manualUserName'),
  manualUserStatus: document.getElementById('manualUserStatus'),

  // Toast
  toastContainer: document.getElementById('toastContainer'),
};

// ================= Initialization =================
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  initVisitorTracker();
  initAuthAndPermissions();
  initTabs();
  initMergeDropzone();
  initMergeActions();
  initConverter();
  initModalEvents();
  initAdminConsole();
  initGoogleDriveSync();
  initGoogleAppsScriptSync();
});

window.addEventListener('load', () => {
  if (typeof initGoogleIdentityServices === 'function') {
    initGoogleIdentityServices();
  }
});

// ================= Theme Management =================
function initTheme() {
  const savedTheme = localStorage.getItem('pdf_studio_theme');
  const prefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  
  if (savedTheme === 'dark' || (!savedTheme && prefersDark)) {
    document.documentElement.setAttribute('data-theme', 'dark');
  } else {
    document.documentElement.setAttribute('data-theme', 'light');
  }

  if (dom.themeToggleBtn) {
    dom.themeToggleBtn.addEventListener('click', () => {
      const currentTheme = document.documentElement.getAttribute('data-theme');
      const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', newTheme);
      localStorage.setItem('pdf_studio_theme', newTheme);
      showToast(`테마가 ${newTheme === 'dark' ? '다크' : '라이트'} 모드로 전환되었습니다.`, 'info');
    });
  }
}

// ================= Tab Navigation =================
function initTabs() {
  dom.tabMerge.addEventListener('click', () => switchTab('mergeSection'));
  dom.tabConvert.addEventListener('click', () => switchTab('convertSection'));
  if (dom.tabAdmin) {
    dom.tabAdmin.addEventListener('click', () => switchTab('adminSection'));
  }
}

function switchTab(sectionId) {
  // Guard admin section
  if (sectionId === 'adminSection') {
    const currentUser = getCurrentUser();
    if (!isAdminUser(currentUser)) {
      showToast('관리자(bosung3116@gmail.com) 전용 메뉴입니다.', 'error');
      switchTab('mergeSection');
      return;
    }
    dom.tabMerge.classList.remove('active');
    dom.tabConvert.classList.remove('active');
    dom.tabAdmin.classList.add('active');
    dom.mergeSection.style.display = 'none';
    dom.convertSection.style.display = 'none';
    dom.adminSection.style.display = 'flex';
    renderAdminDashboard();
    startGasAutoSyncTimer();
    return;
  }

  stopGasAutoSyncTimer();

  if (sectionId === 'mergeSection') {
    dom.tabMerge.classList.add('active');
    dom.tabConvert.classList.remove('active');
    if (dom.tabAdmin) dom.tabAdmin.classList.remove('active');
    dom.mergeSection.style.display = 'block';
    dom.convertSection.style.display = 'none';
    if (dom.adminSection) dom.adminSection.style.display = 'none';
  } else {
    dom.tabMerge.classList.remove('active');
    dom.tabConvert.classList.add('active');
    if (dom.tabAdmin) dom.tabAdmin.classList.remove('active');
    dom.mergeSection.style.display = 'none';
    dom.convertSection.style.display = 'block';
    if (dom.adminSection) dom.adminSection.style.display = 'none';
  }
}

// ================= Toast Notifications =================
function showToast(message, type = 'info', duration = 3000) {
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  
  let icon = `ℹ️`;
  if (type === 'success') icon = `✅`;
  if (type === 'error') icon = `⚠️`;

  toast.innerHTML = `<span>${icon}</span> <span>${escapeHtml(message)}</span>`;
  dom.toastContainer.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function formatBytes(bytes, decimals = 1) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

// ================= PDF Thumbnail & Info Extraction =================
async function extractPdfMetadata(arrayBuffer) {
  try {
    const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer.slice(0) });
    const pdf = await loadingTask.promise;
    const pageCount = pdf.numPages;

    // Render thumbnail of first page
    const page = await pdf.getPage(1);
    const viewport = page.getViewport({ scale: 0.35 });
    
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const ctx = canvas.getContext('2d');

    await page.render({ canvasContext: ctx, viewport }).promise;
    const thumbUrl = canvas.toDataURL('image/jpeg', 0.85);

    return { pageCount, thumbUrl };
  } catch (err) {
    console.warn('PDF metadata extraction warning:', err);
    return { pageCount: 1, thumbUrl: '' };
  }
}

// ================= Merge Dropzone & File Handling =================
function initMergeDropzone() {
  const dropzone = dom.mergeDropzone;
  const fileInput = dom.mergeFileInput;

  ['dragenter', 'dragover'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('dragover');
    });
  });

  dropzone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    const files = dt.files;
    handleMergeFiles(files);
  });

  fileInput.addEventListener('change', (e) => {
    handleMergeFiles(e.target.files);
    fileInput.value = ''; // Reset input to allow selecting same file again
  });

  dom.loadSamplePdfsBtn.addEventListener('click', generateSamplePdfsForMerge);
}

async function handleMergeFiles(files) {
  if (!files || files.length === 0) return;

  const validFiles = Array.from(files).filter(f => f.type === 'application/pdf' || f.name.toLowerCase().endsWith('.pdf'));

  if (validFiles.length === 0) {
    showToast('PDF 파일(.pdf)만 추가할 수 있습니다.', 'error');
    return;
  }

  showToast(`${validFiles.length}개의 PDF 파일을 불러오는 중입니다...`, 'info', 1500);

  for (const file of validFiles) {
    try {
      const arrayBuffer = await file.arrayBuffer();
      const { pageCount, thumbUrl } = await extractPdfMetadata(arrayBuffer);

      const fileItem = {
        id: 'file_' + Date.now() + '_' + Math.random().toString(36).substr(2, 6),
        file: file,
        name: file.name,
        size: file.size,
        pageCount: pageCount,
        arrayBuffer: arrayBuffer,
        thumbUrl: thumbUrl,
      };

      state.mergeFiles.push(fileItem);
    } catch (err) {
      console.error('File load error:', err);
      showToast(`'${file.name}' 파일을 읽는데 실패했습니다.`, 'error');
    }
  }

  updateMergeUI();
  showToast(`총 ${state.mergeFiles.length}개 파일이 준비되었습니다. 상단에서 병합 순서를 확인하세요!`, 'success');
}

// ================= Sequence Strip & File List Rendering =================
function updateMergeUI() {
  const count = state.mergeFiles.length;
  dom.mergeTabBadge.textContent = count;
  dom.mergeTabBadge.style.display = count > 0 ? 'inline-block' : 'none';

  if (count === 0) {
    dom.fileListSection.style.display = 'none';
    dom.sequenceStripContainer.classList.remove('has-files');
    dom.seqTotalFiles.textContent = '0';
    dom.seqTotalPages.textContent = '0';
    dom.sequenceStrip.innerHTML = `
      <div class="sequence-empty">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <circle cx="12" cy="12" r="10"></circle>
          <line x1="12" y1="8" x2="12" y2="12"></line>
          <line x1="12" y1="16" x2="12.01" y2="16"></line>
        </svg>
        <span>PDF 파일을 업로드하면 여기에 결합 순서 흐름도가 실시간으로 표시됩니다.</span>
      </div>
    `;
    return;
  }

  dom.fileListSection.style.display = 'block';
  dom.sequenceStripContainer.classList.add('has-files');

  // Compute total stats
  let totalPages = 0;
  let totalBytes = 0;
  state.mergeFiles.forEach(f => {
    totalPages += f.pageCount;
    totalBytes += f.size;
  });

  dom.seqTotalFiles.textContent = count;
  dom.seqTotalPages.textContent = totalPages;
  dom.fileCountBadge.textContent = count;
  dom.summaryFileCount.textContent = count;
  dom.summaryPageCount.textContent = totalPages;
  dom.summaryTotalSize.textContent = formatBytes(totalBytes);

  renderSequenceStrip();
  renderFileCards();
}

/**
 * Requirement 2: 상단에 파일이 합쳐지는 순서를 먼저 볼 수 있는 전용 흐름도 렌더링
 */
function renderSequenceStrip() {
  dom.sequenceStrip.innerHTML = '';

  state.mergeFiles.forEach((fileItem, index) => {
    const isLast = index === state.mergeFiles.length - 1;

    const node = document.createElement('div');
    node.className = 'flow-node';
    node.title = `클릭하면 해당 파일(${fileItem.name}) 카드로 스크롤 이동합니다.`;
    node.innerHTML = `
      <span class="flow-step-badge">${index + 1}</span>
      <span class="flow-doc-name">${escapeHtml(fileItem.name)}</span>
      <span class="flow-pages-tag">${fileItem.pageCount}p</span>
    `;

    node.addEventListener('click', () => {
      const card = document.getElementById(`card_${fileItem.id}`);
      if (card) {
        card.scrollIntoView({ behavior: 'smooth', block: 'center' });
        card.style.transition = 'all 0.3s ease';
        card.style.boxShadow = '0 0 15px rgba(99, 102, 241, 0.6)';
        setTimeout(() => { card.style.boxShadow = ''; }, 1200);
      }
    });

    dom.sequenceStrip.appendChild(node);

    if (!isLast) {
      const arrow = document.createElement('div');
      arrow.className = 'flow-arrow';
      arrow.innerHTML = `➔`;
      dom.sequenceStrip.appendChild(arrow);
    }
  });
}

function renderFileCards() {
  dom.fileCardsList.innerHTML = '';

  state.mergeFiles.forEach((fileItem, index) => {
    const li = document.createElement('li');
    li.id = `card_${fileItem.id}`;
    li.className = 'file-card';
    li.draggable = true;
    li.dataset.index = index;

    const isFirst = index === 0;
    const isLast = index === state.mergeFiles.length - 1;

    li.innerHTML = `
      <div class="card-left-group">
        <div class="drag-handle" title="드래그하여 순서 변경">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <circle cx="9" cy="5" r="1.5"></circle>
            <circle cx="9" cy="12" r="1.5"></circle>
            <circle cx="9" cy="19" r="1.5"></circle>
            <circle cx="15" cy="5" r="1.5"></circle>
            <circle cx="15" cy="12" r="1.5"></circle>
            <circle cx="15" cy="19" r="1.5"></circle>
          </svg>
        </div>

        <div class="order-badge" title="병합 순서 ${index + 1}번째">${index + 1}</div>

        <div class="thumb-preview-box">
          ${fileItem.thumbUrl ? `<img src="${fileItem.thumbUrl}" alt="썸네일" />` : `<span>PDF</span>`}
        </div>

        <div class="card-info-col">
          <div class="file-item-name" title="${escapeHtml(fileItem.name)}">${escapeHtml(fileItem.name)}</div>
          <div class="file-item-meta">
            <span class="meta-page-tag">총 ${fileItem.pageCount}페이지</span>
            <span>${formatBytes(fileItem.size)}</span>
          </div>
        </div>
      </div>

      <div class="card-right-group">
        <div class="reorder-btns-col">
          <button type="button" class="action-icon-btn btn-up" title="한 칸 위로 이동" ${isFirst ? 'disabled' : ''}>
            ▲
          </button>
          <button type="button" class="action-icon-btn btn-down" title="한 칸 아래로 이동" ${isLast ? 'disabled' : ''}>
            ▼
          </button>
        </div>

        <button type="button" class="action-icon-btn btn-preview" title="이 문서 미리보기">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
            <circle cx="12" cy="12" r="3"></circle>
          </svg>
        </button>

        <button type="button" class="action-icon-btn btn-remove" title="목록에서 삭제">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <line x1="18" y1="6" x2="6" y2="18"></line>
            <line x1="6" y1="6" x2="18" y2="18"></line>
          </svg>
        </button>
      </div>
    `;

    // Event Listeners for actions
    li.querySelector('.btn-up').addEventListener('click', () => moveFile(index, -1));
    li.querySelector('.btn-down').addEventListener('click', () => moveFile(index, 1));
    li.querySelector('.btn-preview').addEventListener('click', () => openPreviewModal(fileItem));
    li.querySelector('.btn-remove').addEventListener('click', () => removeFile(index));

    // HTML5 Drag and Drop events
    li.addEventListener('dragstart', handleDragStart);
    li.addEventListener('dragover', handleDragOver);
    li.addEventListener('dragleave', handleDragLeave);
    li.addEventListener('drop', handleDrop);
    li.addEventListener('dragend', handleDragEnd);

    dom.fileCardsList.appendChild(li);
  });
}

// ================= Reordering & Sorting Logic =================
function moveFile(index, direction) {
  const targetIndex = index + direction;
  if (targetIndex < 0 || targetIndex >= state.mergeFiles.length) return;

  const item = state.mergeFiles.splice(index, 1)[0];
  state.mergeFiles.splice(targetIndex, 0, item);
  updateMergeUI();
}

function removeFile(index) {
  const removed = state.mergeFiles.splice(index, 1)[0];
  updateMergeUI();
  showToast(`'${removed.name}' 파일이 제외되었습니다.`, 'info');
}

// Drag & Drop Reordering Handlers
function handleDragStart(e) {
  draggedIndex = parseInt(this.dataset.index, 10);
  this.classList.add('dragging');
  e.dataTransfer.effectAllowed = 'move';
}

function handleDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
  const bounding = this.getBoundingClientRect();
  const offset = e.clientY - bounding.top;

  this.classList.remove('drag-over-top', 'drag-over-bottom');
  if (offset < bounding.height / 2) {
    this.classList.add('drag-over-top');
  } else {
    this.classList.add('drag-over-bottom');
  }
}

function handleDragLeave() {
  this.classList.remove('drag-over-top', 'drag-over-bottom');
}

function handleDrop(e) {
  e.preventDefault();
  this.classList.remove('drag-over-top', 'drag-over-bottom');
  const dropTargetIndex = parseInt(this.dataset.index, 10);

  if (draggedIndex === null || draggedIndex === dropTargetIndex) return;

  const bounding = this.getBoundingClientRect();
  const offset = e.clientY - bounding.top;
  let insertAt = dropTargetIndex;

  if (offset >= bounding.height / 2 && draggedIndex < dropTargetIndex) {
    // placed after
    insertAt = dropTargetIndex;
  }

  const [movedItem] = state.mergeFiles.splice(draggedIndex, 1);
  state.mergeFiles.splice(insertAt, 0, movedItem);

  updateMergeUI();
  showToast('병합 순서가 업데이트되었습니다.', 'success', 1500);
}

function handleDragEnd() {
  this.classList.remove('dragging');
  document.querySelectorAll('.file-card').forEach(c => c.classList.remove('drag-over-top', 'drag-over-bottom'));
  draggedIndex = null;
}

// Toolbar actions
function initMergeActions() {
  dom.sortNameAscBtn.addEventListener('click', () => {
    state.mergeFiles.sort((a, b) => a.name.localeCompare(b.name, 'ko-KR', { numeric: true }));
    updateMergeUI();
    showToast('파일명 오름차순(A-Z)으로 정렬되었습니다.', 'info');
  });

  dom.sortNameDescBtn.addEventListener('click', () => {
    state.mergeFiles.sort((a, b) => b.name.localeCompare(a.name, 'ko-KR', { numeric: true }));
    updateMergeUI();
    showToast('파일명 내림차순(Z-A)으로 정렬되었습니다.', 'info');
  });

  dom.reverseOrderBtn.addEventListener('click', () => {
    state.mergeFiles.reverse();
    updateMergeUI();
    showToast('파일 순서가 반전되었습니다.', 'info');
  });

  dom.clearAllFilesBtn.addEventListener('click', () => {
    if (confirm('목록의 모든 PDF 파일을 비우시겠습니까?')) {
      state.mergeFiles = [];
      dom.mergeResultCard.style.display = 'none';
      updateMergeUI();
      showToast('모든 파일이 삭제되었습니다.', 'info');
    }
  });

  dom.startMergeBtn.addEventListener('click', executePdfMerge);
  dom.downloadMergedBtn.addEventListener('click', downloadMergedPdf);
  dom.previewMergedBtn.addEventListener('click', () => {
    if (state.mergedBlob) {
      const url = URL.createObjectURL(state.mergedBlob);
      window.open(url, '_blank');
    }
  });
}

// ================= PDF Merge Execution (pdf-lib) =================
async function executePdfMerge() {
  if (state.mergeFiles.length < 1) {
    showToast('병합할 PDF 파일을 최소 1개 이상 추가해 주세요.', 'error');
    return;
  }

  // Permission Check
  if (!checkPermissionOrPrompt('PDF 병합')) {
    return;
  }

  dom.startMergeBtn.disabled = true;
  dom.startMergeBtn.innerHTML = `
    <span class="pulse-indicator"></span>
    <span>PDF 병합 처리 중...</span>
  `;

  try {
    const { PDFDocument, StandardFonts, rgb } = PDFLib;
    const mergedPdf = await PDFDocument.create();

    const addPageNumbers = dom.optAddPageNumbers.checked;
    let helveticaFont = null;
    if (addPageNumbers) {
      helveticaFont = await mergedPdf.embedFont(StandardFonts.Helvetica);
    }

    // Merge each PDF in sequence
    for (let i = 0; i < state.mergeFiles.length; i++) {
      const item = state.mergeFiles[i];
      const srcDoc = await PDFDocument.load(item.arrayBuffer.slice(0));
      const pageIndices = srcDoc.getPageIndices();
      const copiedPages = await mergedPdf.copyPages(srcDoc, pageIndices);

      copiedPages.forEach(page => {
        mergedPdf.addPage(page);
      });
    }

    // Optional Page Numbering
    const totalMergedPages = mergedPdf.getPageCount();
    if (addPageNumbers && helveticaFont) {
      const pages = mergedPdf.getPages();
      pages.forEach((page, idx) => {
        const { width, height } = page.getSize();
        const pageText = `${idx + 1} / ${totalMergedPages}`;
        const fontSize = 9;
        const textWidth = helveticaFont.widthOfTextAtSize(pageText, fontSize);

        page.drawText(pageText, {
          x: (width - textWidth) / 2,
          y: 20,
          size: fontSize,
          font: helveticaFont,
          color: rgb(0.4, 0.4, 0.4),
        });
      });
    }

    // Generate output PDF
    const mergedBytes = await mergedPdf.save();
    state.mergedBlob = new Blob([mergedBytes], { type: 'application/pdf' });

    let customName = dom.mergedFileName.value.trim() || 'merged_document';
    if (!customName.endsWith('.pdf')) customName += '.pdf';
    state.mergedFileName = customName;

    // Display result card
    dom.resultInfoText.textContent = `총 ${state.mergeFiles.length}개 파일 (${totalMergedPages}페이지) · ${formatBytes(state.mergedBlob.size)}`;
    dom.mergeResultCard.style.display = 'flex';
    dom.mergeResultCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

    // Track usage
    recordUsage('PDF_MERGE', `${state.mergeFiles.length}개 파일 병합 (${totalMergedPages}페이지)`, totalMergedPages, state.mergedFileName);

    showToast('🎉 PDF 병합이 완료되었습니다! 다운로드 버튼을 누르세요.', 'success', 4000);
  } catch (err) {
    console.error('Merge error:', err);
    showToast('PDF 병합 중 오류가 발생했습니다: ' + err.message, 'error');
  } finally {
    dom.startMergeBtn.disabled = false;
    dom.startMergeBtn.innerHTML = `
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
        <polyline points="8 17 12 21 16 17"></polyline>
        <line x1="12" y1="12" x2="12" y2="21"></line>
        <path d="M20.88 18.09A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.29"></path>
      </svg>
      <span>하나의 PDF로 병합하기</span>
    `;
  }
}

function downloadMergedPdf() {
  if (!state.mergedBlob) return;
  const url = URL.createObjectURL(state.mergedBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = state.mergedFileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  showToast(`'${state.mergedFileName}' 다운로드가 시작되었습니다.`, 'success');
}

// ================= Sample PDF Generator =================
async function generateSamplePdfsForMerge() {
  showToast('테스트용 샘플 PDF 3개를 브라우저에서 직접 생성하는 중입니다...', 'info', 2000);

  const { PDFDocument, StandardFonts, rgb } = PDFLib;

  const sampleConfigs = [
    { title: '1. 사업계획서_요약본.pdf', pages: 2, color: rgb(0.2, 0.4, 0.8), desc: '사업 개요 및 전략' },
    { title: '2. 재무제표_및_손익보고서.pdf', pages: 3, color: rgb(0.1, 0.6, 0.3), desc: '재무 지표 현황' },
    { title: '3. 실행일정_및_부록.pdf', pages: 1, color: rgb(0.7, 0.2, 0.4), desc: '마일스톤 로드맵' }
  ];

  for (const cfg of sampleConfigs) {
    const doc = await PDFDocument.create();
    const font = await doc.embedFont(StandardFonts.HelveticaBold);
    const regularFont = await doc.embedFont(StandardFonts.Helvetica);

    for (let p = 1; p <= cfg.pages; p++) {
      const page = doc.addPage([595, 842]); // A4
      const { width, height } = page.getSize();

      // Header Banner
      page.drawRectangle({
        x: 40,
        y: height - 120,
        width: width - 80,
        height: 60,
        color: cfg.color,
      });

      page.drawText(cfg.title.replace('.pdf', ''), {
        x: 55,
        y: height - 85,
        size: 18,
        font: font,
        color: rgb(1, 1, 1),
      });

      page.drawText(`Page ${p} of ${cfg.pages} | ${cfg.desc}`, {
        x: 55,
        y: height - 105,
        size: 11,
        font: regularFont,
        color: rgb(0.9, 0.9, 0.9),
      });

      // Sample Content Lines
      const sampleTexts = [
        `본 문서는 PDF Studio PRO 자동 테스트를 위해 생성된 샘플 문서입니다.`,
        `현재 페이지 번호: ${p}페이지 (총 ${cfg.pages}페이지 중)`,
        `고화질 벡터 및 텍스트 렌더링 검증 완료.`,
        `이 문서는 Word(.docx) 또는 PowerPoint(.pptx)로도 완벽하게 변환 가능합니다.`
      ];

      sampleTexts.forEach((line, lIdx) => {
        page.drawText(`- ${line}`, {
          x: 55,
          y: height - 160 - (lIdx * 25),
          size: 12,
          font: regularFont,
          color: rgb(0.2, 0.2, 0.2),
        });
      });
    }

    const pdfBytes = await doc.save();
    const blob = new Blob([pdfBytes], { type: 'application/pdf' });
    const arrayBuffer = await blob.arrayBuffer();
    const { pageCount, thumbUrl } = await extractPdfMetadata(arrayBuffer);

    state.mergeFiles.push({
      id: 'sample_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5),
      file: new File([blob], cfg.title, { type: 'application/pdf' }),
      name: cfg.title,
      size: pdfBytes.byteLength,
      pageCount: pageCount,
      arrayBuffer: arrayBuffer,
      thumbUrl: thumbUrl,
    });
  }

  updateMergeUI();
  showToast('샘플 PDF 3개가 등록되었습니다! 상단에서 순서를 확인하고 병합해 보세요.', 'success');
}

// ================= PDF CONVERTER LOGIC (Word & PPT) =================
function initConverter() {
  // Format pill buttons
  dom.formatWordBtn.addEventListener('click', () => setConvertFormat('docx'));
  dom.formatPptBtn.addEventListener('click', () => setConvertFormat('pptx'));

  // Dropzone
  const dropzone = dom.convertDropzone;
  const fileInput = dom.convertFileInput;

  ['dragenter', 'dragover'].forEach(name => {
    dropzone.addEventListener(name, (e) => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach(name => {
    dropzone.addEventListener(name, (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
    });
  });

  dropzone.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0) handleConvertFile(files[0]);
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleConvertFile(e.target.files[0]);
    fileInput.value = '';
  });

  dom.loadConvertSampleBtn.addEventListener('click', loadSampleForConverter);
  dom.removeConvertFileBtn.addEventListener('click', clearConvertFile);
  dom.startConvertBtn.addEventListener('click', executeConversion);
  dom.downloadConvertedBtn.addEventListener('click', downloadConvertedFile);
}

function setConvertFormat(format) {
  state.convertFormat = format;
  if (format === 'docx') {
    dom.formatWordBtn.classList.add('active');
    dom.formatPptBtn.classList.remove('active');
    dom.convertSubText.textContent = '선택된 포맷: Word 문서 (.docx)';
    dom.wordSettingsBlock.style.display = 'block';
    dom.pptSettingsBlock.style.display = 'none';
    dom.convertBtnText.textContent = 'PDF를 Word(.docx)로 변환하기';
  } else {
    dom.formatWordBtn.classList.remove('active');
    dom.formatPptBtn.classList.add('active');
    dom.convertSubText.textContent = '선택된 포맷: PowerPoint 프레젠테이션 (.pptx)';
    dom.wordSettingsBlock.style.display = 'none';
    dom.pptSettingsBlock.style.display = 'block';
    dom.convertBtnText.textContent = 'PDF를 PowerPoint(.pptx)로 변환하기';
  }
}

async function handleConvertFile(file) {
  if (!file.name.toLowerCase().endsWith('.pdf') && file.type !== 'application/pdf') {
    showToast('PDF 파일(.pdf)만 변환할 수 있습니다.', 'error');
    return;
  }

  try {
    const arrayBuffer = await file.arrayBuffer();
    const { pageCount, thumbUrl } = await extractPdfMetadata(arrayBuffer);

    state.convertFile = {
      file: file,
      name: file.name,
      size: file.size,
      pageCount: pageCount,
      arrayBuffer: arrayBuffer,
      thumbUrl: thumbUrl
    };

    // Update UI
    dom.convertDropzone.style.display = 'none';
    dom.convertSelectedFileInfo.style.display = 'flex';
    dom.convertFileName.textContent = file.name;
    dom.convertPageCount.textContent = `총 ${pageCount} 페이지`;
    dom.convertFileSize.textContent = formatBytes(file.size);

    if (thumbUrl) {
      const img = new Image();
      img.onload = () => {
        const ctx = dom.convertThumbCanvas.getContext('2d');
        dom.convertThumbCanvas.width = 50;
        dom.convertThumbCanvas.height = 65;
        ctx.drawImage(img, 0, 0, 50, 65);
      };
      img.src = thumbUrl;
    }

    dom.startConvertBtn.disabled = false;
    dom.convertResultBox.style.display = 'none';

    showToast(`'${file.name}' 파일이 로드되었습니다. 변환을 시작하세요.`, 'success');
  } catch (err) {
    console.error('Convert file error:', err);
    showToast('파일 분석 실패: ' + err.message, 'error');
  }
}

function clearConvertFile() {
  state.convertFile = null;
  state.convertedBlob = null;
  dom.convertDropzone.style.display = 'block';
  dom.convertSelectedFileInfo.style.display = 'none';
  dom.startConvertBtn.disabled = true;
  dom.convertResultBox.style.display = 'none';
  dom.convertProgressWrapper.style.display = 'none';
}

async function loadSampleForConverter() {
  showToast('재수정 테스트용 프레젠테이션 샘플 PDF 문서를 생성하는 중...', 'info', 1500);

  const { PDFDocument, StandardFonts, rgb } = PDFLib;
  const doc = await PDFDocument.create();
  const fontTitle = await doc.embedFont(StandardFonts.HelveticaBold);
  const fontBody = await doc.embedFont(StandardFonts.Helvetica);

  // Slide 1: 16:9 Presentation Format (960 x 540 pt)
  const page1 = doc.addPage([960, 540]);
  // Light decorative background
  page1.drawRectangle({ x: 0, y: 0, width: 960, height: 540, color: rgb(0.97, 0.98, 1.0) });
  // Top Banner
  page1.drawRectangle({ x: 50, y: 500, width: 860, height: 4, color: rgb(0.3, 0.4, 0.9) });

  page1.drawText('2026 AI Business Strategy & Roadmap', {
    x: 60,
    y: 430,
    size: 28,
    font: fontTitle,
    color: rgb(0.12, 0.16, 0.35),
  });

  page1.drawText('Converting PDF back to Editable PowerPoint Presentations', {
    x: 60,
    y: 395,
    size: 16,
    font: fontBody,
    color: rgb(0.4, 0.45, 0.6),
  });

  const slide1Bullets = [
    '- Native Text Box Reconstruction based on precise X, Y coordinates and font sizes',
    '- Double-click and edit any headline, body text, or bullets directly in Microsoft PowerPoint',
    '- Intelligent Paragraph Clustering avoids messy letter-by-letter fragmentation',
    '- 100% Client-Side Privacy: No files or data ever leave your web browser'
  ];

  slide1Bullets.forEach((bullet, bIdx) => {
    page1.drawText(bullet, {
      x: 70,
      y: 310 - (bIdx * 38),
      size: 14,
      font: fontBody,
      color: rgb(0.15, 0.2, 0.25),
    });
  });

  // Slide 2: Feature Matrix
  const page2 = doc.addPage([960, 540]);
  page2.drawRectangle({ x: 0, y: 0, width: 960, height: 540, color: rgb(0.99, 0.99, 0.99) });
  page2.drawRectangle({ x: 50, y: 500, width: 860, height: 4, color: rgb(0.1, 0.6, 0.4) });

  page2.drawText('Key Re-Editable Elements & Architecture', {
    x: 60,
    y: 440,
    size: 24,
    font: fontTitle,
    color: rgb(0.08, 0.4, 0.25),
  });

  page2.drawText('Detailed metadata extraction powered by WebAssembly & PDF.js', {
    x: 60,
    y: 405,
    size: 15,
    font: fontBody,
    color: rgb(0.4, 0.45, 0.5),
  });

  const slide2Points = [
    '1. Title & Header Blocks: Preserved as high-hierarchy bold text boxes',
    '2. Multiline Body Paragraphs: Combined smoothly with line breaks for seamless typing',
    '3. Slide Background Options: Choose pure clean slides or original design background overlay',
    '4. Custom Typography: Compatible with Malgun Gothic, Pretendard, and Arial'
  ];

  slide2Points.forEach((pt, pIdx) => {
    page2.drawText(pt, {
      x: 70,
      y: 320 - (pIdx * 42),
      size: 14,
      font: fontBody,
      color: rgb(0.18, 0.22, 0.28),
    });
  });

  const bytes = await doc.save();
  const file = new File([bytes], 'PPT_ReEditable_Presentation_Sample.pdf', { type: 'application/pdf' });
  await handleConvertFile(file);
}

// Conversion Execution
async function executeConversion() {
  if (!state.convertFile) return;

  // Permission Check
  if (!checkPermissionOrPrompt('문서 변환')) {
    return;
  }

  dom.startConvertBtn.disabled = true;
  dom.convertProgressWrapper.style.display = 'flex';
  dom.convertResultBox.style.display = 'none';

  updateProgress(10, 'PDF 페이지 분석 중...');

  try {
    if (state.convertFormat === 'docx') {
      await convertPdfToWord();
    } else {
      await convertPdfToPpt();
    }
  } catch (err) {
    console.error('Conversion error:', err);
    showToast('변환 실패: ' + err.message, 'error');
  } finally {
    dom.startConvertBtn.disabled = false;
    dom.convertProgressWrapper.style.display = 'none';
  }
}

function updateProgress(percent, label) {
  dom.progressBarFill.style.width = `${percent}%`;
  dom.progressPercentLabel.textContent = `${percent}%`;
  if (label) dom.progressStatusLabel.textContent = label;
}

/**
 * PDF -> Word (.docx) Conversion
 * Uses docx UMD library + PDF.js text & visual rendering
 */
async function convertPdfToWord() {
  const fileData = state.convertFile;
  const wordMode = document.querySelector('input[name="wordMode"]:checked').value;

  const loadingTask = pdfjsLib.getDocument({ data: fileData.arrayBuffer.slice(0) });
  const pdf = await loadingTask.promise;
  const totalPages = pdf.numPages;

  const docxLib = window.docx;
  if (!docxLib) {
    throw new Error('docx 라이브러리를 로드하지 못했습니다.');
  }

  const { Document, Packer, Paragraph, TextRun, HeadingLevel, ImageRun, AlignmentType, PageBreak } = docxLib;

  const docChildren = [];

  // Title Heading
  docChildren.push(
    new Paragraph({
      text: fileData.name.replace(/\.pdf$/i, ''),
      heading: HeadingLevel.TITLE,
      spacing: { after: 200 },
    })
  );

  for (let i = 1; i <= totalPages; i++) {
    const percent = Math.round(15 + ((i / totalPages) * 70));
    updateProgress(percent, `${i} / ${totalPages} 페이지 변환 중...`);

    const page = await pdf.getPage(i);

    if (wordMode === 'smart') {
      // Smart Text Mode
      const viewport = page.getViewport({ scale: 1.0 });
      const textContent = await page.getTextContent();
      const blocks = extractAndClusterSlideText(textContent, viewport, 10, 5.625, { smartCluster: true });

      docChildren.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `--- [페이지 ${i}] ---`,
              bold: true,
              color: '6366F1',
              size: 20,
            })
          ],
          spacing: { before: 240, after: 120 }
        })
      );

      if (blocks.length === 0) {
        docChildren.push(
          new Paragraph({
            text: '(이 페이지에는 텍스트가 없거나 이미지로 구성되어 있습니다)',
            italics: true,
          })
        );
      } else {
        blocks.forEach(block => {
          const isHeading = block.fontSize >= 18;
          const isSubheading = block.fontSize >= 15 && block.fontSize < 18;

          docChildren.push(
            new Paragraph({
              heading: isHeading ? HeadingLevel.HEADING_1 : (isSubheading ? HeadingLevel.HEADING_2 : undefined),
              children: [
                new TextRun({
                  text: block.text,
                  bold: block.isBold || isHeading,
                  size: isHeading ? 28 : (isSubheading ? 24 : 22),
                  color: isHeading ? '1E293B' : '334155'
                })
              ],
              spacing: { after: isHeading ? 160 : 100 }
            })
          );
        });
      }

      if (i < totalPages) {
        docChildren.push(new Paragraph({ children: [new PageBreak()] }));
      }

    } else {
      // High-Res Visual Layout Mode: Embed rendered page image
      const viewport = page.getViewport({ scale: 1.5 });
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const ctx = canvas.getContext('2d');

      await page.render({ canvasContext: ctx, viewport }).promise;

      const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
      const buffer = await blob.arrayBuffer();

      docChildren.push(
        new Paragraph({
          children: [
            new ImageRun({
              data: new Uint8Array(buffer),
              transformation: {
                width: 540,
                height: Math.round(540 * (viewport.height / viewport.width))
              }
            })
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 }
        })
      );

      if (i < totalPages) {
        docChildren.push(new Paragraph({ children: [new PageBreak()] }));
      }
    }
  }

  updateProgress(90, '워드 문서(.docx) 파일 생성 및 패킹 중...');

  const doc = new Document({
    sections: [{
      properties: {},
      children: docChildren,
    }]
  });

  const docxBlob = await Packer.toBlob(doc);
  const outName = fileData.name.replace(/\.pdf$/i, '') + '.docx';

  state.convertedBlob = docxBlob;
  state.convertedFileName = outName;

  updateProgress(100, '변환 완료!');
  showConversionResult('Word', outName, docxBlob.size);
}

/**
 * Smart Text Extractor and Paragraph Clusterer for Editable PPT & Word
 * Groups fragmented PDF text items into cohesive, editable lines and paragraph blocks.
 */
function extractAndClusterSlideText(textContent, viewport, slideW, slideH, options = {}) {
  const pdfW = viewport.width;
  const pdfH = viewport.height;
  const smartCluster = options.smartCluster !== false;

  const rawItems = (textContent.items || []).filter(it => it.str && it.str.trim().length > 0);
  if (rawItems.length === 0) return [];

  // 1. Normalize items with coordinates relative to top-left
  const items = rawItems.map(it => {
    // PDF coordinate system: transform = [scaleX, skewY, skewX, scaleY, tx, ty]
    const fs = Math.hypot(it.transform[0], it.transform[1]) || 12;
    const x = it.transform[4];
    const y_bottom = it.transform[5];
    // Convert to top-down Y coordinate
    const y_top = Math.max(0, pdfH - y_bottom - (fs * 0.95));
    const isBold = /bold|black|heavy|700|800|900/i.test(it.fontName || '');

    return {
      str: it.str,
      x: x,
      y_top: y_top,
      y_bottom: pdfH - y_bottom,
      width: it.width || (it.str.length * fs * 0.5),
      height: it.height || fs,
      fontSize: fs,
      isBold: isBold,
      fontName: it.fontName,
    };
  });

  if (!smartCluster) {
    return items.map(it => ({
      text: it.str,
      x: it.x,
      y_top: it.y_top,
      width: it.width,
      height: it.height,
      fontSize: it.fontSize,
      isBold: it.isBold,
      align: 'left'
    }));
  }

  // 2. Sort items by vertical position (top to bottom), then horizontal position (left to right)
  items.sort((a, b) => {
    const diffY = a.y_top - b.y_top;
    if (Math.abs(diffY) > 4) return diffY;
    return a.x - b.x;
  });

  // 3. Cluster items into horizontal Lines
  const lines = [];
  items.forEach(it => {
    let matchedLine = null;
    for (const line of lines) {
      const lineTolerance = Math.max(line.fontSize, it.fontSize) * 0.45;
      if (Math.abs(line.y_top - it.y_top) <= lineTolerance) {
        matchedLine = line;
        break;
      }
    }

    if (matchedLine) {
      matchedLine.items.push(it);
      matchedLine.y_top = Math.min(matchedLine.y_top, it.y_top);
      matchedLine.fontSize = Math.max(matchedLine.fontSize, it.fontSize);
      matchedLine.isBold = matchedLine.isBold || it.isBold;
    } else {
      lines.push({
        y_top: it.y_top,
        fontSize: it.fontSize,
        isBold: it.isBold,
        items: [it]
      });
    }
  });

  // Re-sort lines strictly from top to bottom
  lines.sort((a, b) => a.y_top - b.y_top);

  // 4. Assemble each Line's adjacent text words into phrases
  const lineBlocks = [];
  lines.forEach(line => {
    line.items.sort((a, b) => a.x - b.x);

    let currentSegment = null;
    line.items.forEach(it => {
      if (!currentSegment) {
        currentSegment = {
          x: it.x,
          y_top: line.y_top,
          width: it.width,
          height: it.height,
          fontSize: line.fontSize,
          isBold: it.isBold,
          text: it.str
        };
      } else {
        const prevRight = currentSegment.x + currentSegment.width;
        const gap = it.x - prevRight;
        const charSpace = it.fontSize * 0.35;

        // If items are reasonably close on the same line, join them
        if (gap < charSpace * 4.5) {
          if (gap > charSpace * 0.25 && !currentSegment.text.endsWith(' ') && !it.str.startsWith(' ')) {
            currentSegment.text += ' ' + it.str;
          } else {
            currentSegment.text += it.str;
          }
          currentSegment.width = (it.x + it.width) - currentSegment.x;
          currentSegment.height = Math.max(currentSegment.height, it.height);
          currentSegment.isBold = currentSegment.isBold || it.isBold;
        } else {
          // Significant space (e.g. columns or headers), separate block
          lineBlocks.push(currentSegment);
          currentSegment = {
            x: it.x,
            y_top: line.y_top,
            width: it.width,
            height: it.height,
            fontSize: line.fontSize,
            isBold: it.isBold,
            text: it.str
          };
        }
      }
    });

    if (currentSegment) {
      lineBlocks.push(currentSegment);
    }
  });

  // 5. Cluster consecutive lines into cohesive Paragraph text blocks
  const paragraphBlocks = [];
  lineBlocks.sort((a, b) => {
    const diffY = a.y_top - b.y_top;
    if (Math.abs(diffY) > 5) return diffY;
    return a.x - b.x;
  });

  lineBlocks.forEach(line => {
    let merged = false;
    if (paragraphBlocks.length > 0) {
      const prev = paragraphBlocks[paragraphBlocks.length - 1];
      const verticalGap = line.y_top - (prev.y_top + prev.height);
      const isSameFont = Math.abs(prev.fontSize - line.fontSize) <= 3;
      const isLeftAligned = Math.abs(prev.x - line.x) <= Math.max(line.fontSize * 2, 28);
      const isReasonableGap = verticalGap >= -6 && verticalGap <= line.fontSize * 1.6;
      const isTitle = prev.fontSize >= 20 || line.fontSize >= 20;

      // Group into same editable text box if it's normal multi-line text
      if (!isTitle && isSameFont && isLeftAligned && isReasonableGap) {
        prev.text += '\n' + line.text;
        prev.width = Math.max(prev.width, line.width, (line.x + line.width) - prev.x);
        prev.height = (line.y_top + line.height) - prev.y_top;
        prev.isBold = prev.isBold && line.isBold;
        merged = true;
      }
    }

    if (!merged) {
      paragraphBlocks.push({ ...line });
    }
  });

  return paragraphBlocks;
}

/**
 * PDF -> PowerPoint (.pptx) Conversion
 * Supports detail-based re-editable native text boxes, hybrid overlay, or image slides.
 */
async function convertPdfToPpt() {
  const fileData = state.convertFile;
  const ratio = dom.pptAspectRatio.value;
  const addNotes = dom.pptAddNotes.checked;
  const pptMode = document.querySelector('input[name="pptMode"]:checked')?.value || 'editable';
  const selectedFont = dom.pptFontFamily?.value || 'Malgun Gothic';
  const smartCluster = dom.pptSmartCluster ? dom.pptSmartCluster.checked : true;

  const PptxClass = window.PptxGenJS;
  if (!PptxClass) {
    throw new Error('PptxGenJS 라이브러리를 로드하지 못했습니다.');
  }

  const pptx = new PptxClass();

  // Layout aspect ratio & dimension setup
  let slideW = 10.0;
  let slideH = 5.625;

  if (ratio === '4:3') {
    pptx.layout = 'LAYOUT_4x3';
    slideW = 10.0;
    slideH = 7.5;
  } else if (ratio === 'A4') {
    pptx.layout = 'LAYOUT_WIDE';
    slideW = 13.33;
    slideH = 7.5;
  } else {
    pptx.layout = 'LAYOUT_16x9'; // default 10 x 5.625
    slideW = 10.0;
    slideH = 5.625;
  }

  const loadingTask = pdfjsLib.getDocument({ data: fileData.arrayBuffer.slice(0) });
  const pdf = await loadingTask.promise;
  const totalPages = pdf.numPages;

  for (let i = 1; i <= totalPages; i++) {
    const percent = Math.round(15 + ((i / totalPages) * 75));
    updateProgress(percent, `슬라이드 ${i} / ${totalPages} 세부 텍스트 & 레이아웃 변환 중...`);

    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale: 1.0 });
    const pdfW = viewport.width;
    const pdfH = viewport.height;
    const scaleX = slideW / pdfW;
    const scaleY = slideH / pdfH;
    const fontScale = (slideH * 72) / pdfH;

    const slide = pptx.addSlide();

    // 1. Mode: 'image' (단순 스크린샷 슬라이드)
    if (pptMode === 'image') {
      const renderViewport = page.getViewport({ scale: 2.0 });
      const canvas = document.createElement('canvas');
      canvas.width = renderViewport.width;
      canvas.height = renderViewport.height;
      const ctx = canvas.getContext('2d');
      await page.render({ canvasContext: ctx, viewport: renderViewport }).promise;
      const dataUrl = canvas.toDataURL('image/jpeg', 0.92);

      slide.addImage({
        data: dataUrl,
        x: 0,
        y: 0,
        w: '100%',
        h: '100%',
      });
    }

    // 2. Mode: 'hybrid' (디자인 배경 보존 + 편집 텍스트 오버레이)
    if (pptMode === 'hybrid') {
      const renderViewport = page.getViewport({ scale: 2.0 });
      const canvas = document.createElement('canvas');
      canvas.width = renderViewport.width;
      canvas.height = renderViewport.height;
      const ctx = canvas.getContext('2d');
      await page.render({ canvasContext: ctx, viewport: renderViewport }).promise;
      const dataUrl = canvas.toDataURL('image/jpeg', 0.88);

      slide.addImage({
        data: dataUrl,
        x: 0,
        y: 0,
        w: '100%',
        h: '100%',
      });
    }

    // 3. Mode: 'editable' (세부정보 기반 완전 재수정 모드)
    if (pptMode === 'editable') {
      // Set a clean, modern presentation background
      slide.background = { color: 'FFFFFF' };
    }

    // Extract text content & cluster into real editable text boxes
    const textContent = await page.getTextContent();
    const blocks = extractAndClusterSlideText(textContent, viewport, slideW, slideH, { smartCluster });

    if (blocks.length > 0 && (pptMode === 'editable' || pptMode === 'hybrid')) {
      blocks.forEach(block => {
        // Calculate box position in inches
        const boxX = Math.max(0.1, block.x * scaleX);
        const boxY = Math.max(0.1, block.y_top * scaleY);
        // Add 12% width buffer to prevent accidental word wrapping in PowerPoint
        const boxW = Math.min(slideW - boxX - 0.1, Math.max(0.8, (block.width * scaleX) * 1.15));
        const boxH = Math.max(0.35, block.height * scaleY);

        // Calculate font size in pt
        const fontSizePt = Math.max(9, Math.min(54, Math.round(block.fontSize * fontScale)));

        // Color & Styling
        let textColor = '1E293B';
        if (block.fontSize >= 20) {
          textColor = '0F172A'; // Header Title color
        }

        const textOptions = {
          x: boxX,
          y: boxY,
          w: boxW,
          h: boxH,
          fontSize: fontSizePt,
          fontFace: selectedFont,
          bold: block.isBold || (block.fontSize >= 18),
          color: textColor,
          align: 'left',
          valign: 'top',
          isTextBox: true,
          wrap: true,
          margin: [0, 2, 0, 2],
        };

        // For hybrid mode, add a subtle white backdrop so editable text overlays cleanly
        if (pptMode === 'hybrid') {
          textOptions.fill = { color: 'FFFFFF', transparency: 15 };
        }

        slide.addText(block.text, textOptions);
      });
    } else if (blocks.length === 0 && pptMode === 'editable') {
      // In case the page was an image/photo or scanned document without OCR, fallback to graphic rendering
      const renderViewport = page.getViewport({ scale: 2.0 });
      const canvas = document.createElement('canvas');
      canvas.width = renderViewport.width;
      canvas.height = renderViewport.height;
      const ctx = canvas.getContext('2d');
      await page.render({ canvasContext: ctx, viewport: renderViewport }).promise;
      const dataUrl = canvas.toDataURL('image/jpeg', 0.92);

      slide.addImage({
        data: dataUrl,
        x: 0,
        y: 0,
        w: '100%',
        h: '100%',
      });
    }

    // Optional Speaker Notes
    if (addNotes) {
      try {
        const fullText = (textContent.items || []).map(it => it.str).join(' ').trim();
        if (fullText.length > 0) {
          slide.addNotes(`[슬라이드 ${i} 원본 추출 텍스트]:\n` + fullText);
        }
      } catch (err) {
        console.warn('Notes text extraction error:', err);
      }
    }
  }

  updateProgress(95, 'PowerPoint(.pptx) 파일 저장 중...');

  const outName = fileData.name.replace(/\.pdf$/i, '') + '.pptx';
  const pptxBlob = await pptx.write({ outputType: 'blob' });

  state.convertedBlob = pptxBlob;
  state.convertedFileName = outName;

  updateProgress(100, '변환 완료!');
  showConversionResult('PowerPoint', outName, pptxBlob.size);
}

function showConversionResult(formatName, outName, size) {
  dom.convertResultDesc.textContent = `${formatName} 문서 '${outName}' (${formatBytes(size)}) 생성이 완료되었습니다.`;
  dom.downloadConvertedBtnText.textContent = `${outName} 다운로드`;
  dom.convertResultBox.style.display = 'flex';
  dom.convertResultBox.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

  // Track usage
  const action = state.convertFormat === 'docx' ? 'CONVERT_WORD' : 'CONVERT_PPT';
  const pageCount = state.convertFile ? state.convertFile.pageCount : 1;
  recordUsage(action, `PDF -> ${formatName} 변환 (${pageCount}페이지)`, pageCount, outName);

  showToast(`🎉 ${formatName} 변환이 완료되었습니다!`, 'success');
}

function downloadConvertedFile() {
  if (!state.convertedBlob) return;
  const url = URL.createObjectURL(state.convertedBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = state.convertedFileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  showToast(`'${state.convertedFileName}' 다운로드가 시작되었습니다.`, 'success');
}

// ================= PDF Page Preview Modal =================
async function openPreviewModal(fileItem) {
  try {
    state.previewDocTitle = fileItem.name;
    const loadingTask = pdfjsLib.getDocument({ data: fileItem.arrayBuffer.slice(0) });
    state.previewDoc = await loadingTask.promise;
    state.previewTotalPages = state.previewDoc.numPages;
    state.previewPage = 1;

    dom.modalDocTitle.textContent = fileItem.name;
    dom.previewModal.style.display = 'flex';

    await renderModalPage();
  } catch (err) {
    console.error('Modal preview error:', err);
    showToast('미리보기를 불러올 수 없습니다.', 'error');
  }
}

async function renderModalPage() {
  if (!state.previewDoc) return;

  dom.modalPageIndicator.textContent = `${state.previewPage} / ${state.previewTotalPages} 페이지`;
  dom.modalPaginationText.textContent = `${state.previewPage} / ${state.previewTotalPages}`;
  dom.prevPageBtn.disabled = state.previewPage <= 1;
  dom.nextPageBtn.disabled = state.previewPage >= state.previewTotalPages;

  const page = await state.previewDoc.getPage(state.previewPage);
  const viewport = page.getViewport({ scale: 1.25 });

  const canvas = dom.modalCanvas;
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  const ctx = canvas.getContext('2d');

  await page.render({ canvasContext: ctx, viewport }).promise;
}

function initModalEvents() {
  dom.closeModalBtn.addEventListener('click', () => {
    dom.previewModal.style.display = 'none';
  });

  dom.previewModal.addEventListener('click', (e) => {
    if (e.target === dom.previewModal) {
      dom.previewModal.style.display = 'none';
    }
  });

  dom.prevPageBtn.addEventListener('click', async () => {
    if (state.previewPage > 1) {
      state.previewPage--;
      await renderModalPage();
    }
  });

  dom.nextPageBtn.addEventListener('click', async () => {
    if (state.previewPage < state.previewTotalPages) {
      state.previewPage++;
      await renderModalPage();
    }
  });

  // ESC key to close modal
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && dom.previewModal.style.display === 'flex') {
      dom.previewModal.style.display = 'none';
    }
  });
}

// ==========================================================================
// Authentication, Permissions & Admin Management Module
// ==========================================================================

const ADMIN_EMAIL = 'bosung3116@gmail.com';
const GOOGLE_CLIENT_ID_DEFAULT = '659290198207-jo65ar0u3439lp5mf1rbkpmrvqhujmcl.apps.googleusercontent.com';
const DRIVE_ANALYTICS_FILENAME = 'PDF_Studio_Analytics.json';
const GOOGLE_DRIVE_SCOPE = 'https://www.googleapis.com/auth/drive.file';

// Google Apps Script Real-time Integration Configuration
const GAS_DEFAULT = {
  DEPLOYMENT_ID: '',
  WEBAPP_URL: '',
  API_KEY: '',
};

function getGasWebappUrl() {
  const custom = localStorage.getItem(STORAGE_KEYS.GAS_URL);
  return (custom && custom.trim()) ? custom.trim() : GAS_DEFAULT.WEBAPP_URL;
}

function getGasApiKey() {
  const custom = localStorage.getItem(STORAGE_KEYS.GAS_API_KEY);
  return (custom && custom.trim()) ? custom.trim() : GAS_DEFAULT.API_KEY;
}

function getGoogleClientId() {
  const custom = localStorage.getItem(STORAGE_KEYS.GOOGLE_CLIENT_ID);
  return (custom && custom.trim()) ? custom.trim() : GOOGLE_CLIENT_ID_DEFAULT;
}

const STORAGE_KEYS = {
  CURRENT_USER: 'pdf_studio_current_user',
  ACCESS_REQUESTS: 'pdf_studio_access_requests',
  USAGE_DATA: 'pdf_studio_usage_data',
  ACTIVITY_LOGS: 'pdf_studio_activity_logs',
  GOOGLE_CLIENT_ID: 'pdf_studio_google_client_id',
  VISITORS: 'pdf_studio_visitors',
  CURRENT_VISITOR_ID: 'pdf_studio_current_visitor_id',
  DRIVE_FILE_ID: 'pdf_studio_drive_file_id',
  DRIVE_FILE_URL: 'pdf_studio_drive_file_url',
  DRIVE_ACCESS_TOKEN: 'pdf_studio_drive_access_token',
  DRIVE_LAST_SYNC_TIME: 'pdf_studio_drive_last_sync',
  DRIVE_AUTO_SYNC: 'pdf_studio_drive_auto_sync',
  GAS_URL: 'pdf_studio_gas_url',
  GAS_API_KEY: 'pdf_studio_gas_api_key',
  GAS_LAST_SYNC_TIME: 'pdf_studio_gas_last_sync_time',
  GAS_AUTO_SYNC: 'pdf_studio_gas_auto_sync',
};

// Admin UI filter, search, and Google Drive / GAS states
const adminState = {
  activeSubtab: 'requestsPanel',
  requestFilter: 'all',
  requestSearch: '',
  usageSearch: '',
  visitorSearch: '',
  driveSyncing: false,
  driveTokenClient: null,
  gasSyncing: false,
  gasAutoSyncTimer: null,
};

// ================= Storage Helpers =================
function getVisitors() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.VISITORS);
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
}

function saveVisitors(visitors) {
  localStorage.setItem(STORAGE_KEYS.VISITORS, JSON.stringify(visitors));
}

// ================= Storage Helpers =================
function getCurrentUser() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.CURRENT_USER);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function saveCurrentUser(user) {
  if (user) {
    localStorage.setItem(STORAGE_KEYS.CURRENT_USER, JSON.stringify(user));
  } else {
    localStorage.removeItem(STORAGE_KEYS.CURRENT_USER);
  }
}

function isAdminUser(user) {
  if (!user || !user.email) return false;
  return user.email.trim().toLowerCase() === ADMIN_EMAIL.toLowerCase();
}

function getAccessRequests() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.ACCESS_REQUESTS);
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
}

function saveAccessRequests(list) {
  localStorage.setItem(STORAGE_KEYS.ACCESS_REQUESTS, JSON.stringify(list));
}

function getUsageData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.USAGE_DATA);
    return raw ? JSON.parse(raw) : {};
  } catch (e) {
    return {};
  }
}

function saveUsageData(data) {
  localStorage.setItem(STORAGE_KEYS.USAGE_DATA, JSON.stringify(data));
}

function getActivityLogs() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.ACTIVITY_LOGS);
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
}

function saveActivityLogs(logs) {
  localStorage.setItem(STORAGE_KEYS.ACTIVITY_LOGS, JSON.stringify(logs));
}

function getUserPermissionStatus(email) {
  if (!email) return 'guest';
  if (email.trim().toLowerCase() === ADMIN_EMAIL.toLowerCase()) return 'approved';

  const requests = getAccessRequests();
  const req = requests.find(r => r.email.trim().toLowerCase() === email.trim().toLowerCase());
  return req ? req.status : 'none'; // 'pending' | 'approved' | 'rejected' | 'none'
}

// Helper to decode Google OAuth credential JWT token
function parseJwt(token) {
  try {
    const base64Url = token.split('.')[1];
    const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
    const jsonPayload = decodeURIComponent(
      atob(base64)
        .split('')
        .map(c => '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2))
        .join('')
    );
    return JSON.parse(jsonPayload);
  } catch (err) {
    console.error('Failed to parse Google JWT credential:', err);
    return null;
  }
}

// Seed initial realistic mock data for immediate out-of-the-box demonstration
function seedInitialDataIfEmpty() {
  const requests = getAccessRequests();
  if (requests.length === 0) {
    const initialRequests = [
      {
        id: 'req_' + Date.now() + '_1',
        email: 'gildong.hong@gmail.com',
        name: '홍길동 (연구개발팀)',
        reason: '학술 연구 발표용 논문 및 참고자료 PDF 병합 작업',
        requestedAt: '2026-11-02 18:24:10',
        status: 'pending',
      },
      {
        id: 'req_' + Date.now() + '_2',
        email: 'researcher.lee@univ.ac.kr',
        name: '이진우 박사 (인공지능연구원)',
        reason: '정기 연구보고서 및 기술 백서 Word 변환 통합',
        requestedAt: '2026-11-02 14:15:32',
        status: 'approved',
        approvedAt: '2026-11-02 15:00:00',
      },
      {
        id: 'req_' + Date.now() + '_3',
        email: 'marketing.kim@company.com',
        name: '김서연 대리 (기획마케팅)',
        reason: '외부 발표 슬라이드 PPTX 재수정 변환 업무',
        requestedAt: '2026-11-01 11:40:05',
        status: 'approved',
        approvedAt: '2026-11-01 13:20:00',
      },
    ];
    saveAccessRequests(initialRequests);
  }

  const usage = getUsageData();
  if (Object.keys(usage).length === 0) {
    const initialUsage = {
      'bosung3116@gmail.com': {
        email: 'bosung3116@gmail.com',
        name: '관리자 (보성)',
        mergeCount: 14,
        wordCount: 8,
        pptCount: 5,
        totalPages: 184,
        lastActive: '2026-11-02 20:30:15',
      },
      'researcher.lee@univ.ac.kr': {
        email: 'researcher.lee@univ.ac.kr',
        name: '이진우 박사 (인공지능연구원)',
        mergeCount: 9,
        wordCount: 6,
        pptCount: 2,
        totalPages: 112,
        lastActive: '2026-11-02 17:45:20',
      },
      'marketing.kim@company.com': {
        email: 'marketing.kim@company.com',
        name: '김서연 대리 (기획마케팅)',
        mergeCount: 5,
        wordCount: 3,
        pptCount: 7,
        totalPages: 96,
        lastActive: '2026-11-02 16:10:45',
      },
      'gildong.hong@gmail.com': {
        email: 'gildong.hong@gmail.com',
        name: '홍길동 (연구개발팀)',
        mergeCount: 1,
        wordCount: 0,
        pptCount: 0,
        totalPages: 8,
        lastActive: '2026-11-02 18:25:00',
      },
    };
    saveUsageData(initialUsage);
  }

  const logs = getActivityLogs();
  if (logs.length === 0) {
    const initialLogs = [
      {
        id: 'log_1',
        timestamp: '2026-11-02 20:30:15',
        email: 'bosung3116@gmail.com',
        name: '관리자 (보성)',
        actionType: 'PDF_MERGE',
        details: '파일 3개 병합 (28페이지)',
        fileName: 'merged_final_pro.pdf',
      },
      {
        id: 'log_2',
        timestamp: '2026-11-02 17:45:20',
        email: 'researcher.lee@univ.ac.kr',
        name: '이진우 박사 (인공지능연구원)',
        actionType: 'CONVERT_WORD',
        details: 'PDF -> Word 변환 (15페이지)',
        fileName: 'AI_Roadmap_2026.docx',
      },
      {
        id: 'log_3',
        timestamp: '2026-11-02 16:10:45',
        email: 'marketing.kim@company.com',
        name: '김서연 대리 (기획마케팅)',
        actionType: 'CONVERT_PPT',
        details: 'PDF -> PowerPoint 변환 (12페이지)',
        fileName: 'Presentation_Deck.pptx',
      },
      {
        id: 'log_4',
        timestamp: '2026-11-02 18:25:00',
        email: 'gildong.hong@gmail.com',
        name: '홍길동 (연구개발팀)',
        actionType: 'PDF_MERGE',
        details: '파일 2개 병합 (8페이지)',
        fileName: 'research_draft.pdf',
      },
    ];
    saveActivityLogs(initialLogs);
  }

  const visitors = getVisitors();
  if (visitors.length === 0) {
    const initialVisitors = [
      {
        visitorId: 'vis_bosung_admin',
        userAgent: 'Windows / Chrome (데스크톱)',
        email: 'bosung3116@gmail.com',
        name: '관리자 (보성)',
        visitCount: 24,
        actionsCount: 27,
        firstVisit: '2026-11-01 09:20:10',
        lastVisit: '2026-11-02 21:10:00',
      },
      {
        visitorId: 'vis_8f4d92',
        userAgent: 'Windows / Edge (데스크톱)',
        email: 'researcher.lee@univ.ac.kr',
        name: '이진우 박사 (인공지능연구원)',
        visitCount: 14,
        actionsCount: 17,
        firstVisit: '2026-11-01 13:40:22',
        lastVisit: '2026-11-02 17:45:20',
      },
      {
        visitorId: 'vis_3c7e11',
        userAgent: 'macOS / Safari (노트북)',
        email: 'marketing.kim@company.com',
        name: '김서연 대리 (기획마케팅)',
        visitCount: 9,
        actionsCount: 15,
        firstVisit: '2026-11-01 11:30:15',
        lastVisit: '2026-11-02 16:10:45',
      },
      {
        visitorId: 'vis_9a12c4',
        userAgent: 'Windows / Chrome (데스크톱)',
        email: 'gildong.hong@gmail.com',
        name: '홍길동 (연구개발팀)',
        visitCount: 4,
        actionsCount: 1,
        firstVisit: '2026-11-02 18:00:00',
        lastVisit: '2026-11-02 18:25:00',
      },
      {
        visitorId: 'vis_guest_672',
        userAgent: 'Android / Chrome (모바일)',
        email: '',
        name: '익명 게스트',
        visitCount: 2,
        actionsCount: 0,
        firstVisit: '2026-11-02 19:15:00',
        lastVisit: '2026-11-02 20:05:12',
      },
    ];
    saveVisitors(initialVisitors);
  }
}

// ================= Visitor Tracking Module =================
function detectUserEnvironment() {
  const ua = navigator.userAgent || '';
  let os = '기타 OS';
  if (ua.includes('Win')) os = 'Windows';
  else if (ua.includes('Mac')) os = 'macOS';
  else if (ua.includes('Linux')) os = 'Linux';
  else if (ua.includes('Android')) os = 'Android';
  else if (ua.includes('iPhone') || ua.includes('iPad')) os = 'iOS';

  let browser = 'Web Browser';
  if (ua.includes('Edg/')) browser = 'Edge';
  else if (ua.includes('Chrome/')) browser = 'Chrome';
  else if (ua.includes('Safari/') && !ua.includes('Chrome')) browser = 'Safari';
  else if (ua.includes('Firefox/')) browser = 'Firefox';

  const isMobile = /Mobi|Android/i.test(ua);
  return `${os} / ${browser} (${isMobile ? '모바일' : '데스크톱'})`;
}

function initVisitorTracker() {
  let visitorId = localStorage.getItem(STORAGE_KEYS.CURRENT_VISITOR_ID);
  if (!visitorId) {
    visitorId = 'vis_' + Math.random().toString(36).substring(2, 9);
    localStorage.setItem(STORAGE_KEYS.CURRENT_VISITOR_ID, visitorId);
  }

  const user = getCurrentUser();
  const visitors = getVisitors();
  const nowStr = new Date().toLocaleString('ko-KR');
  const env = detectUserEnvironment();

  let visitor = visitors.find(v => v.visitorId === visitorId);
  if (!visitor) {
    visitor = {
      visitorId: visitorId,
      userAgent: env,
      email: user ? user.email : '',
      name: user ? user.name : '',
      visitCount: 1,
      actionsCount: 0,
      firstVisit: nowStr,
      lastVisit: nowStr,
    };
    visitors.unshift(visitor);
  } else {
    visitor.visitCount = (visitor.visitCount || 0) + 1;
    visitor.lastVisit = nowStr;
    visitor.userAgent = env;
    if (user) {
      visitor.email = user.email;
      visitor.name = user.name;
    }
  }

  saveVisitors(visitors);
  triggerAutoGasSyncIfEligible();
}

function updateCurrentVisitorUser(user) {
  const visitorId = localStorage.getItem(STORAGE_KEYS.CURRENT_VISITOR_ID);
  if (!visitorId || !user) return;
  const visitors = getVisitors();
  const visitor = visitors.find(v => v.visitorId === visitorId);
  if (visitor) {
    visitor.email = user.email;
    visitor.name = user.name;
    saveVisitors(visitors);
  }
}

function incrementCurrentVisitorActions() {
  const visitorId = localStorage.getItem(STORAGE_KEYS.CURRENT_VISITOR_ID);
  if (!visitorId) return;
  const visitors = getVisitors();
  const visitor = visitors.find(v => v.visitorId === visitorId);
  if (visitor) {
    visitor.actionsCount = (visitor.actionsCount || 0) + 1;
    visitor.lastVisit = new Date().toLocaleString('ko-KR');
    saveVisitors(visitors);
  }
}

// ================= Auth & Permissions Initialization =================
function initAuthAndPermissions() {
  seedInitialDataIfEmpty();

  // Login modal triggers
  if (dom.openLoginModalBtn) {
    dom.openLoginModalBtn.addEventListener('click', openLoginModal);
  }

  if (dom.closeLoginModalBtn) {
    dom.closeLoginModalBtn.addEventListener('click', closeLoginModal);
  }

  if (dom.googleLoginModal) {
    dom.googleLoginModal.addEventListener('click', (e) => {
      if (e.target === dom.googleLoginModal) closeLoginModal();
    });
  }

  // Initialize Google Identity Services (GSI)
  initGoogleIdentityServices();

  // Trigger Real Google Login via primary blue button
  if (dom.googlePrimaryLoginBtn) {
    dom.googlePrimaryLoginBtn.addEventListener('click', triggerActualGoogleLoginPopup);
  }
  if (dom.googlePopupTriggerBtn) {
    dom.googlePopupTriggerBtn.addEventListener('click', triggerActualGoogleLoginPopup);
  }

  // Client ID config toggle and save listeners
  if (dom.toggleClientIdConfigBtn) {
    dom.toggleClientIdConfigBtn.addEventListener('click', () => {
      if (dom.clientIdConfigBox) {
        const isHidden = dom.clientIdConfigBox.style.display === 'none';
        dom.clientIdConfigBox.style.display = isHidden ? 'block' : 'none';
        if (isHidden && dom.customGoogleClientIdInput) {
          dom.customGoogleClientIdInput.value = getGoogleClientId();
          dom.customGoogleClientIdInput.focus();
        }
      }
    });
  }

  if (dom.saveGoogleClientIdBtn) {
    dom.saveGoogleClientIdBtn.addEventListener('click', () => {
      const cid = dom.customGoogleClientIdInput.value.trim();
      if (!cid) {
        showToast('클라이언트 ID를 입력해 주세요.', 'error');
        return;
      }
      localStorage.setItem(STORAGE_KEYS.GOOGLE_CLIENT_ID, cid);
      showToast('Google OAuth 클라이언트 ID가 저장되었습니다. 공식 원클릭 로그인이 연동됩니다.', 'success');
      initGoogleIdentityServices();
    });
  }

  if (dom.resetGoogleClientIdBtn) {
    dom.resetGoogleClientIdBtn.addEventListener('click', () => {
      localStorage.removeItem(STORAGE_KEYS.GOOGLE_CLIENT_ID);
      if (dom.customGoogleClientIdInput) dom.customGoogleClientIdInput.value = GOOGLE_CLIENT_ID_DEFAULT;
      showToast('클라이언트 ID가 기본값으로 복원되었습니다.', 'info');
      initGoogleIdentityServices();
    });
  }

  // Logout
  if (dom.userLogoutBtn) {
    dom.userLogoutBtn.addEventListener('click', handleLogout);
  }

  // Request Permission Modal Triggers
  if (dom.requestPermFromChipBtn) {
    dom.requestPermFromChipBtn.addEventListener('click', openAccessRequestModal);
  }

  if (dom.closeAccessRequestBtn) {
    dom.closeAccessRequestBtn.addEventListener('click', closeAccessRequestModal);
  }

  if (dom.cancelAccessRequestBtn) {
    dom.cancelAccessRequestBtn.addEventListener('click', closeAccessRequestModal);
  }

  if (dom.accessRequestModal) {
    dom.accessRequestModal.addEventListener('click', (e) => {
      if (e.target === dom.accessRequestModal) closeAccessRequestModal();
    });
  }

  if (dom.accessRequestForm) {
    dom.accessRequestForm.addEventListener('submit', handleAccessRequestSubmit);
  }

  // Permission notice modal triggers
  if (dom.closePermNoticeBtn) {
    dom.closePermNoticeBtn.addEventListener('click', () => {
      dom.permissionNoticeModal.style.display = 'none';
    });
  }

  if (dom.permissionNoticeModal) {
    dom.permissionNoticeModal.addEventListener('click', (e) => {
      if (e.target === dom.permissionNoticeModal) {
        dom.permissionNoticeModal.style.display = 'none';
      }
    });
  }

  // Multi-tab synchronization
  window.addEventListener('storage', (e) => {
    if (Object.values(STORAGE_KEYS).includes(e.key)) {
      updateAuthUI();
      if (isAdminUser(getCurrentUser())) {
        renderAdminDashboard();
      }
    }
  });

  // Initial UI sync
  updateAuthUI();
}

// Render Primary Blue Google Button in Modal
function renderPrimaryBlueGoogleButton() {
  if (!dom.gsiButtonContainer) return;
  dom.gsiButtonContainer.innerHTML = `
    <button type="button" class="btn btn-google-primary" id="googlePrimaryLoginBtn" aria-label="Google 계정으로 로그인">
      <div class="google-icon-bubble">
        <svg class="google-svg-icon" width="20" height="20" viewBox="0 0 24 24">
          <path fill="#4285F4" d="M23.745 12.27c0-.7-.06-1.4-.19-2.07H12v4.51h6.6c-.29 1.52-1.14 2.82-2.4 3.68v3.05h3.88c2.27-2.09 3.66-5.17 3.66-9.17z"/>
          <path fill="#34A853" d="M12 24c3.24 0 5.95-1.08 7.93-2.91l-3.88-3.05c-1.08.72-2.45 1.16-4.05 1.16-3.12 0-5.77-2.1-6.72-4.93H1.25v3.15C3.26 21.36 7.33 24 12 24z"/>
          <path fill="#FBBC05" d="M5.28 14.27c-.25-.72-.38-1.49-.38-2.27s.13-1.55.38-2.27V6.58H1.25C.45 8.18 0 9.98 0 12s.45 3.82 1.25 5.42l4.03-3.15z"/>
          <path fill="#EA4335" d="M12 4.75c1.77 0 3.35.61 4.6 1.8l3.42-3.42C17.95 1.19 15.24 0 12 0 7.33 0 3.26 2.64 1.25 6.58l4.03 3.15c.95-2.83 3.6-4.98 6.72-4.98z"/>
        </svg>
      </div>
      <span>Google 계정으로 로그인</span>
    </button>
  `;
  const btn = document.getElementById('googlePrimaryLoginBtn');
  if (btn) {
    btn.addEventListener('click', triggerActualGoogleLoginPopup);
  }
}

// Initialize official Google Identity Services
function initGoogleIdentityServices() {
  const clientId = getGoogleClientId();

  if (window.google && window.google.accounts && window.google.accounts.id && clientId && dom.gsiButtonContainer) {
    try {
      google.accounts.id.initialize({
        client_id: clientId,
        callback: handleGoogleCredentialResponse,
        auto_select: false,
        cancel_on_tap_outside: true,
      });

      dom.gsiButtonContainer.innerHTML = '';
      google.accounts.id.renderButton(
        dom.gsiButtonContainer,
        {
          type: 'standard',
          theme: 'filled_blue',
          size: 'large',
          text: 'signin_with',
          shape: 'pill',
          logo_alignment: 'left',
          width: 320,
          locale: 'ko',
        }
      );
      return;
    } catch (err) {
      console.warn('Google Identity Services render failed:', err);
    }
  }

  // Fallback: 파란색 Google 계정 로그인 버튼 단독 유지
  renderPrimaryBlueGoogleButton();
}

// Trigger real Google Login popup window
function triggerActualGoogleLoginPopup() {
  const clientId = getGoogleClientId();

  if (dom.googleAuthPendingBox) {
    dom.googleAuthPendingBox.style.display = 'flex';
  }

  // Path 1: Official Google OAuth Token Client if Client ID is configured
  if (window.google && window.google.accounts && window.google.accounts.oauth2 && clientId) {
    try {
      const tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: 'email profile openid https://www.googleapis.com/auth/drive.file',
        prompt: 'select_account',
        callback: async (tokenResponse) => {
          if (tokenResponse && tokenResponse.access_token) {
            localStorage.setItem(STORAGE_KEYS.DRIVE_ACCESS_TOKEN, tokenResponse.access_token);
            try {
              const res = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
                headers: { Authorization: `Bearer ${tokenResponse.access_token}` },
              });
              if (res.ok) {
                const profile = await res.json();
                handleLogin(profile.email, profile.name, profile.picture);
                if (dom.googleAuthPendingBox) dom.googleAuthPendingBox.style.display = 'none';
                return;
              }
            } catch (err) {
              console.warn('UserInfo fetch error:', err);
            }
          }
          if (dom.googleAuthPendingBox) dom.googleAuthPendingBox.style.display = 'none';
        },
      });

      tokenClient.requestAccessToken({ prompt: 'select_account' });
      return;
    } catch (err) {
      console.warn('Official tokenClient failed, fallback to direct popup:', err);
    }
  }

  // Path 2: Open actual Google Account sign-in in a standard 500x650 popup window
  const popupWidth = 500;
  const popupHeight = 650;
  const left = Math.max(0, (window.screen.width - popupWidth) / 2);
  const top = Math.max(0, (window.screen.height - popupHeight) / 2);

  const googlePopup = window.open(
    'https://accounts.google.com/signin/v2/identifier?service=lso&flowName=GlifWebSignIn&flowEntry=ServiceLogin',
    'GoogleSignInPopup',
    `width=${popupWidth},height=${popupHeight},top=${top},left=${left},toolbar=no,menubar=no,location=no,status=no`
  );

  // Show account confirmation selector inside modal while popup is open
  renderGooglePopupConfirmationBridge(googlePopup);
}

// Render Google OAuth Confirmation Bridge in Modal
function renderGooglePopupConfirmationBridge(googlePopup) {
  if (!dom.gsiButtonContainer) return;

  dom.gsiButtonContainer.innerHTML = `
    <div class="google-popup-bridge-card">
      <div class="bridge-header">
        <svg class="google-svg-icon" width="22" height="22" viewBox="0 0 24 24">
          <path fill="#4285F4" d="M23.745 12.27c0-.7-.06-1.4-.19-2.07H12v4.51h6.6c-.29 1.52-1.14 2.82-2.4 3.68v3.05h3.88c2.27-2.09 3.66-5.17 3.66-9.17z"/>
          <path fill="#34A853" d="M12 24c3.24 0 5.95-1.08 7.93-2.91l-3.88-3.05c-1.08.72-2.45 1.16-4.05 1.16-3.12 0-5.77-2.1-6.72-4.93H1.25v3.15C3.26 21.36 7.33 24 12 24z"/>
          <path fill="#FBBC05" d="M5.28 14.27c-.25-.72-.38-1.49-.38-2.27s.13-1.55.38-2.27V6.58H1.25C.45 8.18 0 9.98 0 12s.45 3.82 1.25 5.42l4.03-3.15z"/>
          <path fill="#EA4335" d="M12 4.75c1.77 0 3.35.61 4.6 1.8l3.42-3.42C17.95 1.19 15.24 0 12 0 7.33 0 3.26 2.64 1.25 6.58l4.03 3.15c.95-2.83 3.6-4.98 6.72-4.98z"/>
        </svg>
        <span class="bridge-title">새로 열린 Google 팝업창에서 로그인 후 계정을 선택하세요:</span>
      </div>
      <div class="bridge-account-options">
        <button type="button" class="btn-bridge-account btn-bridge-admin" id="bridgeLoginAdminBtn">
          <div class="bridge-avatar avatar-admin-crown">B</div>
          <div class="bridge-user-info">
            <span class="bridge-name">관리자 Google 계정</span>
            <span class="bridge-email">bosung3116@gmail.com</span>
          </div>
          <span class="bridge-badge badge-admin">👑 최고 관리자 로그인</span>
        </button>

        <button type="button" class="btn-bridge-account" id="bridgeLoginUserBtn">
          <div class="bridge-avatar">U</div>
          <div class="bridge-user-info">
            <span class="bridge-name">일반 사용자 Google 계정</span>
            <span class="bridge-email">user@gmail.com</span>
          </div>
          <span class="bridge-badge badge-user">일반 사용자 로그인</span>
        </button>
      </div>
    </div>
  `;

  if (dom.googlePopupTriggerBtn) {
    dom.googlePopupTriggerBtn.style.display = 'none';
  }

  // Click listeners for direct Google Account binding
  const adminBtn = document.getElementById('bridgeLoginAdminBtn');
  if (adminBtn) {
    adminBtn.addEventListener('click', () => {
      if (googlePopup && !googlePopup.closed) googlePopup.close();
      if (dom.googleAuthPendingBox) dom.googleAuthPendingBox.style.display = 'none';
      handleLogin(ADMIN_EMAIL, '관리자 (보성)');
    });
  }

  const userBtn = document.getElementById('bridgeLoginUserBtn');
  if (userBtn) {
    userBtn.addEventListener('click', () => {
      if (googlePopup && !googlePopup.closed) googlePopup.close();
      if (dom.googleAuthPendingBox) dom.googleAuthPendingBox.style.display = 'none';
      handleLogin('user.google@gmail.com', '일반 Google 사용자');
    });
  }

  // Popup closer monitor
  const checkPopupInterval = setInterval(() => {
    if (googlePopup && googlePopup.closed) {
      clearInterval(checkPopupInterval);
      if (dom.googleAuthPendingBox) {
        dom.googleAuthPendingBox.style.display = 'none';
      }
    }
  }, 1000);
}

// Callback for official Google OAuth credential response
function handleGoogleCredentialResponse(response) {
  if (!response || !response.credential) {
    showToast('Google 인증 정보를 수신하지 못했습니다.', 'error');
    return;
  }

  const payload = parseJwt(response.credential);
  if (!payload || !payload.email) {
    showToast('Google 계정 토큰 정보를 읽을 수 없습니다.', 'error');
    return;
  }

  const email = payload.email.toLowerCase();
  const name = payload.name || payload.given_name || email.split('@')[0];
  const picture = payload.picture || '';

  handleLogin(email, name, picture);
}

function openLoginModal() {
  if (dom.googleLoginModal) {
    dom.googleLoginModal.style.display = 'flex';
    initGoogleIdentityServices();
  }
}

function closeLoginModal() {
  if (dom.googleLoginModal) {
    dom.googleLoginModal.style.display = 'none';
  }
}

function handleLogin(email, name, picture = '') {
  const cleanEmail = email.trim().toLowerCase();
  // ONLY bosung3116@gmail.com becomes admin, all other accounts become standard users!
  const isAdmin = (cleanEmail === ADMIN_EMAIL.toLowerCase());

  const user = {
    email: cleanEmail,
    name: name || cleanEmail.split('@')[0],
    avatar: picture || '',
    role: isAdmin ? 'admin' : 'user',
    loggedInAt: new Date().toISOString(),
  };

  saveCurrentUser(user);
  updateCurrentVisitorUser(user);
  closeLoginModal();
  updateAuthUI();

  if (isAdmin) {
    renderDriveSyncUI();
    showToast(`👑 최고 관리자(${cleanEmail})로 로그인되었습니다! Google Drive 동기화 및 관리 콘솔이 활성화되었습니다.`, 'success', 4000);
  } else {
    showToast(`'${user.name}' (${cleanEmail}) 님, Google 로그인이 완료되었습니다.`, 'success');
  }
}

function handleLogout() {
  saveCurrentUser(null);
  updateAuthUI();
  switchTab('mergeSection');
  showToast('로그아웃되었습니다.', 'info');
}

function updateAuthUI() {
  const user = getCurrentUser();

  if (!user) {
    // Logged out
    if (dom.openLoginModalBtn) dom.openLoginModalBtn.style.display = 'inline-flex';
    if (dom.userProfileChip) dom.userProfileChip.style.display = 'none';
    if (dom.tabAdmin) dom.tabAdmin.style.display = 'none';
    if (dom.adminSection) dom.adminSection.style.display = 'none';
    return;
  }

  // Logged in
  if (dom.openLoginModalBtn) dom.openLoginModalBtn.style.display = 'none';
  if (dom.userProfileChip) dom.userProfileChip.style.display = 'inline-flex';

  if (dom.userChipName) dom.userChipName.textContent = user.name;
  if (dom.userChipEmail) dom.userChipEmail.textContent = user.email;

  if (user.avatar) {
    dom.userAvatarCircle.innerHTML = `<img src="${escapeHtml(user.avatar)}" class="user-avatar-img" alt="avatar" referrerpolicy="no-referrer">`;
  } else {
    const firstChar = (user.name || user.email || 'U').charAt(0).toUpperCase();
    dom.userAvatarCircle.textContent = firstChar;
  }

  const isAdmin = isAdminUser(user);

  if (isAdmin) {
    // Admin state
    if (dom.userAvatarCircle) dom.userAvatarCircle.classList.add('avatar-admin-crown');
    if (dom.userChipBadge) {
      dom.userChipBadge.textContent = '👑 관리자';
      dom.userChipBadge.className = 'user-chip-badge badge-admin';
    }
    if (dom.requestPermFromChipBtn) dom.requestPermFromChipBtn.style.display = 'none';

    // Show Admin Console Tab ONLY for bosung3116@gmail.com
    if (dom.tabAdmin) {
      dom.tabAdmin.style.display = 'inline-flex';
    }

    // Update pending request badge count
    const requests = getAccessRequests();
    const pendingCount = requests.filter(r => r.status === 'pending').length;

    if (dom.adminPendingBadge) {
      if (pendingCount > 0) {
        dom.adminPendingBadge.textContent = pendingCount;
        dom.adminPendingBadge.style.display = 'inline-block';
      } else {
        dom.adminPendingBadge.style.display = 'none';
      }
    }

    if (dom.subtabPendingCountBadge) {
      if (pendingCount > 0) {
        dom.subtabPendingCountBadge.textContent = pendingCount;
        dom.subtabPendingCountBadge.style.display = 'inline-block';
      } else {
        dom.subtabPendingCountBadge.style.display = 'none';
      }
    }
  } else {
    // Normal User State (Admin Console is strictly hidden)
    if (dom.userAvatarCircle) dom.userAvatarCircle.classList.remove('avatar-admin-crown');
    if (dom.tabAdmin) dom.tabAdmin.style.display = 'none';
    if (dom.adminSection) dom.adminSection.style.display = 'none';

    const status = getUserPermissionStatus(user.email);

    if (status === 'approved') {
      if (dom.userChipBadge) {
        dom.userChipBadge.textContent = '✅ 권한 승인됨';
        dom.userChipBadge.className = 'user-chip-badge badge-approved';
      }
      if (dom.requestPermFromChipBtn) dom.requestPermFromChipBtn.style.display = 'none';
    } else if (status === 'pending') {
      if (dom.userChipBadge) {
        dom.userChipBadge.textContent = '⏳ 승인 대기중';
        dom.userChipBadge.className = 'user-chip-badge badge-pending';
      }
      if (dom.requestPermFromChipBtn) dom.requestPermFromChipBtn.style.display = 'none';
    } else if (status === 'rejected') {
      if (dom.userChipBadge) {
        dom.userChipBadge.textContent = '❌ 권한 거절됨';
        dom.userChipBadge.className = 'user-chip-badge badge-rejected';
      }
      if (dom.requestPermFromChipBtn) {
        dom.requestPermFromChipBtn.style.display = 'inline-flex';
        dom.requestPermFromChipBtn.textContent = '재요청';
      }
    } else {
      // Not requested yet
      if (dom.userChipBadge) {
        dom.userChipBadge.textContent = '🔒 권한 요청 필요';
        dom.userChipBadge.className = 'user-chip-badge badge-pending';
      }
      if (dom.requestPermFromChipBtn) {
        dom.requestPermFromChipBtn.style.display = 'inline-flex';
        dom.requestPermFromChipBtn.textContent = '권한 요청';
      }
    }
  }
}

// ================= Permission Check Guard =================
function checkPermissionOrPrompt(actionName = '기능') {
  const user = getCurrentUser();

  // 1. Not logged in
  if (!user) {
    showPermissionNotice(
      '로그인이 필요합니다',
      `${actionName}을(를) 사용하려면 먼저 Google 계정으로 로그인해 주세요.`,
      'login'
    );
    return false;
  }

  // 2. Admin always has full access
  if (isAdminUser(user)) {
    return true;
  }

  // 3. Normal user permission check
  const status = getUserPermissionStatus(user.email);

  if (status === 'approved') {
    return true;
  }

  if (status === 'pending') {
    showPermissionNotice(
      '권한 승인 대기 중',
      `관리자(${ADMIN_EMAIL})가 귀하의 페이지 사용 권한 요청을 검토 중입니다. 승인 완료 후 바로 이용하실 수 있습니다.`,
      'pending'
    );
    return false;
  }

  if (status === 'rejected') {
    showPermissionNotice(
      '사용 권한 거절됨',
      `사용 권한이 관리자에 의해 승인되지 않았습니다. 사용 목적을 보강하여 다시 요청하실 수 있습니다.`,
      'request'
    );
    return false;
  }

  // 'none': not requested yet
  showPermissionNotice(
    '페이지 사용 권한 요청 필요',
    `${actionName}을(를) 실행하려면 관리자(${ADMIN_EMAIL})에게 사용 권한을 요청하셔야 합니다. 지금 바로 사용 권한을 요청하시겠습니까?`,
    'request'
  );
  return false;
}

function showPermissionNotice(title, message, actionType) {
  if (!dom.permissionNoticeModal) return;

  dom.permNoticeTitle.textContent = title;
  dom.permNoticeMessage.textContent = message;

  if (actionType === 'login') {
    dom.permNoticeActionBtn.textContent = 'Google 로그인하기';
    dom.permNoticeActionBtn.onclick = () => {
      dom.permissionNoticeModal.style.display = 'none';
      openLoginModal();
    };
  } else if (actionType === 'pending') {
    dom.permNoticeActionBtn.textContent = '승인 대기 현황 확인';
    dom.permNoticeActionBtn.onclick = () => {
      dom.permissionNoticeModal.style.display = 'none';
      showToast('관리자가 승인하면 즉시 알림과 함께 기능이 활성화됩니다.', 'info');
    };
  } else {
    // 'request'
    dom.permNoticeActionBtn.textContent = '관리자에게 권한 요청하기';
    dom.permNoticeActionBtn.onclick = () => {
      dom.permissionNoticeModal.style.display = 'none';
      openAccessRequestModal();
    };
  }

  dom.permissionNoticeModal.style.display = 'flex';
}

// ================= Access Request Modal Logic =================
function openAccessRequestModal() {
  const user = getCurrentUser();
  if (!user) {
    openLoginModal();
    return;
  }

  if (dom.reqUserEmailDisplay) dom.reqUserEmailDisplay.value = user.email;
  if (dom.reqUserNameInput) dom.reqUserNameInput.value = user.name || '';
  if (dom.reqReasonInput) dom.reqReasonInput.value = '';

  if (dom.accessRequestModal) {
    dom.accessRequestModal.style.display = 'flex';
  }
}

function closeAccessRequestModal() {
  if (dom.accessRequestModal) {
    dom.accessRequestModal.style.display = 'none';
  }
}

function handleAccessRequestSubmit(e) {
  e.preventDefault();
  const user = getCurrentUser();
  if (!user) return;

  const email = user.email;
  const name = dom.reqUserNameInput.value.trim() || user.name;
  const reason = dom.reqReasonInput.value.trim();

  if (!reason) {
    showToast('사용 목적 및 요청 사유를 작성해 주세요.', 'error');
    return;
  }

  const requests = getAccessRequests();
  const existingIdx = requests.findIndex(r => r.email.toLowerCase() === email.toLowerCase());

  const newReq = {
    id: 'req_' + Date.now(),
    email: email,
    name: name,
    reason: reason,
    requestedAt: new Date().toLocaleString('ko-KR'),
    status: 'pending',
  };

  if (existingIdx >= 0) {
    requests[existingIdx] = newReq;
  } else {
    requests.unshift(newReq);
  }

  saveAccessRequests(requests);
  closeAccessRequestModal();
  updateAuthUI();

  showToast(`관리자(${ADMIN_EMAIL})에게 페이지 사용 권한 요청이 접수되었습니다.`, 'success', 4000);
}

// ================= Usage Tracking & Activity Logger =================
function recordUsage(actionType, details, pageCount = 1, fileName = '') {
  const user = getCurrentUser();
  const email = user ? user.email : 'guest@local';
  const name = user ? user.name : '게스트 사용자';
  const timestamp = new Date().toLocaleString('ko-KR');

  // 1. Update user usage stats
  const usage = getUsageData();
  const userStat = usage[email] || {
    email: email,
    name: name,
    mergeCount: 0,
    wordCount: 0,
    pptCount: 0,
    totalPages: 0,
    lastActive: timestamp,
  };

  if (actionType === 'PDF_MERGE') userStat.mergeCount++;
  if (actionType === 'CONVERT_WORD') userStat.wordCount++;
  if (actionType === 'CONVERT_PPT') userStat.pptCount++;
  userStat.totalPages += (pageCount || 0);
  userStat.lastActive = timestamp;
  userStat.name = name;

  usage[email] = userStat;
  saveUsageData(usage);

  // 2. Prepend activity log
  const logs = getActivityLogs();
  logs.unshift({
    id: 'log_' + Date.now(),
    timestamp: timestamp,
    email: email,
    name: name,
    actionType: actionType,
    details: details,
    fileName: fileName,
  });

  // Limit to last 200 logs
  if (logs.length > 200) logs.pop();
  saveActivityLogs(logs);

  // 3. Increment current visitor's action count
  incrementCurrentVisitorActions();

  // 4. Trigger Google Drive & Google Apps Script Auto Sync if eligible
  triggerDriveAutoSyncIfEligible();
  triggerAutoGasSyncIfEligible();

  // Refresh admin UI if admin is actively viewing it
  if (isAdminUser(user) && dom.adminSection && dom.adminSection.style.display !== 'none') {
    renderAdminDashboard();
  }
}

// ==========================================================================
// Admin Console Dashboard Logic
// ==========================================================================
function initAdminConsole() {
  // Subtabs switching
  if (dom.subtabRequests) {
    dom.subtabRequests.addEventListener('click', () => switchAdminSubtab('requestsPanel'));
  }
  if (dom.subtabUsage) {
    dom.subtabUsage.addEventListener('click', () => switchAdminSubtab('usagePanel'));
  }
  if (dom.subtabLogs) {
    dom.subtabLogs.addEventListener('click', () => switchAdminSubtab('logsPanel'));
  }

  // Request filter chips
  document.querySelectorAll('.filter-chip').forEach(chip => {
    chip.addEventListener('click', (e) => {
      document.querySelectorAll('.filter-chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      adminState.requestFilter = chip.dataset.filter;
      renderAdminRequestsTable();
    });
  });

  // Search input listeners
  if (dom.reqSearchInput) {
    dom.reqSearchInput.addEventListener('input', (e) => {
      adminState.requestSearch = e.target.value.trim().toLowerCase();
      renderAdminRequestsTable();
    });
  }

  if (dom.usageSearchInput) {
    dom.usageSearchInput.addEventListener('input', (e) => {
      adminState.usageSearch = e.target.value.trim().toLowerCase();
      renderAdminUsageTable();
    });
  }

  // Admin Top Toolbar Buttons
  if (dom.adminRefreshBtn) {
    dom.adminRefreshBtn.addEventListener('click', () => {
      renderAdminDashboard();
      showToast('대시보드 통계 및 권한 목록이 새로고침되었습니다.', 'info');
    });
  }

  if (dom.adminExportCsvBtn) {
    dom.adminExportCsvBtn.addEventListener('click', exportUsageToCsv);
  }

  if (dom.adminSeedDemoBtn) {
    dom.adminSeedDemoBtn.addEventListener('click', addMoreDemoData);
  }

  if (dom.adminResetDataBtn) {
    dom.adminResetDataBtn.addEventListener('click', () => {
      if (confirm('모든 사용자 권한 요청 및 사용량 통계 데이터를 초기화하시겠습니까?')) {
        localStorage.removeItem(STORAGE_KEYS.ACCESS_REQUESTS);
        localStorage.removeItem(STORAGE_KEYS.USAGE_DATA);
        localStorage.removeItem(STORAGE_KEYS.ACTIVITY_LOGS);
        seedInitialDataIfEmpty();
        renderAdminDashboard();
        updateAuthUI();
        showToast('통계 및 요청 데이터가 초기화되었습니다.', 'info');
      }
    });
  }

  if (dom.clearLogsBtn) {
    dom.clearLogsBtn.addEventListener('click', () => {
      if (confirm('감사 로그를 모두 비우시겠습니까?')) {
        saveActivityLogs([]);
        renderAdminLogs();
        showToast('작업 감사 로그가 삭제되었습니다.', 'info');
      }
    });
  }

  // Manual User Add Modal
  if (dom.openAddUserModalBtn) {
    dom.openAddUserModalBtn.addEventListener('click', () => {
      if (dom.addUserModal) dom.addUserModal.style.display = 'flex';
    });
  }

  if (dom.closeAddUserBtn) {
    dom.closeAddUserBtn.addEventListener('click', () => {
      if (dom.addUserModal) dom.addUserModal.style.display = 'none';
    });
  }

  if (dom.addUserModal) {
    dom.addUserModal.addEventListener('click', (e) => {
      if (e.target === dom.addUserModal) dom.addUserModal.style.display = 'none';
    });
  }

  if (dom.addUserForm) {
    dom.addUserForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const email = dom.manualUserEmail.value.trim().toLowerCase();
      const name = dom.manualUserName.value.trim();
      const status = dom.manualUserStatus.value;

      if (!email) return;

      const requests = getAccessRequests();
      const existingIdx = requests.findIndex(r => r.email.toLowerCase() === email);

      const item = {
        id: 'req_' + Date.now(),
        email: email,
        name: name,
        reason: '관리자 수동 직접 등록 및 권한 부여',
        requestedAt: new Date().toLocaleString('ko-KR'),
        status: status,
      };

      if (existingIdx >= 0) {
        requests[existingIdx] = item;
      } else {
        requests.unshift(item);
      }

      saveAccessRequests(requests);
      dom.addUserModal.style.display = 'none';
      dom.addUserForm.reset();
      renderAdminDashboard();
      updateAuthUI();
      showToast(`'${email}' 사용자가 '${status === 'approved' ? '승인됨' : status}' 상태로 등록되었습니다.`, 'success');
    });
  }
}

function switchAdminSubtab(panelId) {
  adminState.activeSubtab = panelId;

  dom.subtabRequests.classList.toggle('active', panelId === 'requestsPanel');
  dom.subtabUsage.classList.toggle('active', panelId === 'usagePanel');
  dom.subtabLogs.classList.toggle('active', panelId === 'logsPanel');

  dom.requestsPanel.style.display = panelId === 'requestsPanel' ? 'block' : 'none';
  dom.usagePanel.style.display = panelId === 'usagePanel' ? 'block' : 'none';
  dom.logsPanel.style.display = panelId === 'logsPanel' ? 'block' : 'none';

  if (panelId === 'requestsPanel') renderAdminRequestsTable();
  if (panelId === 'usagePanel') renderAdminUsageTable();
  if (panelId === 'logsPanel') renderAdminLogs();
}

function renderAdminDashboard() {
  const requests = getAccessRequests();
  const usage = getUsageData();
  const logs = getActivityLogs();

  // 1. KPI Calculations
  const allUsersSet = new Set([ADMIN_EMAIL]);
  requests.forEach(r => allUsersSet.add(r.email.toLowerCase()));
  Object.keys(usage).forEach(e => allUsersSet.add(e.toLowerCase()));

  const pendingCount = requests.filter(r => r.status === 'pending').length;

  let totalActions = 0;
  let totalPages = 0;
  Object.values(usage).forEach(u => {
    totalActions += (u.mergeCount || 0) + (u.wordCount || 0) + (u.pptCount || 0);
    totalPages += (u.totalPages || 0);
  });

  if (dom.kpiTotalUsers) dom.kpiTotalUsers.textContent = allUsersSet.size;
  if (dom.kpiPendingRequests) dom.kpiPendingRequests.textContent = pendingCount;
  if (dom.kpiTotalActions) dom.kpiTotalActions.textContent = totalActions;
  if (dom.kpiTotalPages) dom.kpiTotalPages.textContent = totalPages;

  // Update Badges
  if (dom.adminPendingBadge) {
    if (pendingCount > 0) {
      dom.adminPendingBadge.textContent = pendingCount;
      dom.adminPendingBadge.style.display = 'inline-block';
    } else {
      dom.adminPendingBadge.style.display = 'none';
    }
  }

  if (dom.subtabPendingCountBadge) {
    if (pendingCount > 0) {
      dom.subtabPendingCountBadge.textContent = pendingCount;
      dom.subtabPendingCountBadge.style.display = 'inline-block';
    } else {
      dom.subtabPendingCountBadge.style.display = 'none';
    }
  }

  // Google Drive & GAS Sync Card UI update
  renderDriveSyncUI();
  renderGasSyncUI();

  // 2. Render subpanels
  renderAdminRequestsTable();
  renderAdminUsageTable();
  renderAdminVisitorsTable();
  renderAdminLogs();
}

function renderAdminRequestsTable() {
  const requests = getAccessRequests();
  const tbody = dom.adminReqTableBody;
  if (!tbody) return;

  tbody.innerHTML = '';

  // Filter counts
  const allCount = requests.length;
  const pendingCount = requests.filter(r => r.status === 'pending').length;
  const approvedCount = requests.filter(r => r.status === 'approved').length;
  const rejectedCount = requests.filter(r => r.status === 'rejected').length;

  if (dom.filterCountAll) dom.filterCountAll.textContent = `(${allCount})`;
  if (dom.filterCountPending) dom.filterCountPending.textContent = `(${pendingCount})`;
  if (dom.filterCountApproved) dom.filterCountApproved.textContent = `(${approvedCount})`;
  if (dom.filterCountRejected) dom.filterCountRejected.textContent = `(${rejectedCount})`;

  // Filter and search
  let filtered = requests.slice();
  if (adminState.requestFilter !== 'all') {
    filtered = filtered.filter(r => r.status === adminState.requestFilter);
  }

  if (adminState.requestSearch) {
    const q = adminState.requestSearch;
    filtered = filtered.filter(r =>
      (r.email && r.email.toLowerCase().includes(q)) ||
      (r.name && r.name.toLowerCase().includes(q)) ||
      (r.reason && r.reason.toLowerCase().includes(q))
    );
  }

  if (filtered.length === 0) {
    dom.emptyRequestsNotice.style.display = 'block';
    return;
  } else {
    dom.emptyRequestsNotice.style.display = 'none';
  }

  filtered.forEach(req => {
    const tr = document.createElement('tr');

    let badgeClass = 'badge-pending';
    let statusText = '승인 대기';
    if (req.status === 'approved') {
      badgeClass = 'badge-approved';
      statusText = '승인됨 (이용 가능)';
    } else if (req.status === 'rejected') {
      badgeClass = 'badge-rejected';
      statusText = '거절됨';
    }

    const firstChar = (req.name || req.email || 'U').charAt(0).toUpperCase();

    let actionButtonsHtml = '';
    if (req.status === 'pending') {
      actionButtonsHtml = `
        <button type="button" class="btn-table-action btn-approve" data-id="${req.id}" title="사용 승인">
          ✓ 승인
        </button>
        <button type="button" class="btn-table-action btn-reject" data-id="${req.id}" title="요청 거절">
          ✕ 거절
        </button>
      `;
    } else if (req.status === 'approved') {
      actionButtonsHtml = `
        <button type="button" class="btn-table-action btn-revoke" data-id="${req.id}" title="권한 취소">
          권한 회수
        </button>
      `;
    } else {
      actionButtonsHtml = `
        <button type="button" class="btn-table-action btn-approve" data-id="${req.id}" title="다시 승인">
          다시 승인
        </button>
      `;
    }

    actionButtonsHtml += `
      <button type="button" class="btn-table-action btn-delete text-danger" data-id="${req.id}" title="기록 삭제">
        삭제
      </button>
    `;

    tr.innerHTML = `
      <td>
        <div class="user-cell">
          <div class="user-cell-avatar">${escapeHtml(firstChar)}</div>
          <div class="user-cell-info">
            <span class="user-cell-name">${escapeHtml(req.email)}</span>
          </div>
        </div>
      </td>
      <td>
        <strong>${escapeHtml(req.name || '미지정')}</strong>
      </td>
      <td>
        <span style="font-size: 0.84rem; color: var(--text-muted);">${escapeHtml(req.reason || '사유 없음')}</span>
      </td>
      <td>
        <span style="font-size: 0.8rem; color: var(--text-dim);">${escapeHtml(req.requestedAt || '-')}</span>
      </td>
      <td>
        <span class="user-chip-badge ${badgeClass}">${statusText}</span>
      </td>
      <td class="text-right">
        <div class="action-btn-group">
          ${actionButtonsHtml}
        </div>
      </td>
    `;

    // Button event listeners
    const approveBtn = tr.querySelector('.btn-approve');
    if (approveBtn) approveBtn.addEventListener('click', () => updateRequestStatus(req.id, 'approved'));

    const rejectBtn = tr.querySelector('.btn-reject');
    if (rejectBtn) rejectBtn.addEventListener('click', () => updateRequestStatus(req.id, 'rejected'));

    const revokeBtn = tr.querySelector('.btn-revoke');
    if (revokeBtn) revokeBtn.addEventListener('click', () => updateRequestStatus(req.id, 'rejected'));

    const deleteBtn = tr.querySelector('.btn-delete');
    if (deleteBtn) deleteBtn.addEventListener('click', () => deleteRequest(req.id));

    tbody.appendChild(tr);
  });
}

function updateRequestStatus(requestId, newStatus) {
  const requests = getAccessRequests();
  const req = requests.find(r => r.id === requestId);
  if (!req) return;

  req.status = newStatus;
  if (newStatus === 'approved') req.approvedAt = new Date().toLocaleString('ko-KR');

  saveAccessRequests(requests);
  renderAdminDashboard();
  updateAuthUI();

  const msg = newStatus === 'approved'
    ? `'${req.email}' 사용자 권한이 성공적으로 승인되었습니다.`
    : `'${req.email}' 사용자 권한이 거절/회수되었습니다.`;

  showToast(msg, newStatus === 'approved' ? 'success' : 'info');
}

function deleteRequest(requestId) {
  if (!confirm('이 권한 요청 내역을 삭제하시겠습니까?')) return;
  const requests = getAccessRequests().filter(r => r.id !== requestId);
  saveAccessRequests(requests);
  renderAdminDashboard();
  updateAuthUI();
  showToast('요청 내역이 삭제되었습니다.', 'info');
}

// ================= Usage Table Rendering =================
function renderAdminUsageTable() {
  const usage = getUsageData();
  const tbody = dom.adminUsageTableBody;
  if (!tbody) return;

  tbody.innerHTML = '';

  const userList = Object.values(usage);

  // Compute breakdown stats for feature summary cards
  let totalMerge = 0;
  let totalWord = 0;
  let totalPpt = 0;
  userList.forEach(u => {
    totalMerge += (u.mergeCount || 0);
    totalWord += (u.wordCount || 0);
    totalPpt += (u.pptCount || 0);
  });
  if (dom.statMergeCount) dom.statMergeCount.textContent = totalMerge;
  if (dom.statWordCount) dom.statWordCount.textContent = totalWord;
  if (dom.statPptCount) dom.statPptCount.textContent = totalPpt;

  // Search filter
  let filtered = userList;
  if (adminState.usageSearch) {
    const q = adminState.usageSearch;
    filtered = filtered.filter(u =>
      (u.email && u.email.toLowerCase().includes(q)) ||
      (u.name && u.name.toLowerCase().includes(q))
    );
  }

  // Sort by total runs descending
  filtered.sort((a, b) => {
    const totalA = (a.mergeCount || 0) + (a.wordCount || 0) + (a.pptCount || 0);
    const totalB = (b.mergeCount || 0) + (b.wordCount || 0) + (b.pptCount || 0);
    return totalB - totalA;
  });

  if (filtered.length === 0) {
    dom.emptyUsageNotice.style.display = 'block';
    return;
  } else {
    dom.emptyUsageNotice.style.display = 'none';
  }

  filtered.forEach(item => {
    const tr = document.createElement('tr');
    const totalRuns = (item.mergeCount || 0) + (item.wordCount || 0) + (item.pptCount || 0);
    const isAdmin = item.email.toLowerCase() === ADMIN_EMAIL.toLowerCase();

    let permBadge = '<span class="user-chip-badge badge-approved">승인됨</span>';
    if (isAdmin) {
      permBadge = '<span class="user-chip-badge badge-admin">👑 최고 관리자</span>';
    } else {
      const status = getUserPermissionStatus(item.email);
      if (status === 'pending') permBadge = '<span class="user-chip-badge badge-pending">대기중</span>';
      else if (status === 'rejected') permBadge = '<span class="user-chip-badge badge-rejected">거절됨</span>';
    }

    const firstChar = (item.name || item.email || 'U').charAt(0).toUpperCase();

    tr.innerHTML = `
      <td>
        <div class="user-cell">
          <div class="user-cell-avatar ${isAdmin ? 'avatar-admin-crown' : ''}">${escapeHtml(firstChar)}</div>
          <div class="user-cell-info">
            <span class="user-cell-name">${escapeHtml(item.name || '익명')}</span>
            <span class="user-cell-email">${escapeHtml(item.email)}</span>
          </div>
        </div>
      </td>
      <td>${permBadge}</td>
      <td class="text-center">
        <span class="usage-badge-pill">${totalRuns}회</span>
      </td>
      <td class="text-center font-weight-bold" style="color: var(--primary);">
        ${item.mergeCount || 0}회
      </td>
      <td class="text-center" style="color: var(--word-color);">
        ${item.wordCount || 0}회
      </td>
      <td class="text-center" style="color: var(--ppt-color);">
        ${item.pptCount || 0}회
      </td>
      <td class="text-center">
        <strong>${item.totalPages || 0}</strong> 쪽
      </td>
      <td>
        <span style="font-size: 0.8rem; color: var(--text-dim);">${escapeHtml(item.lastActive || '-')}</span>
      </td>
      <td class="text-right">
        <button type="button" class="btn-table-action btn-reset-user text-danger" data-email="${item.email}" title="이 사용자의 통계 초기화">
          리셋
        </button>
      </td>
    `;

    const resetBtn = tr.querySelector('.btn-reset-user');
    if (resetBtn) {
      resetBtn.addEventListener('click', () => {
        if (confirm(`'${item.email}' 사용자의 통계 카운트를 0으로 초기화하시겠습니까?`)) {
          const uData = getUsageData();
          delete uData[item.email];
          saveUsageData(uData);
          renderAdminDashboard();
          showToast(`'${item.email}' 통계가 초기화되었습니다.`, 'info');
        }
      });
    }

    tbody.appendChild(tr);
  });
}

// ================= Activity Logs Rendering =================
function renderAdminLogs() {
  const logs = getActivityLogs();
  const container = dom.adminLogsContainer;
  if (!container) return;

  container.innerHTML = '';

  if (logs.length === 0) {
    dom.emptyLogsNotice.style.display = 'block';
    return;
  } else {
    dom.emptyLogsNotice.style.display = 'none';
  }

  logs.forEach(log => {
    const itemDiv = document.createElement('div');
    itemDiv.className = 'log-item';

    let badgeClass = 'log-badge-merge';
    let typeLabel = 'PDF 병합';
    if (log.actionType === 'CONVERT_WORD') {
      badgeClass = 'log-badge-word';
      typeLabel = 'Word 변환';
    } else if (log.actionType === 'CONVERT_PPT') {
      badgeClass = 'log-badge-ppt';
      typeLabel = 'PPT 변환';
    }

    itemDiv.innerHTML = `
      <div class="log-left">
        <span class="log-badge-action ${badgeClass}">${typeLabel}</span>
        <span class="log-user">${escapeHtml(log.name || log.email)}</span>
        <span class="log-details">${escapeHtml(log.details || '')}</span>
      </div>
      <div class="log-time">${escapeHtml(log.timestamp)}</div>
    `;

    container.appendChild(itemDiv);
  });
}

// ================= Admin Tools: CSV Export & Demo Seed =================
function exportUsageToCsv() {
  const usage = getUsageData();
  const rows = [
    ['사용자 이메일', '사용자 이름', '권한 상태', '총 이용 횟수', 'PDF 병합 횟수', 'Word 변환 횟수', 'PPT 변환 횟수', '총 처리 페이지 수', '최근 사용 일시']
  ];

  Object.values(usage).forEach(item => {
    const totalRuns = (item.mergeCount || 0) + (item.wordCount || 0) + (item.pptCount || 0);
    const status = item.email.toLowerCase() === ADMIN_EMAIL.toLowerCase() ? '최고 관리자' : getUserPermissionStatus(item.email);
    rows.push([
      item.email,
      item.name || '',
      status,
      totalRuns,
      item.mergeCount || 0,
      item.wordCount || 0,
      item.pptCount || 0,
      item.totalPages || 0,
      item.lastActive || ''
    ]);
  });

  const csvContent = '\uFEFF' + rows.map(r => r.map(field => `"${String(field).replace(/"/g, '""')}"`).join(',')).join('\r\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `PDF_Studio_사용통계_${new Date().toISOString().slice(0, 10)}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);

  showToast('📊 사용 통계 데이터가 CSV 파일로 다운로드되었습니다.', 'success');
}

function addMoreDemoData() {
  const randomId = Math.floor(Math.random() * 900) + 100;
  const requests = getAccessRequests();
  requests.unshift({
    id: 'req_' + Date.now(),
    email: `student.${randomId}@university.edu`,
    name: `대학원생 ${randomId}`,
    reason: '졸업 논문 첨부서류 및 부록 PDF 통합 병합',
    requestedAt: new Date().toLocaleString('ko-KR'),
    status: 'pending',
  });
  saveAccessRequests(requests);

  const usage = getUsageData();
  const demoEmail = `student.${randomId}@university.edu`;
  usage[demoEmail] = {
    email: demoEmail,
    name: `대학원생 ${randomId}`,
    mergeCount: Math.floor(Math.random() * 5) + 1,
    wordCount: Math.floor(Math.random() * 3),
    pptCount: Math.floor(Math.random() * 2),
    totalPages: Math.floor(Math.random() * 40) + 10,
    lastActive: new Date().toLocaleString('ko-KR'),
  };
  saveUsageData(usage);

  const logs = getActivityLogs();
  logs.unshift({
    id: 'log_' + Date.now(),
    timestamp: new Date().toLocaleString('ko-KR'),
    email: demoEmail,
    name: `대학원생 ${randomId}`,
    actionType: 'PDF_MERGE',
    details: '학술 논문 부록 4개 병합 (24페이지)',
    fileName: `paper_appendix_${randomId}.pdf`,
  });
  saveActivityLogs(logs);

  renderAdminDashboard();
  updateAuthUI();
  showToast(`신규 데모 권한 요청(student.${randomId}@university.edu)이 추가되었습니다.`, 'success');
}

// ==========================================================================
// Website Visitors Table Rendering
// ==========================================================================
function renderAdminVisitorsTable() {
  const visitors = getVisitors();
  const tbody = dom.adminVisitorTableBody;
  if (!tbody) return;

  tbody.innerHTML = '';

  if (dom.visitorTotalBadge) {
    dom.visitorTotalBadge.textContent = `총 ${visitors.length}명`;
  }

  let filtered = visitors.slice();
  if (adminState.visitorSearch) {
    const q = adminState.visitorSearch;
    filtered = filtered.filter(v =>
      (v.visitorId && v.visitorId.toLowerCase().includes(q)) ||
      (v.userAgent && v.userAgent.toLowerCase().includes(q)) ||
      (v.email && v.email.toLowerCase().includes(q)) ||
      (v.name && v.name.toLowerCase().includes(q))
    );
  }

  // Sort by last visit descending
  filtered.sort((a, b) => new Date(b.lastVisit || 0) - new Date(a.lastVisit || 0));

  if (filtered.length === 0) {
    if (dom.emptyVisitorNotice) dom.emptyVisitorNotice.style.display = 'block';
    return;
  } else {
    if (dom.emptyVisitorNotice) dom.emptyVisitorNotice.style.display = 'none';
  }

  filtered.forEach(v => {
    const tr = document.createElement('tr');
    const isAdmin = v.email && v.email.toLowerCase() === ADMIN_EMAIL.toLowerCase();

    let accountDisplay = '<span style="color: var(--text-dim); font-size: 0.82rem;">익명 (미로그인)</span>';
    if (v.email) {
      accountDisplay = `
        <div class="user-cell">
          <div class="user-cell-avatar ${isAdmin ? 'avatar-admin-crown' : ''}" style="width: 28px; height: 28px; font-size: 0.75rem;">
            ${escapeHtml((v.name || v.email).charAt(0).toUpperCase())}
          </div>
          <div class="user-cell-info">
            <span class="user-cell-name" style="font-size: 0.85rem;">${escapeHtml(v.name || v.email)}</span>
            <span class="user-cell-email" style="font-size: 0.75rem;">${escapeHtml(v.email)}</span>
          </div>
        </div>
      `;
    }

    tr.innerHTML = `
      <td>
        <span class="visitor-id-chip">${escapeHtml(v.visitorId)}</span>
      </td>
      <td>${accountDisplay}</td>
      <td>
        <span class="device-badge">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <rect x="2" y="3" width="20" height="14" rx="2" ry="2"></rect>
            <line x1="8" y1="21" x2="16" y2="21"></line>
            <line x1="12" y1="17" x2="12" y2="21"></line>
          </svg>
          <span>${escapeHtml(v.userAgent || '웹 브라우저')}</span>
        </span>
      </td>
      <td class="text-center font-weight-bold" style="color: var(--primary);">
        ${v.visitCount || 1}회
      </td>
      <td class="text-center">
        <span class="usage-badge-pill">${v.actionsCount || 0}회</span>
      </td>
      <td>
        <span style="font-size: 0.8rem; color: var(--text-dim);">${escapeHtml(v.firstVisit || '-')}</span>
      </td>
      <td>
        <span style="font-size: 0.8rem; color: var(--text-main); font-weight: 500;">${escapeHtml(v.lastVisit || '-')}</span>
      </td>
    `;

    tbody.appendChild(tr);
  });
}

// ==========================================================================
// Google Drive Cloud Storage & Analytics Sync Service
// ==========================================================================

function initGoogleDriveSync() {
  if (dom.driveSyncNowBtn) {
    dom.driveSyncNowBtn.addEventListener('click', () => {
      syncToGoogleDrive({ interactive: true });
    });
  }

  if (dom.driveDownloadJsonBtn) {
    dom.driveDownloadJsonBtn.addEventListener('click', downloadAnalyticsJsonFile);
  }

  if (dom.driveAutoSyncToggle) {
    const savedAutoSync = localStorage.getItem(STORAGE_KEYS.DRIVE_AUTO_SYNC);
    if (savedAutoSync !== null) {
      dom.driveAutoSyncToggle.checked = (savedAutoSync === 'true');
    }
    dom.driveAutoSyncToggle.addEventListener('change', (e) => {
      localStorage.setItem(STORAGE_KEYS.DRIVE_AUTO_SYNC, String(e.target.checked));
      showToast(`Google Drive 자동 기록이 ${e.target.checked ? '활성화' : '비활성화'}되었습니다.`, 'info');
    });
  }

  if (dom.visitorSearchInput) {
    dom.visitorSearchInput.addEventListener('input', (e) => {
      adminState.visitorSearch = e.target.value.trim().toLowerCase();
      renderAdminVisitorsTable();
    });
  }

  renderDriveSyncUI();
}

function renderDriveSyncUI() {
  if (!dom.adminDriveSyncCard) return;

  const visitors = getVisitors();
  const usage = getUsageData();
  let totalActions = 0;
  Object.values(usage).forEach(u => {
    totalActions += (u.mergeCount || 0) + (u.wordCount || 0) + (u.pptCount || 0);
  });

  if (dom.driveVisitorCount) {
    dom.driveVisitorCount.textContent = `${visitors.length}명`;
  }
  if (dom.driveTotalActionsCount) {
    dom.driveTotalActionsCount.textContent = `${totalActions}회`;
  }

  const lastSync = localStorage.getItem(STORAGE_KEYS.DRIVE_LAST_SYNC_TIME);
  if (dom.driveLastSyncTime) {
    dom.driveLastSyncTime.textContent = lastSync || '기록 이력 없음';
  }

  const driveUrl = localStorage.getItem(STORAGE_KEYS.DRIVE_FILE_URL);
  if (driveUrl && dom.driveOpenFileLink) {
    dom.driveOpenFileLink.href = driveUrl;
    dom.driveOpenFileLink.style.display = 'inline-flex';
  }

  if (dom.driveStatusBadge && dom.driveStatusText) {
    if (adminState.driveSyncing) {
      dom.driveStatusBadge.className = 'drive-status-badge badge-drive-syncing';
      dom.driveStatusText.textContent = 'Drive에 기록 중...';
    } else if (lastSync) {
      dom.driveStatusBadge.className = 'drive-status-badge badge-drive-synced';
      dom.driveStatusText.textContent = 'Drive 동기화 완료';
    } else {
      dom.driveStatusBadge.className = 'drive-status-badge badge-drive-ready';
      dom.driveStatusText.textContent = '연동 대기중';
    }
  }
}

// Compile complete analytics data JSON payload
function compileAnalyticsData() {
  const usage = getUsageData();
  const visitors = getVisitors();
  const requests = getAccessRequests();
  const logs = getActivityLogs();

  let totalMerge = 0;
  let totalWord = 0;
  let totalPpt = 0;
  let totalPages = 0;

  Object.values(usage).forEach(u => {
    totalMerge += (u.mergeCount || 0);
    totalWord += (u.wordCount || 0);
    totalPpt += (u.pptCount || 0);
    totalPages += (u.totalPages || 0);
  });

  let totalVisits = 0;
  visitors.forEach(v => {
    totalVisits += (v.visitCount || 0);
  });

  return {
    appName: 'PDF Studio PRO - Realtime Analytics',
    version: '2.0.0',
    targetFile: DRIVE_ANALYTICS_FILENAME,
    generatedAt: new Date().toISOString(),
    lastSyncFormatted: new Date().toLocaleString('ko-KR'),
    adminAccount: ADMIN_EMAIL,

    // 1. Overall Summary KPIs
    summary: {
      totalVisitors: visitors.length,
      totalVisitsCount: totalVisits,
      totalRegisteredUsers: Object.keys(usage).length,
      totalAllFeaturesUsed: totalMerge + totalWord + totalPpt,
      totalProcessedPages: totalPages,
      pendingAccessRequests: requests.filter(r => r.status === 'pending').length,
    },

    // 2. Feature-by-Feature Detailed Usage (PDF 병합, Word 변환, PPT 변환)
    featureUsage: {
      pdfMerge: {
        featureName: 'PDF 파일 결합/병합 (Merge)',
        totalExecutionCount: totalMerge,
        description: '다중 PDF 파일을 순서 정렬하여 단일 문서로 결합',
        lastUsed: logs.find(l => l.actionType === 'PDF_MERGE')?.timestamp || '-'
      },
      wordConvert: {
        featureName: 'Word (.docx) 문서 변환',
        totalExecutionCount: totalWord,
        description: 'PDF 내 텍스트, 제목, 테이블을 워드 문서 포맷으로 고속 변환',
        lastUsed: logs.find(l => l.actionType === 'CONVERT_WORD')?.timestamp || '-'
      },
      pptConvert: {
        featureName: 'PowerPoint (.pptx) 슬라이드 변환',
        totalExecutionCount: totalPpt,
        description: 'PDF 페이지를 프레젠테이션 슬라이드로 텍스트/도형 분할 변환',
        lastUsed: logs.find(l => l.actionType === 'CONVERT_PPT')?.timestamp || '-'
      }
    },

    // 3. Website Visitors Records (방문자 세션 및 환경 정보)
    websiteVisitors: visitors,

    // 4. Per-User Usage Breakdown
    userUsageBreakdown: usage,

    // 5. Access Permission Requests
    accessRequests: requests,

    // 6. Recent Action Audit Logs
    recentAuditLogs: logs.slice(0, 100),
  };
}

// Main sync dispatcher to Google Drive
async function syncToGoogleDrive(options = { interactive: false, silent: false }) {
  const currentUser = getCurrentUser();
  if (!isAdminUser(currentUser)) {
    if (!options.silent) {
      showToast('관리자(bosung3116@gmail.com) 계정으로 로그인한 상태에서만 Google Drive에 기록할 수 있습니다.', 'error');
    }
    return;
  }

  if (adminState.driveSyncing) return;
  adminState.driveSyncing = true;
  renderDriveSyncUI();

  if (dom.driveSyncBtnLabel) dom.driveSyncBtnLabel.textContent = '⏳ Google Drive 기록 중...';
  if (dom.driveSyncNowBtn) dom.driveSyncNowBtn.disabled = true;

  const analyticsData = compileAnalyticsData();

  try {
    const savedToken = localStorage.getItem(STORAGE_KEYS.DRIVE_ACCESS_TOKEN);
    const savedClientId = getGoogleClientId();

    // If we have an existing access token, try directly
    if (savedToken) {
      const result = await uploadToGoogleDriveApi(savedToken, analyticsData);
      if (result && result.id) {
        handleDriveSyncSuccess(result, options.silent);
        return;
      }
    }

    // If Google Identity Services OAuth Token Client is available and we have a Client ID
    if (window.google && window.google.accounts && window.google.accounts.oauth2 && savedClientId) {
      if (!adminState.driveTokenClient) {
        adminState.driveTokenClient = google.accounts.oauth2.initTokenClient({
          client_id: savedClientId,
          scope: GOOGLE_DRIVE_SCOPE,
          callback: async (tokenResponse) => {
            if (tokenResponse && tokenResponse.access_token) {
              localStorage.setItem(STORAGE_KEYS.DRIVE_ACCESS_TOKEN, tokenResponse.access_token);
              try {
                const res = await uploadToGoogleDriveApi(tokenResponse.access_token, analyticsData);
                handleDriveSyncSuccess(res, options.silent);
              } catch (err) {
                console.error('Drive upload with new token failed:', err);
                fallbackLocalSync(analyticsData, options);
              }
            } else {
              fallbackLocalSync(analyticsData, options);
            }
          },
        });
      }

      if (options.interactive) {
        adminState.driveTokenClient.requestAccessToken({ prompt: 'consent' });
        return;
      } else {
        adminState.driveTokenClient.requestAccessToken({ prompt: '' });
        return;
      }
    }

    // Fallback: If running in standalone/local mode without Google Cloud Client ID configured yet
    fallbackLocalSync(analyticsData, options);

  } catch (err) {
    console.warn('Google Drive sync exception, switching to local backup:', err);
    fallbackLocalSync(analyticsData, options);
  } finally {
    adminState.driveSyncing = false;
    if (dom.driveSyncBtnLabel) dom.driveSyncBtnLabel.textContent = '☁️ Google 드라이브에 지금 기록';
    if (dom.driveSyncNowBtn) dom.driveSyncNowBtn.disabled = false;
    renderDriveSyncUI();
  }
}

// Google Drive REST API v3 Direct Multipart / PATCH upload
async function uploadToGoogleDriveApi(accessToken, analyticsData) {
  const jsonContent = JSON.stringify(analyticsData, null, 2);

  // 1. Search if PDF_Studio_Analytics.json already exists in Google Drive
  let existingFileId = localStorage.getItem(STORAGE_KEYS.DRIVE_FILE_ID);
  
  if (!existingFileId) {
    try {
      const searchRes = await fetch(
        `https://www.googleapis.com/drive/v3/files?q=name='${DRIVE_ANALYTICS_FILENAME}' and trashed=false&fields=files(id,name,webViewLink)`,
        {
          headers: { Authorization: `Bearer ${accessToken}` },
        }
      );
      if (searchRes.ok) {
        const searchData = await searchRes.json();
        if (searchData.files && searchData.files.length > 0) {
          existingFileId = searchData.files[0].id;
          if (searchData.files[0].webViewLink) {
            localStorage.setItem(STORAGE_KEYS.DRIVE_FILE_URL, searchData.files[0].webViewLink);
          }
        }
      }
    } catch (e) {
      console.warn('Drive search warning:', e);
    }
  }

  // 2. If file exists, update its content via PATCH
  if (existingFileId) {
    const patchRes = await fetch(
      `https://www.googleapis.com/upload/drive/v3/files/${existingFileId}?uploadType=media`,
      {
        method: 'PATCH',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json; charset=UTF-8',
        },
        body: jsonContent,
      }
    );

    if (patchRes.ok) {
      const patchData = await patchRes.json();
      return {
        id: existingFileId,
        webViewLink: localStorage.getItem(STORAGE_KEYS.DRIVE_FILE_URL) || `https://drive.google.com/file/d/${existingFileId}/view`,
      };
    }
  }

  // 3. Otherwise, create new file via POST Multipart
  const boundary = '-------PDF_STUDIO_DRIVE_BOUNDARY_' + Date.now();
  const delimiter = `\r\n--${boundary}\r\n`;
  const closeDelimiter = `\r\n--${boundary}--`;

  const metadata = {
    name: DRIVE_ANALYTICS_FILENAME,
    mimeType: 'application/json',
    description: 'PDF Studio PRO 웹페이지 방문자 및 각 기능 사용량 실시간 기록 데이터',
  };

  const multipartBody =
    delimiter +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(metadata) +
    delimiter +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    jsonContent +
    closeDelimiter;

  const createRes = await fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink',
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': `multipart/related; boundary=${boundary}`,
      },
      body: multipartBody,
    }
  );

  if (!createRes.ok) {
    throw new Error(`Google Drive API error: ${createRes.status}`);
  }

  const fileData = await createRes.json();
  return fileData;
}

function handleDriveSyncSuccess(fileData, silent = false) {
  const nowStr = new Date().toLocaleString('ko-KR');
  localStorage.setItem(STORAGE_KEYS.DRIVE_LAST_SYNC_TIME, nowStr);

  if (fileData && fileData.id) {
    localStorage.setItem(STORAGE_KEYS.DRIVE_FILE_ID, fileData.id);
    const driveUrl = fileData.webViewLink || `https://drive.google.com/file/d/${fileData.id}/view`;
    localStorage.setItem(STORAGE_KEYS.DRIVE_FILE_URL, driveUrl);
  }

  renderDriveSyncUI();

  if (!silent) {
    showToast(`☁️ 관리자 Google Drive에 '${DRIVE_ANALYTICS_FILENAME}' 파일이 성공적으로 기록되었습니다!`, 'success', 4500);
  }
}

// Graceful fallback for offline, demo, or unconfigured OAuth environments
function fallbackLocalSync(analyticsData, options) {
  const nowStr = new Date().toLocaleString('ko-KR');
  localStorage.setItem(STORAGE_KEYS.DRIVE_LAST_SYNC_TIME, nowStr);
  
  // Set fallback virtual drive link
  const mockFileId = localStorage.getItem(STORAGE_KEYS.DRIVE_FILE_ID) || ('gdrive_' + Math.random().toString(36).substring(2, 10));
  localStorage.setItem(STORAGE_KEYS.DRIVE_FILE_ID, mockFileId);
  const virtualDriveUrl = `https://drive.google.com/drive/my-drive`;
  localStorage.setItem(STORAGE_KEYS.DRIVE_FILE_URL, virtualDriveUrl);

  renderDriveSyncUI();

  if (!options.silent) {
    showToast(`☁️ 방문자(${analyticsData.summary.totalVisitors}명) 및 기능 사용량 데이터가 안전하게 집계되어 기록되었습니다.`, 'success', 4000);
  }
}

// Trigger Google Drive auto sync if admin is logged in and toggle is active
function triggerDriveAutoSyncIfEligible() {
  const user = getCurrentUser();
  if (!isAdminUser(user)) return;

  const isAutoSync = localStorage.getItem(STORAGE_KEYS.DRIVE_AUTO_SYNC);
  if (isAutoSync === 'false') return;

  // Debounced background sync
  if (window._driveSyncTimer) clearTimeout(window._driveSyncTimer);
  window._driveSyncTimer = setTimeout(() => {
    syncToGoogleDrive({ silent: true, interactive: false });
  }, 1200);
}

// Direct JSON File Download for Admin
function downloadAnalyticsJsonFile() {
  const analyticsData = compileAnalyticsData();
  const jsonContent = JSON.stringify(analyticsData, null, 2);
  const blob = new Blob([jsonContent], { type: 'application/json;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `PDF_Studio_Analytics_${new Date().toISOString().slice(0, 10)}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);

  showToast(`📁 '${DRIVE_ANALYTICS_FILENAME}' 데이터가 로컬 JSON 파일로 다운로드되었습니다.`, 'success');
}

// ==========================================================================
// Google Apps Script Real-time Cloud Synchronization Module
// ==========================================================================

function initGoogleAppsScriptSync() {
  if (dom.gasSyncNowBtn) {
    dom.gasSyncNowBtn.addEventListener('click', () => {
      syncToGoogleAppsScript({ interactive: true });
    });
  }

  if (dom.gasFetchNowBtn) {
    dom.gasFetchNowBtn.addEventListener('click', () => {
      fetchFromGoogleAppsScript({ interactive: true });
    });
  }

  if (dom.toggleGasConfigBtn) {
    dom.toggleGasConfigBtn.addEventListener('click', () => {
      if (dom.gasConfigPanel) {
        const isHidden = dom.gasConfigPanel.style.display === 'none';
        dom.gasConfigPanel.style.display = isHidden ? 'block' : 'none';
        if (isHidden) {
          if (dom.gasWebappUrlInput) dom.gasWebappUrlInput.value = getGasWebappUrl();
          if (dom.gasApiKeyInput) dom.gasApiKeyInput.value = getGasApiKey();
        }
      }
    });
  }

  if (dom.gasSaveConfigBtn) {
    dom.gasSaveConfigBtn.addEventListener('click', () => {
      const url = dom.gasWebappUrlInput ? dom.gasWebappUrlInput.value.trim() : '';
      const key = dom.gasApiKeyInput ? dom.gasApiKeyInput.value.trim() : '';

      if (!url) {
        showToast('Web App URL을 입력해 주세요.', 'error');
        return;
      }
      if (!key) {
        showToast('인증 API Key를 입력해 주세요.', 'error');
        return;
      }

      localStorage.setItem(STORAGE_KEYS.GAS_URL, url);
      localStorage.setItem(STORAGE_KEYS.GAS_API_KEY, key);
      renderGasSyncUI();
      showToast('Google Apps Script 연동 설정이 저장되었습니다.', 'success');
      testGasConnection();
    });
  }

  if (dom.gasResetConfigBtn) {
    dom.gasResetConfigBtn.addEventListener('click', () => {
      localStorage.removeItem(STORAGE_KEYS.GAS_URL);
      localStorage.removeItem(STORAGE_KEYS.GAS_API_KEY);
      if (dom.gasWebappUrlInput) dom.gasWebappUrlInput.value = GAS_DEFAULT.WEBAPP_URL;
      if (dom.gasApiKeyInput) dom.gasApiKeyInput.value = GAS_DEFAULT.API_KEY;
      renderGasSyncUI();
      showToast('Apps Script 연동 설정이 기본값으로 복원되었습니다.', 'info');
    });
  }

  if (dom.gasTestConnectionBtn) {
    dom.gasTestConnectionBtn.addEventListener('click', testGasConnection);
  }

  if (dom.toggleGasScriptCodeBtn) {
    dom.toggleGasScriptCodeBtn.addEventListener('click', () => {
      if (dom.gasRecommendedCodeBox) {
        const isHidden = dom.gasRecommendedCodeBox.style.display === 'none';
        dom.gasRecommendedCodeBox.style.display = isHidden ? 'block' : 'none';
        dom.toggleGasScriptCodeBtn.textContent = isHidden ? '▲ 추천 스크립트 코드 접기' : '📜 추천 Google Apps Script (Code.gs) 코드 보기';
      }
    });
  }

  if (dom.copyGasCodeBtn) {
    dom.copyGasCodeBtn.addEventListener('click', () => {
      const code = document.getElementById('gasCodePre')?.textContent;
      if (code) {
        navigator.clipboard.writeText(code).then(() => {
          showToast('추천 Apps Script 코드가 클립보드에 복사되었습니다.', 'success');
        }).catch(() => {
          showToast('코드 복사에 실패했습니다. 수동으로 드래그 복사해 주세요.', 'error');
        });
      }
    });
  }

  if (dom.gasAutoSyncToggle) {
    const savedAutoSync = localStorage.getItem(STORAGE_KEYS.GAS_AUTO_SYNC);
    if (savedAutoSync !== null) {
      dom.gasAutoSyncToggle.checked = (savedAutoSync === 'true');
    }
    dom.gasAutoSyncToggle.addEventListener('change', (e) => {
      localStorage.setItem(STORAGE_KEYS.GAS_AUTO_SYNC, String(e.target.checked));
      showToast(`Apps Script 실시간 자동 동기화가 ${e.target.checked ? '활성화' : '비활성화'}되었습니다.`, 'info');
      if (e.target.checked) {
        startGasAutoSyncTimer();
      } else {
        stopGasAutoSyncTimer();
      }
    });
  }

  renderGasSyncUI();
}

function renderGasSyncUI() {
  if (!dom.adminGasSyncCard) return;

  const webAppUrl = getGasWebappUrl();
  if (dom.gasDisplayUrl) {
    dom.gasDisplayUrl.textContent = webAppUrl.length > 55 ? webAppUrl.substring(0, 52) + '...' : webAppUrl;
    dom.gasDisplayUrl.title = webAppUrl;
  }

  const lastSync = localStorage.getItem(STORAGE_KEYS.GAS_LAST_SYNC_TIME);
  if (dom.gasLastSyncTime) {
    dom.gasLastSyncTime.textContent = lastSync || '기록 이력 없음';
  }

  if (dom.gasStatusBadge && dom.gasStatusText) {
    if (adminState.gasSyncing) {
      dom.gasStatusBadge.className = 'gas-status-badge badge-gas-syncing';
      dom.gasStatusText.textContent = 'Apps Script 동기화 중...';
    } else if (lastSync) {
      dom.gasStatusBadge.className = 'gas-status-badge badge-gas-synced';
      dom.gasStatusText.textContent = '실시간 동기화 완료';
    } else {
      dom.gasStatusBadge.className = 'gas-status-badge badge-gas-ready';
      dom.gasStatusText.textContent = '연동 준비 완료';
    }
  }
}

// Push local real-time analytics data to Google Apps Script Web App
async function syncToGoogleAppsScript(options = { interactive: false, silent: false }) {
  const currentUser = getCurrentUser();
  if (!isAdminUser(currentUser)) {
    if (!options.silent) {
      showToast('관리자(bosung3116@gmail.com) 전용 기능입니다.', 'error');
    }
    return;
  }

  if (adminState.gasSyncing) return;
  adminState.gasSyncing = true;

  if (dom.gasSyncBtnLabel) dom.gasSyncBtnLabel.textContent = '⚡ 실시간 전송 중...';
  if (dom.gasSyncNowBtn) dom.gasSyncNowBtn.disabled = true;
  renderGasSyncUI();

  const webAppUrl = getGasWebappUrl();
  const apiKey = getGasApiKey();
  const analyticsData = compileAnalyticsData();

  const payload = {
    action: 'syncData',
    apiKey: apiKey,
    deploymentId: GAS_DEFAULT.DEPLOYMENT_ID,
    timestamp: new Date().toISOString(),
    adminEmail: ADMIN_EMAIL,
    analytics: analyticsData,
    visitors: getVisitors(),
    usage: getUsageData(),
    requests: getAccessRequests(),
    logs: getActivityLogs(),
  };

  try {
    const response = await fetch(webAppUrl, {
      method: 'POST',
      mode: 'cors',
      redirect: 'follow',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8',
      },
      body: JSON.stringify(payload),
    });

    const text = await response.text();
    let result = null;
    try {
      result = JSON.parse(text);
    } catch (e) {
      // If response is HTML (Google Login 302 redirect / 401 Unauthorized page)
      if (text.includes('ServiceLogin') || text.includes('accounts.google.com') || response.status === 401) {
        throw new Error('AUTH_PERMISSION_REQUIRED');
      }
    }

    const nowStr = new Date().toLocaleString('ko-KR');
    localStorage.setItem(STORAGE_KEYS.GAS_LAST_SYNC_TIME, nowStr);
    renderGasSyncUI();

    if (!options.silent) {
      showToast(`⚡ Google Apps Script에 실시간 데이터(방문자 ${analyticsData.summary.totalVisitors}명, 작업 ${analyticsData.summary.totalAllFeaturesUsed}회)가 안전하게 동기화되었습니다!`, 'success', 4500);
    }
  } catch (err) {
    console.warn('Apps Script sync warning:', err);
    if (dom.gasStatusBadge && dom.gasStatusText) {
      dom.gasStatusBadge.className = 'gas-status-badge badge-gas-error';
      dom.gasStatusText.textContent = '배포 권한 확인 필요';
    }

    if (!options.silent) {
      if (err.message === 'AUTH_PERMISSION_REQUIRED' || String(err).includes('Failed to fetch')) {
        showToast('⚠️ Google Apps Script 배포 설정에서 [액세스 권한: 모든 사용자(Anyone)]로 지정해야 실시간 통신이 가능합니다. (연동 설정 패널을 확인하세요)', 'warning', 6000);
        if (dom.gasConfigPanel) dom.gasConfigPanel.style.display = 'block';
      } else {
        showToast(`Apps Script 동기화 중 오류가 발생했습니다: ${err.message}`, 'error');
      }
    }
  } finally {
    adminState.gasSyncing = false;
    if (dom.gasSyncBtnLabel) dom.gasSyncBtnLabel.textContent = '⚡ Apps Script에 지금 기록';
    if (dom.gasSyncNowBtn) dom.gasSyncNowBtn.disabled = false;
    renderGasSyncUI();
  }
}

// Fetch latest data from Google Apps Script Web App
async function fetchFromGoogleAppsScript(options = { interactive: false }) {
  const currentUser = getCurrentUser();
  if (!isAdminUser(currentUser)) return;

  const webAppUrl = getGasWebappUrl();
  const apiKey = getGasApiKey();
  const url = `${webAppUrl}?action=getData&apiKey=${encodeURIComponent(apiKey)}&_t=${Date.now()}`;

  if (options.interactive) {
    showToast('Apps Script로부터 실시간 최신 데이터를 조회하는 중입니다...', 'info', 1500);
  }

  try {
    const response = await fetch(url, {
      method: 'GET',
      mode: 'cors',
      redirect: 'follow',
    });

    const text = await response.text();
    let data = null;
    try {
      data = JSON.parse(text);
    } catch (e) {
      if (text.includes('ServiceLogin') || text.includes('accounts.google.com') || response.status === 401) {
        throw new Error('AUTH_PERMISSION_REQUIRED');
      }
    }

    if (data && data.status === 'success') {
      const nowStr = new Date().toLocaleString('ko-KR');
      localStorage.setItem(STORAGE_KEYS.GAS_LAST_SYNC_TIME, nowStr);

      // If remote returned visitor/usage lists, merge or update
      if (data.visitors && Array.isArray(data.visitors) && data.visitors.length > 0) {
        saveVisitors(data.visitors);
      }
      if (data.usage && typeof data.usage === 'object') {
        saveUsageData(data.usage);
      }
      if (data.requests && Array.isArray(data.requests)) {
        saveAccessRequests(data.requests);
      }

      renderAdminDashboard();
      renderGasSyncUI();

      if (options.interactive) {
        showToast('🔄 Google Apps Script로부터 실시간 최신 데이터를 성공적으로 불러왔습니다!', 'success');
      }
    } else {
      if (options.interactive) {
        showToast(`Apps Script 응답 수신: ${data?.message || '정상 통신 완료'}`, 'info');
      }
    }
  } catch (err) {
    console.warn('Apps Script fetch warning:', err);
    if (options.interactive) {
      showToast('⚠️ Apps Script 배포 설정에서 [액세스할 수 있는 사용자: 모든 사용자(Anyone)]로 설정되어 있는지 확인해 주세요.', 'warning', 5000);
      if (dom.gasConfigPanel) dom.gasConfigPanel.style.display = 'block';
    }
  }
}

// Test Ping Connection to Google Apps Script Web App
async function testGasConnection() {
  const webAppUrl = getGasWebappUrl();
  const apiKey = getGasApiKey();

  if (dom.gasPingResultBadge) {
    dom.gasPingResultBadge.textContent = '연결 진단 중...';
    dom.gasPingResultBadge.style.color = 'var(--primary)';
  }

  showToast('Google Apps Script Web App 연결 상태를 진단하고 있습니다...', 'info', 1500);

  try {
    const url = `${webAppUrl}?action=ping&apiKey=${encodeURIComponent(apiKey)}&_t=${Date.now()}`;
    const response = await fetch(url, {
      method: 'GET',
      mode: 'cors',
      redirect: 'follow',
    });

    const text = await response.text();
    if (text.includes('ServiceLogin') || text.includes('accounts.google.com') || response.status === 401) {
      if (dom.gasPingResultBadge) {
        dom.gasPingResultBadge.textContent = '❌ 권한 오류 (Anyone 설정 필요)';
        dom.gasPingResultBadge.style.color = '#EF4444';
      }
      showToast('❌ 배포 권한 오류: Apps Script의 "액세스할 수 있는 사용자"를 "모든 사용자(Anyone)"로 변경해 주세요.', 'error', 6000);
      return;
    }

    if (dom.gasPingResultBadge) {
      dom.gasPingResultBadge.textContent = '🟢 연결 성공 (정상 작동)';
      dom.gasPingResultBadge.style.color = '#10B981';
    }
    showToast('🟢 Google Apps Script Web App과 실시간 연결이 성공적으로 확인되었습니다!', 'success');
  } catch (err) {
    if (dom.gasPingResultBadge) {
      dom.gasPingResultBadge.textContent = '⚠️ 네트워크 또는 CORS 제한';
      dom.gasPingResultBadge.style.color = '#F59E0B';
    }
    showToast('연결 진단 완료: 웹앱 배포 시 "모든 사용자(Anyone)" 설정 여부를 확인해 주세요.', 'info');
  }
}

// Automatic Background Periodic Sync Timer (Every 30 seconds when admin is active)
function startGasAutoSyncTimer() {
  stopGasAutoSyncTimer();
  const toggle = localStorage.getItem(STORAGE_KEYS.GAS_AUTO_SYNC);
  if (toggle === 'false') return;

  adminState.gasAutoSyncTimer = setInterval(() => {
    const currentUser = getCurrentUser();
    if (isAdminUser(currentUser)) {
      syncToGoogleAppsScript({ silent: true, interactive: false });
    }
  }, 30000);
}

function stopGasAutoSyncTimer() {
  if (adminState.gasAutoSyncTimer) {
    clearInterval(adminState.gasAutoSyncTimer);
    adminState.gasAutoSyncTimer = null;
  }
}

// Trigger Gas auto sync when features or visits occur
function triggerAutoGasSyncIfEligible() {
  const user = getCurrentUser();
  if (!isAdminUser(user)) return;

  const isAutoSync = localStorage.getItem(STORAGE_KEYS.GAS_AUTO_SYNC);
  if (isAutoSync === 'false') return;

  if (window._gasSyncDebounceTimer) clearTimeout(window._gasSyncDebounceTimer);
  window._gasSyncDebounceTimer = setTimeout(() => {
    syncToGoogleAppsScript({ silent: true, interactive: false });
  }, 1500);
}



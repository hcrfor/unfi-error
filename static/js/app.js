/**
 * 산림 도시 조사 데이터 오류 검수 시스템 클라이언트 스크립트
 * 파일명: static/js/app.js
 * 설명: 웹 드래그 앤 드롭 업로드, 샘플 검수, 로컬 폴더 검수 및 실시간 결과 대시보드 렌더링
 */

// 전역 상태
let currentInspectionData = null;
let currentMode = 'upload'; // 'upload' 또는 'local'
let selectedFiles = [];     // 사용자가 선택한 File 객체 배열

document.addEventListener('DOMContentLoaded', () => {
  initEventListeners();
  initDropzone();
});

function initEventListeners() {
  const btnBrowse = document.getElementById('btnBrowseFolder');
  const btnStart = document.getElementById('btnStartInspection');
  const btnDownload = document.getElementById('btnDownloadReport');
  const btnOpenFolder = document.getElementById('btnOpenFolder');
  const sampleSearchInput = document.getElementById('sampleSearchInput');
  const sheetFilterSelect = document.getElementById('sheetFilterSelect');
  const errorSearchInput = document.getElementById('errorSearchInput');

  // 파일 입력 필드 변경 이벤트
  const fileInput = document.getElementById('fileUploadInput');
  const folderInput = document.getElementById('folderUploadInput');

  if (fileInput) {
    fileInput.addEventListener('change', (e) => handleFilesSelected(e.target.files));
  }
  if (folderInput) {
    folderInput.addEventListener('change', (e) => handleFilesSelected(e.target.files));
  }

  // 1. 폴더 찾아보기 버튼 (로컬 모드)
  if (btnBrowse) {
    btnBrowse.addEventListener('click', handleBrowseFolder);
  }

  // 2. 검수 시작 버튼
  if (btnStart) {
    btnStart.addEventListener('click', handleStartInspection);
  }

  // 3. 엑셀 리포트 다운로드
  if (btnDownload) {
    btnDownload.addEventListener('click', () => {
      if (currentInspectionData && currentInspectionData.report_filename) {
        const folder = currentInspectionData.output_dir || '';
        const filename = encodeURIComponent(currentInspectionData.report_filename);
        window.location.href = `/api/download-report?folder=${encodeURIComponent(folder)}&filename=${filename}`;
      }
    });
  }

  // 4. 결과 폴더 열기
  if (btnOpenFolder) {
    btnOpenFolder.addEventListener('click', handleOpenFolder);
  }

  // 5. 표본점 요약 테이블 검색
  if (sampleSearchInput) {
    sampleSearchInput.addEventListener('input', filterSampleTable);
  }

  // 6. 상세 오류 테이블 필터 (시트명 & 검색어)
  if (sheetFilterSelect) {
    sheetFilterSelect.addEventListener('change', filterErrorTable);
  }
  if (errorSearchInput) {
    errorSearchInput.addEventListener('input', filterErrorTable);
  }
}

/**
 * 모드 전환 함수 (웹 업로드 vs 로컬 경로)
 */
function switchMode(mode) {
  currentMode = mode;
  const tabUpload = document.getElementById('tabUploadMode');
  const tabLocal = document.getElementById('tabLocalMode');
  const secUpload = document.getElementById('uploadModeSection');
  const secLocal = document.getElementById('localModeSection');

  if (mode === 'upload') {
    tabUpload.classList.add('active');
    tabLocal.classList.remove('active');
    secUpload.style.display = 'block';
    secLocal.style.display = 'none';
  } else {
    tabLocal.classList.add('active');
    tabUpload.classList.remove('active');
    secLocal.style.display = 'block';
    secUpload.style.display = 'none';
  }
}

/**
 * 드래그 앤 드롭 존 초기화
 */
function initDropzone() {
  const dropZone = document.getElementById('dropZone');
  if (!dropZone) return;

  ['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add('dragover');
    }, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove('dragover');
    }, false);
  });

  dropZone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    if (dt && dt.files && dt.files.length > 0) {
      handleFilesSelected(dt.files);
    }
  });
}

/**
 * 파일 선택 처리 함수
 */
function handleFilesSelected(fileList) {
  const validFiles = [];
  for (let i = 0; i < fileList.length; i++) {
    const f = fileList[i];
    const name = f.name;
    if (name.toLowerCase().endsWith('.xlsx') && !name.startsWith('~$') && !name.includes('오류_검수_결과')) {
      validFiles.push(f);
    }
  }

  if (validFiles.length === 0) {
    alert('선택된 항목 중 유효한 .xlsx 조사 엑셀 파일이 없습니다.');
    return;
  }

  selectedFiles = validFiles;

  // UI 업데이트
  const bar = document.getElementById('selectedFilesBar');
  const countBadge = document.getElementById('selectedFilesCount');
  const preview = document.getElementById('selectedFilesNames');

  bar.style.display = 'flex';
  countBadge.innerText = `${selectedFiles.length}개 파일 선택됨`;

  const names = selectedFiles.map(f => f.name);
  if (names.length <= 3) {
    preview.innerText = names.join(', ');
  } else {
    preview.innerText = `${names.slice(0, 3).join(', ')} 외 ${names.length - 3}개`;
  }
}

/**
 * 선택 파일 초기화
 */
function clearSelectedFiles() {
  selectedFiles = [];
  document.getElementById('selectedFilesBar').style.display = 'none';
  document.getElementById('fileUploadInput').value = '';
  document.getElementById('folderUploadInput').value = '';
}

/**
 * 빠른 폴더 선택 칩 (로컬 모드)
 */
function setQuickFolder(folderPath) {
  document.getElementById('targetFolderInput').value = folderPath;
}

/**
 * 내장 샘플 데이터(excel_1-4)로 즉시 체험
 */
async function handleRunSample() {
  const btnStart = document.getElementById('btnStartInspection');
  const progressWrap = document.getElementById('progressWrap');
  const progressMsg = document.getElementById('progressMsg');
  const progressPercent = document.getElementById('progressPercent');
  const progressBarFill = document.getElementById('progressBarFill');
  const statusPill = document.querySelector('.status-indicator');
  const statusText = document.getElementById('systemStatusText');

  btnStart.disabled = true;
  btnStart.classList.add('loading');
  progressWrap.style.display = 'block';
  progressMsg.innerText = '내장 샘플(excel_1-4) 100여 개 파일을 분석하고 있습니다...';
  progressPercent.innerText = '분석 중';
  progressBarFill.style.width = '65%';

  statusPill.classList.add('busy');
  statusText.innerText = '샘플 데이터 검수 중...';

  try {
    const res = await fetch('/api/sample-inspection', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    const result = await res.json();
    if (!res.ok || !result.success) {
      throw new Error(result.error || '샘플 검수 중 오류가 발생했습니다.');
    }

    onInspectionSuccess(result.data);
  } catch (err) {
    onInspectionError(err);
  } finally {
    btnStart.disabled = false;
    btnStart.classList.remove('loading');
  }
}

/**
 * 윈도우 기본 폴더 선택 창 호출 (/api/select-folder)
 */
async function handleBrowseFolder() {
  const btnBrowse = document.getElementById('btnBrowseFolder');
  btnBrowse.disabled = true;
  btnBrowse.innerText = '선택 중...';

  try {
    const res = await fetch('/api/select-folder', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    const data = await res.json();
    if (data.success && data.folder_path) {
      document.getElementById('targetFolderInput').value = data.folder_path;
    } else if (data.message) {
      alert(data.message);
    }
  } catch (err) {
    console.error('폴더 선택 창 호출 오류:', err);
    alert('폴더 선택 창을 여는 중 문제가 발생했습니다.');
  } finally {
    btnBrowse.disabled = false;
    btnBrowse.innerText = '폴더 찾아보기';
  }
}

/**
 * 정밀 검수 시작 버튼 클릭 핸들러
 */
async function handleStartInspection() {
  if (currentMode === 'upload') {
    if (selectedFiles.length === 0) {
      alert('검수할 엑셀 파일을 드래그하여 놓거나 [파일 직접 선택] 또는 [폴더째 선택] 버튼을 눌러주세요.');
      return;
    }
    await executeUploadInspection();
  } else {
    await executeLocalFolderInspection();
  }
}

/**
 * 웹 파일 업로드 검수 실행
 */
async function executeUploadInspection() {
  const btnStart = document.getElementById('btnStartInspection');
  const progressWrap = document.getElementById('progressWrap');
  const progressMsg = document.getElementById('progressMsg');
  const progressPercent = document.getElementById('progressPercent');
  const progressBarFill = document.getElementById('progressBarFill');
  const statusPill = document.querySelector('.status-indicator');
  const statusText = document.getElementById('systemStatusText');

  btnStart.disabled = true;
  btnStart.classList.add('loading');
  progressWrap.style.display = 'block';
  progressMsg.innerText = `선택된 ${selectedFiles.length}개 엑셀 파일을 업로드하고 정밀 검수를 진행 중입니다...`;
  progressPercent.innerText = '분석 중';
  progressBarFill.style.width = '70%';

  statusPill.classList.add('busy');
  statusText.innerText = '파일 업로드 및 검수 중...';

  const formData = new FormData();
  for (let i = 0; i < selectedFiles.length; i++) {
    formData.append('files', selectedFiles[i]);
  }

  try {
    const res = await fetch('/api/upload-and-inspect', {
      method: 'POST',
      body: formData
    });
    const result = await res.json();
    if (!res.ok || !result.success) {
      throw new Error(result.error || '검수 중 오류가 발생했습니다.');
    }

    onInspectionSuccess(result.data);
  } catch (err) {
    onInspectionError(err);
  } finally {
    btnStart.disabled = false;
    btnStart.classList.remove('loading');
  }
}

/**
 * 로컬 경로 검수 실행
 */
async function executeLocalFolderInspection() {
  const folderInput = document.getElementById('targetFolderInput');
  const targetFolder = folderInput.value.trim();

  if (!targetFolder) {
    alert('검수할 폴더 경로를 입력해주세요.');
    folderInput.focus();
    return;
  }

  const btnStart = document.getElementById('btnStartInspection');
  const progressWrap = document.getElementById('progressWrap');
  const progressMsg = document.getElementById('progressMsg');
  const progressPercent = document.getElementById('progressPercent');
  const progressBarFill = document.getElementById('progressBarFill');
  const statusPill = document.querySelector('.status-indicator');
  const statusText = document.getElementById('systemStatusText');

  btnStart.disabled = true;
  btnStart.classList.add('loading');
  progressWrap.style.display = 'block';
  progressMsg.innerText = '로컬 폴더의 엑셀 파일들을 분석하고 있습니다...';
  progressPercent.innerText = '분석 중';
  progressBarFill.style.width = '70%';

  statusPill.classList.add('busy');
  statusText.innerText = '검수 엔진 실행 중...';

  try {
    const res = await fetch('/api/start-inspection', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ target_folder: targetFolder })
    });
    const result = await res.json();
    if (!res.ok || !result.success) {
      throw new Error(result.error || '검수 중 서버 오류가 발생했습니다.');
    }

    onInspectionSuccess(result.data);
  } catch (err) {
    onInspectionError(err);
  } finally {
    btnStart.disabled = false;
    btnStart.classList.remove('loading');
  }
}

function onInspectionSuccess(data) {
  const progressBarFill = document.getElementById('progressBarFill');
  const progressPercent = document.getElementById('progressPercent');
  const progressMsg = document.getElementById('progressMsg');
  const statusPill = document.querySelector('.status-indicator');
  const statusText = document.getElementById('systemStatusText');

  progressBarFill.style.width = '100%';
  progressPercent.innerText = '100%';
  progressMsg.innerText = '검수가 성공적으로 완료되었습니다!';
  statusPill.classList.remove('busy');
  statusText.innerText = '검수 완료';

  currentInspectionData = data;
  renderDashboard(data);
}

function onInspectionError(err) {
  console.error('검수 실패:', err);
  alert(`검수 중 오류가 발생했습니다:\n${err.message}`);
  document.getElementById('progressWrap').style.display = 'none';
  const statusPill = document.querySelector('.status-indicator');
  statusPill.classList.remove('busy');
  document.getElementById('systemStatusText').innerText = '오류 발생';
}

/**
 * 검수 결과 대시보드 렌더링
 */
function renderDashboard(data) {
  const resultsSection = document.getElementById('resultsSection');
  resultsSection.style.display = 'block';

  // 1. 주요 지표 카드
  document.getElementById('metricTotalFiles').innerText = data.total_files.toLocaleString();
  document.getElementById('metricPassFiles').innerText = data.pass_files.toLocaleString();
  document.getElementById('metricErrorFiles').innerText = data.error_files.toLocaleString();
  document.getElementById('metricTotalErrors').innerText = data.total_errors.toLocaleString();

  // 2. 리포트 파일명 표시
  document.getElementById('reportFileName').innerText = data.report_filename;

  // 3. 규칙별 통계 렌더링
  renderRuleStats(data.rule_stats, data.total_errors);

  // 4. 표본점별 요약 테이블 렌더링
  renderSampleTable(data.file_summary);

  // 5. 상세 오류 테이블 렌더링
  renderErrorTable(data.errors);

  // 결과 영역으로 부드럽게 스크롤 이동
  resultsSection.scrollIntoView({ behavior: 'smooth' });
}

/**
 * 규칙별 통계 바 리스트 렌더링
 */
function renderRuleStats(ruleStats, totalErrors) {
  const container = document.getElementById('ruleStatsList');
  container.innerHTML = '';

  const ruleEntries = Object.entries(ruleStats || {}).sort((a, b) => b[1] - a[1]);
  document.getElementById('ruleTotalCountBadge').innerText = `${ruleEntries.length}개 규칙 집계`;

  if (ruleEntries.length === 0) {
    container.innerHTML = '<div style="padding: 20px; text-align: center; color: #10b981; font-weight: 500;">🎉 발견된 오류가 없습니다. 모든 규칙을 완벽하게 통과했습니다!</div>';
    return;
  }

  const maxCount = ruleEntries[0][1] || 1;

  ruleEntries.forEach(([ruleName, count]) => {
    const ratio = totalErrors > 0 ? ((count / totalErrors) * 100).toFixed(1) : 0;
    const barWidth = ((count / maxCount) * 100).toFixed(1);

    const item = document.createElement('div');
    item.className = 'rule-stat-item';
    item.innerHTML = `
      <div class="rule-stat-top">
        <span class="rule-stat-name">${escapeHtml(ruleName)}</span>
        <span class="rule-stat-count">${count.toLocaleString()}건 (${ratio}%)</span>
      </div>
      <div class="rule-stat-bar-bg">
        <div class="rule-stat-bar-fill" style="width: ${barWidth}%;"></div>
      </div>
    `;
    container.appendChild(item);
  });
}

/**
 * 표본점별 요약 테이블 렌더링
 */
function renderSampleTable(fileSummary) {
  const tbody = document.getElementById('sampleSummaryBody');
  tbody.innerHTML = '';

  (fileSummary || []).forEach((item, idx) => {
    const tr = document.createElement('tr');
    const isError = item.오류건수 > 0;
    const statusBadge = isError 
      ? '<span class="badge badge-danger">오류 발견</span>' 
      : '<span class="badge badge-success">정상 통과</span>';

    tr.innerHTML = `
      <td class="cell-center">${idx + 1}</td>
      <td class="cell-code"><strong>${escapeHtml(item.표본점번호)}</strong></td>
      <td>${escapeHtml(item.파일명)}</td>
      <td class="cell-center">${escapeHtml(item.팀장 || '-')}</td>
      <td class="cell-center" style="font-size: 13px;">${escapeHtml(item.팀원 || '-')}</td>
      <td class="cell-right" style="font-weight: 600; color: ${isError ? '#dc2626' : '#059669'};">${item.오류건수}</td>
      <td class="cell-center">${statusBadge}</td>
    `;
    tbody.appendChild(tr);
  });
}

/**
 * 표본점 테이블 실시간 검색 필터
 */
function filterSampleTable() {
  const keyword = document.getElementById('sampleSearchInput').value.trim().toLowerCase();
  const rows = document.querySelectorAll('#sampleSummaryBody tr');

  rows.forEach(tr => {
    const text = tr.innerText.toLowerCase();
    tr.style.display = text.includes(keyword) ? '' : 'none';
  });
}

/**
 * 상세 오류 테이블 렌더링
 */
function renderErrorTable(errors) {
  const tbody = document.getElementById('errorDetailBody');
  tbody.innerHTML = '';
  document.getElementById('detailErrorCountBadge').innerText = `${(errors || []).length.toLocaleString()}건`;

  if (!errors || errors.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="11" style="text-align: center; padding: 36px; color: #10b981; font-weight: 600;">
          ✨ 완벽합니다! 발견된 상세 오류가 없습니다.
        </td>
      </tr>
    `;
    return;
  }

  errors.forEach((err, idx) => {
    const tr = document.createElement('tr');
    tr.dataset.sheet = err.시트명 || '';

    tr.innerHTML = `
      <td class="cell-center">${idx + 1}</td>
      <td class="cell-code">${escapeHtml(err.표본점번호)}</td>
      <td style="font-size: 13px;">${escapeHtml(err.파일명)}</td>
      <td class="cell-center">${escapeHtml(err.팀장 || '-')}</td>
      <td class="cell-center" style="font-size: 13px;">${escapeHtml(err.팀원 || '-')}</td>
      <td class="cell-center"><span class="badge badge-sheet">${escapeHtml(err.시트명)}</span></td>
      <td class="cell-center" style="font-family: monospace;">${escapeHtml(String(err.행번호))}</td>
      <td style="font-size: 13px; font-weight: 500;">${escapeHtml(err.검증규칙)}</td>
      <td><span class="cell-field-tag">${escapeHtml(err.오류항목)}</span></td>
      <td style="font-size: 13px; color: #1e293b;">${escapeHtml(err.오류내용)}</td>
      <td class="cell-code" style="color: #b91c1c;">${escapeHtml(String(err.입력값_상세))}</td>
    `;
    tbody.appendChild(tr);
  });
}

/**
 * 상세 오류 테이블 복합 필터 (시트명 + 검색어)
 */
function filterErrorTable() {
  const selectedSheet = document.getElementById('sheetFilterSelect').value;
  const keyword = document.getElementById('errorSearchInput').value.trim().toLowerCase();
  const rows = document.querySelectorAll('#errorDetailBody tr');

  let visibleCount = 0;
  rows.forEach(tr => {
    const rowSheet = tr.dataset.sheet || '';
    const rowText = tr.innerText.toLowerCase();

    const sheetMatch = (selectedSheet === 'ALL' || rowSheet === selectedSheet);
    const textMatch = (!keyword || rowText.includes(keyword));

    if (sheetMatch && textMatch) {
      tr.style.display = '';
      visibleCount++;
    } else {
      tr.style.display = 'none';
    }
  });

  document.getElementById('detailErrorCountBadge').innerText = `${visibleCount.toLocaleString()}건`;
}

/**
 * 결과 폴더 열기 (/api/open-folder)
 */
async function handleOpenFolder() {
  if (!currentInspectionData) return;
  const folder = currentInspectionData.output_dir || '';

  try {
    const res = await fetch('/api/open-folder', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ folder_path: folder })
    });
    const data = await res.json();
    if (!data.success && data.error) {
      alert(data.error);
    }
  } catch (err) {
    alert('폴더를 여는 중 문제가 발생했습니다.');
  }
}

/**
 * XSS 방지 HTML 이스케이프 유틸리티
 */
function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

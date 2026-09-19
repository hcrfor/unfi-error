/**
 * 산림 도시 조사 데이터 오류 검수 시스템 클라이언트 스크립트
 * 파일명: static/js/app.js
 * 설명: 검수 요청, 폴더 선택 대화상자 호출, 실시간 진행 표시 및 대시보드 렌더링
 */

// 전역 상태
let currentInspectionData = null;

document.addEventListener('DOMContentLoaded', () => {
  initEventListeners();
});

function initEventListeners() {
  const btnBrowse = document.getElementById('btnBrowseFolder');
  const btnStart = document.getElementById('btnStartInspection');
  const btnDownload = document.getElementById('btnDownloadReport');
  const btnOpenFolder = document.getElementById('btnOpenFolder');
  const sampleSearchInput = document.getElementById('sampleSearchInput');
  const sheetFilterSelect = document.getElementById('sheetFilterSelect');
  const errorSearchInput = document.getElementById('errorSearchInput');

  // 1. 폴더 찾아보기 버튼
  btnBrowse.addEventListener('click', handleBrowseFolder);

  // 2. 검수 시작 버튼
  btnStart.addEventListener('click', handleStartInspection);

  // 3. 엑셀 리포트 다운로드
  btnDownload.addEventListener('click', () => {
    if (currentInspectionData && currentInspectionData.report_filename) {
      const folder = currentInspectionData.output_dir || '';
      const filename = encodeURIComponent(currentInspectionData.report_filename);
      window.location.href = `/api/download-report?folder=${encodeURIComponent(folder)}&filename=${filename}`;
    }
  });

  // 4. 결과 폴더 열기 (윈도우 탐색기 - 다운로드 폴더)
  btnOpenFolder.addEventListener('click', handleOpenFolder);

  // 5. 표본점 요약 테이블 검색
  sampleSearchInput.addEventListener('input', filterSampleTable);

  // 6. 상세 오류 테이블 필터 (시트명 & 검색어)
  sheetFilterSelect.addEventListener('change', filterErrorTable);
  errorSearchInput.addEventListener('input', filterErrorTable);
}

// 빠른 폴더 선택 칩
function setQuickFolder(folderPath) {
  document.getElementById('targetFolderInput').value = folderPath;
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
 * 정밀 검수 시작 (/api/start-inspection)
 */
async function handleStartInspection() {
  const folderInput = document.getElementById('targetFolderInput');
  const targetFolder = folderInput.value.trim();

  if (!targetFolder) {
    alert('검수할 폴더 경로를 입력하거나 선택해주세요.');
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

  // UI 상태 업데이트: 검수 진행 중
  btnStart.disabled = true;
  btnStart.classList.add('loading');
  progressWrap.style.display = 'block';
  progressMsg.innerText = '조사 엑셀 파일들을 불러오고 정밀 분석을 진행 중입니다...';
  progressPercent.innerText = '분석 중';
  progressBarFill.style.width = '75%';

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

    // 검수 완료 처리
    progressBarFill.style.width = '100%';
    progressPercent.innerText = '100%';
    progressMsg.innerText = '검수가 성공적으로 완료되었습니다!';
    statusPill.classList.remove('busy');
    statusText.innerText = '검수 완료';

    currentInspectionData = result.data;

    // 대시보드 렌더링
    renderDashboard(result.data);

  } catch (err) {
    console.error('검수 실패:', err);
    alert(`검수 중 오류가 발생했습니다:\n${err.message}`);
    progressWrap.style.display = 'none';
    statusPill.classList.remove('busy');
    statusText.innerText = '오류 발생';
  } finally {
    btnStart.disabled = false;
    btnStart.classList.remove('loading');
  }
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
  document.getElementById('detailErrorCountBadge').innerText = `${(errors || []).length}건`;

  if (!errors || errors.length === 0) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align: center; padding: 30px; color: #10b981; font-weight: 600;">발견된 오류 항목이 없습니다.</td></tr>';
    return;
  }

  errors.forEach((err, idx) => {
    const tr = document.createElement('tr');
    tr.dataset.sheet = err.시트명;

    tr.innerHTML = `
      <td class="cell-center">${idx + 1}</td>
      <td class="cell-code"><strong>${escapeHtml(err.표본점번호)}</strong></td>
      <td style="font-size: 12px;">${escapeHtml(err.파일명)}</td>
      <td><span class="badge">${escapeHtml(err.시트명)}</span></td>
      <td class="cell-center">${err.행번호}</td>
      <td style="font-weight: 600; color: #1e293b;">${escapeHtml(err.검증규칙)}</td>
      <td style="color: #dc2626; font-weight: 600;">${escapeHtml(err.오류항목)}</td>
      <td>${escapeHtml(err.오류내용)}</td>
      <td class="cell-code" style="max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${escapeHtml(err.입력값_상세)}">${escapeHtml(err.입력값_상세)}</td>
    `;
    tbody.appendChild(tr);
  });
}

/**
 * 상세 오류 테이블 복합 필터 (시트 셀렉트 + 검색창)
 */
function filterErrorTable() {
  const selectedSheet = document.getElementById('sheetFilterSelect').value;
  const keyword = document.getElementById('errorSearchInput').value.trim().toLowerCase();
  const rows = document.querySelectorAll('#errorDetailBody tr');

  let visibleCount = 0;
  rows.forEach(tr => {
    const rowSheet = tr.dataset.sheet;
    const rowText = tr.innerText.toLowerCase();

    const matchesSheet = (selectedSheet === 'ALL' || rowSheet === selectedSheet);
    const matchesKeyword = (!keyword || rowText.includes(keyword));

    if (matchesSheet && matchesKeyword) {
      tr.style.display = '';
      visibleCount++;
    } else {
      tr.style.display = 'none';
    }
  });

  document.getElementById('detailErrorCountBadge').innerText = `${visibleCount}건`;
}

/**
 * 탐색기로 결과 폴더 열기 (/api/open-folder)
 */
async function handleOpenFolder() {
  const folder = currentInspectionData?.output_dir || '';
  try {
    const res = await fetch('/api/open-folder', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ folder_path: folder })
    });
    const data = await res.json();
    if (!data.success) {
      alert(data.error || '폴더를 여는 중 오류가 발생했습니다.');
    }
  } catch (err) {
    console.error('폴더 열기 요청 오류:', err);
  }
}

function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

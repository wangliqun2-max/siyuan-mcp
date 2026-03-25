/**
 * script.js - Siyuan Tender Parameter Extraction Tool
 * 4-step workflow: Upload → Locate Section → Extract Params → Download Excel
 */

const API_BASE = 'http://localhost:5001/api';

// ==================== STATE ====================
let state = {
    files: [],           // Array<File>
    productType: 'transformer',
    sessionId: null,
    section: null,
    extracted: null,
    totalPages: 0,
    filePageRanges: []   // [{filename, start_page, end_page, pages}]
};

// ==================== INIT ====================
document.addEventListener('DOMContentLoaded', () => {
    setupUploadZone();
    setupFileInput();
    setupParamTooltip();
});

// ==================== UPLOAD ZONE ====================
function setupUploadZone() {
    const zone = document.getElementById('uploadZone');
    zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('dragover');
        addFiles(e.dataTransfer.files);
    });
    zone.addEventListener('click', (e) => {
        if (e.target.closest('.file-list-item') || e.target.classList.contains('btn-link')) return;
        document.getElementById('fileInput').click();
    });
}

function setupFileInput() {
    document.getElementById('fileInput').addEventListener('change', (e) => {
        addFiles(e.target.files);
        e.target.value = ''; // allow re-selecting same file
    });
}

function addFiles(fileList) {
    const allowed = ['.pdf', '.docx', '.doc'];
    const existing = new Set(state.files.map(f => f.name));
    let added = 0;
    for (const file of fileList) {
        const ext = '.' + file.name.split('.').pop().toLowerCase();
        if (!allowed.includes(ext)) { showError(`不支持的格式：${file.name}`); continue; }
        if (existing.has(file.name)) { showError(`文件已存在：${file.name}`); continue; }
        state.files.push(file);
        existing.add(file.name);
        added++;
    }
    if (added) renderFileList();
}

function removeFile(filename) {
    state.files = state.files.filter(f => f.name !== filename);
    renderFileList();
}

function renderFileList() {
    const listEl = document.getElementById('fileList');
    const btnLocate = document.getElementById('btnLocate');
    if (!state.files.length) {
        listEl.innerHTML = '';
        btnLocate.disabled = true;
        return;
    }
    btnLocate.disabled = false;
    listEl.innerHTML = state.files.map((f, i) => `
        <div class="file-list-item">
            <span class="file-list-icon">${f.name.endsWith('.pdf') ? '📄' : '📝'}</span>
            <span class="file-list-name" title="${f.name}">${f.name}</span>
            <span class="file-list-size">${(f.size / 1024 / 1024).toFixed(1)} MB</span>
            <button class="btn-remove" onclick="removeFile('${f.name.replace(/'/g, "\\'")}')">✕</button>
        </div>
    `).join('');
}

// ==================== STEP 1: LOCATE SECTION ====================
async function locateSection() {
    if (!state.files.length) return;

    showLoading(`正在解析 ${state.files.length} 个文件并定位变压器章节...`, 2);
    setLoadingStep(1, 'active');

    const formData = new FormData();
    state.files.forEach(f => formData.append('files', f));
    formData.append('product_type', state.productType);

    try {
        setTimeout(() => setLoadingStep(1, 'done'), 500);
        setTimeout(() => setLoadingStep(2, 'active'), 800);

        const response = await fetch(`${API_BASE}/locate-section`, {
            method: 'POST',
            body: formData
        });

        const data = await response.json();
        if (!data.success) throw new Error(data.error || '章节定位失败');

        setLoadingStep(2, 'done');
        state.sessionId = data.session_id;
        state.section = data.section;
        state.totalPages = data.total_pages;
        state.filePageRanges = data.file_page_ranges || [];

        setTimeout(() => showSectionResult(data), 400);
    } catch (err) {
        hideLoading();
        showError(err.message);
    }
}

function showSectionResult(data) {
    hideLoading();
    const section = data.section;

    // File summary line
    const ranges = data.file_page_ranges || [];
    const fileSummary = ranges.length > 1
        ? `${ranges.length} 个文件 · ` + ranges.map(r => `${r.filename}(${r.pages}页)`).join(' + ')
        : (data.filename || '');
    document.getElementById('totalPages').textContent = `共 ${data.total_pages} 页${ranges.length > 1 ? ' · ' + ranges.length + '个文件' : ''}`;

    document.getElementById('sectionTitle').textContent = section.section_title || '（未知章节标题）';
    document.getElementById('sectionPages').textContent = `第 ${section.start_page} — ${section.end_page} 页`;
    document.getElementById('sectionNotes').textContent = section.notes || '';
    document.getElementById('startPage').value = section.start_page || 1;
    document.getElementById('endPage').value = section.end_page || data.total_pages;
    document.getElementById('endPage').max = data.total_pages;
    document.getElementById('startPage').max = data.total_pages;

    const conf = section.confidence || 0;
    const badge = document.getElementById('confidenceBadge');
    if (conf >= 0.8) { badge.className = 'confidence-badge high'; badge.textContent = `✓ 置信度 ${Math.round(conf*100)}%`; }
    else if (conf >= 0.5) { badge.className = 'confidence-badge medium'; badge.textContent = `⚠ 置信度 ${Math.round(conf*100)}%`; }
    else { badge.className = 'confidence-badge low'; badge.textContent = `⚠ 低置信度 ${Math.round(conf*100)}%`; }

    document.getElementById('step1').classList.add('hidden');
    document.getElementById('step2').classList.remove('hidden');
    setStepIndicator(2);
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ==================== STEP 2: EXTRACT PARAMS ====================
async function extractParams() {
    const startPage = parseInt(document.getElementById('startPage').value);
    const endPage = parseInt(document.getElementById('endPage').value);

    if (!startPage || !endPage || startPage > endPage) {
        showError('请检查页码范围，起始页不能大于结束页');
        return;
    }

    // Show loading (step 3-4 stages)
    showLoading('AI 正在精准提取技术规格参数...', 4);
    setStepIndicator(3);   // advance sidebar to step 3, light up connector 2→3
    setLoadingStep(1, 'done');
    setLoadingStep(2, 'done');
    setLoadingStep(3, 'active');

    try {
        // 自动学习：将确认的章节标题保存到后端知识库（静默，失败不阻断流程）
        if (state.section && state.section.section_title) {
            try {
                await fetch(`${API_BASE}/confirm-section`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        section_title: state.section.section_title,
                        product_type: 'power_transformer'
                    })
                });
                console.log('[TitleMap] Section title confirmed:', state.section.section_title);
            } catch (e) {
                console.warn('[TitleMap] Failed to save section title (non-critical):', e);
            }
        }

        const response = await fetch(`${API_BASE}/extract-params`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                session_id: state.sessionId,
                start_page: startPage,
                end_page: endPage,
                product_type: state.productType
            })
        });

        const data = await response.json();

        if (!data.success) {
            throw new Error(data.error || '参数提取失败');
        }

        setLoadingStep(3, 'done');
        setLoadingStep(4, 'active');
        state.extracted = data.extracted;

        setTimeout(() => {
            setLoadingStep(4, 'done');
            setTimeout(() => showResults(data), 400);
        }, 600);

    } catch (err) {
        hideLoading();
        showError(err.message);
    }
}


// Global doc viewer state
let _allChunks = {};
let _extractedData = {};

function showResults(data) {
    hideLoading();

    const { extracted, sections, all_chunks, stats } = data;
    _allChunks = all_chunks || {};
    _extractedData = extracted || {};

    // Stats pills
    const pills = document.getElementById('statsPills');
    pills.innerHTML = `
        <span class="stat-pill" style="background:var(--success-bg);color:var(--success)">✓ 找到 ${stats.found}</span>
        <span class="stat-pill" style="background:var(--danger-bg);color:var(--danger)">✗ 未找到 ${stats.not_found}</span>
        <span class="stat-pill" style="background:rgba(148,163,184,0.1);color:var(--text-secondary)">共 ${stats.total} 项</span>
    `;
    document.getElementById('resultSubtitle').textContent =
        `提取到 ${stats.found}/${stats.total} 个参数`;

    // ── Left panel: hierarchical param groups ───────────────────────────
    const groupsEl = document.getElementById('paramGroups');
    groupsEl.innerHTML = '';

    // If no sections data, create a single flat group
    const sectionsData = (sections && sections.length) ? sections : [{ title: '参数', params: Object.keys(extracted) }];

    for (const section of sectionsData) {
        const paramNames = section.params || [];
        if (!paramNames.length) continue;

        const foundInSection = paramNames.filter(p => extracted[p]?.found).length;
        const sectionEl = document.createElement('div');
        sectionEl.className = 'param-section';

        const toggleEl = document.createElement('div');
        toggleEl.className = 'section-toggle';
        toggleEl.innerHTML = `
            <span class="section-toggle-icon">▾</span>
            <span class="section-name">${escapeHtml(section.title)}</span>
            <span class="section-badge ${foundInSection > 0 ? 'has-found' : ''}">
                ${foundInSection}/${paramNames.length}
            </span>
        `;
        toggleEl.addEventListener('click', () => {
            const collapsed = toggleEl.classList.toggle('collapsed');
            paramsEl.classList.toggle('collapsed', collapsed);
        });

        const paramsEl = document.createElement('div');
        paramsEl.className = 'section-params';

        for (const paramName of paramNames) {
            const paramData = extracted[paramName] || {};
            const found = paramData.found === true;
            const value = paramData.value;
            const unit = paramData.unit || '';

            // Display only Chinese portion if name starts with English
            const displayName = paramName.replace(/^[^\u4e00-\u9fff\uff00-\uffef]+/, '').trim() || paramName;

            const rowEl = document.createElement('div');
            rowEl.className = 'param-row';
            rowEl.dataset.param = paramName;
            rowEl.innerHTML = `
                <span class="param-status-dot ${found ? 'found' : 'not-found'}"></span>
                <span class="param-row-name">${escapeHtml(displayName)}</span>
                ${found && value ? `<span class="param-row-value">${escapeHtml(String(value))}</span>` : ''}
                ${found && unit ? `<span class="param-row-unit">${escapeHtml(unit)}</span>` : ''}
            `;
            rowEl.addEventListener('click', () => scrollDocToParam(paramName, rowEl));
            paramsEl.appendChild(rowEl);
        }

        sectionEl.appendChild(toggleEl);
        sectionEl.appendChild(paramsEl);
        groupsEl.appendChild(sectionEl);
    }

    // ── Right panel: doc viewer ──────────────────────────────────────────
    renderDocViewer();

    // Show results card
    document.getElementById('step2').classList.add('hidden');
    document.getElementById('step3').classList.remove('hidden');
    setStepIndicator(4);
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

function renderDocViewer() {
    const viewer = document.getElementById('docViewer');
    if (!Object.keys(_allChunks).length) {
        viewer.innerHTML = '<div class="doc-placeholder">无文档文本数据</div>';
        return;
    }
    const pages = Object.keys(_allChunks).map(Number).sort((a, b) => a - b);
    viewer.innerHTML = pages.map(page => {
        const chunk = _allChunks[String(page)];
        // Support both old string format and new {text, source_file} format
        const text = typeof chunk === 'string' ? chunk : (chunk.text || '');
        const srcFile = typeof chunk === 'object' ? (chunk.source_file || '') : '';
        const fileLabel = srcFile ? `<span class="doc-file-badge">${escapeHtml(srcFile)}</span>` : '';
        return `<div class="doc-chunk" id="doc-page-${page}">
  <div class="doc-chunk-label">第 ${page} 页 ${fileLabel}</div>${escapeHtml(text)}</div>`;
    }).join('');
}

function scrollDocToParam(paramName, rowEl) {
    // Deactivate previously active row
    document.querySelectorAll('.param-row.active').forEach(r => r.classList.remove('active'));
    rowEl.classList.add('active');

    const paramData = _extractedData[paramName] || {};
    const sourceText = (paramData.source_text || '').trim();
    if (!sourceText) return;

    const viewer = document.getElementById('docViewer');

    // Re-render to get clean DOM, then highlight
    renderDocViewer();

    const searchStr = sourceText.substring(0, 80);
    const chunks = viewer.querySelectorAll('.doc-chunk');
    let foundChunk = null;

    for (const chunk of chunks) {
        const fullText = chunk.textContent;
        if (fullText.includes(searchStr)) {
            // Highlight match in innerHTML
            const label = chunk.querySelector('.doc-chunk-label');
            let rawHtml = chunk.innerHTML;
            const escapedSearch = escapeHtml(searchStr);
            const idx = rawHtml.indexOf(escapedSearch);
            if (idx !== -1) {
                rawHtml = rawHtml.slice(0, idx)
                    + `<mark class="doc-highlight">${escapedSearch}</mark>`
                    + rawHtml.slice(idx + escapedSearch.length);
                chunk.innerHTML = rawHtml;
            }
            foundChunk = chunk;
            break;
        }
    }

    if (foundChunk) {
        const target = foundChunk.querySelector('mark.doc-highlight') || foundChunk;
        target.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
}


function toggleAllSections(expand) {
    document.querySelectorAll('.section-toggle').forEach(t => {
        const paramsEl = t.nextElementSibling;
        if (!paramsEl) return;
        t.classList.toggle('collapsed', !expand);
        paramsEl.classList.toggle('collapsed', !expand);
    });
}

// ==================== DOWNLOAD CSV ====================
async function downloadCsv() {
    if (!state.extracted) return;

    const btn = document.querySelector('.btn-export');
    const originalText = btn.textContent;
    btn.textContent = '⏳ 生成中...';
    btn.disabled = true;

    try {
        const response = await fetch(`${API_BASE}/export-csv`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ extracted: state.extracted })
        });

        if (!response.ok) {
            const err = await response.json();
            throw new Error(err.error || 'CSV 生成失败');
        }

        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `变压器参数提取结果_${new Date().toISOString().split('T')[0]}.csv`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        btn.textContent = '✓ 下载成功';
        setTimeout(() => {
            btn.textContent = originalText;
            btn.disabled = false;
        }, 2000);

    } catch (err) {
        showError(err.message);
        btn.textContent = originalText;
        btn.disabled = false;
    }
}

// ==================== NAVIGATION ====================
function goBack() {
    document.getElementById('step2').classList.add('hidden');
    document.getElementById('step1').classList.remove('hidden');
    setStepIndicator(1);
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

function resetAll() {
    state = { files: [], productType: 'transformer', sessionId: null, section: null, extracted: null, totalPages: 0, filePageRanges: [] };
    document.getElementById('fileInput').value = '';
    document.getElementById('fileList').innerHTML = '';
    document.getElementById('btnLocate').disabled = true;
    document.getElementById('step3').classList.add('hidden');
    document.getElementById('step2').classList.add('hidden');
    document.getElementById('step1').classList.remove('hidden');
    setStepIndicator(1);
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ==================== LOADING STATE ====================
function showLoading(title, activeStep) {
    document.getElementById('step1').classList.add('hidden');
    document.getElementById('step2').classList.add('hidden');
    document.getElementById('step3').classList.add('hidden');
    document.getElementById('loadingState').classList.remove('hidden');
    document.getElementById('loadingTitle').textContent = title;

    // Reset all steps
    for (let i = 1; i <= 4; i++) {
        document.getElementById(`ls${i}`).textContent = '⏳';
    }
}

function hideLoading() {
    document.getElementById('loadingState').classList.add('hidden');
}

function setLoadingStep(stepNum, status) {
    const el = document.getElementById(`ls${stepNum}`);
    if (!el) return;
    if (status === 'active') el.textContent = '🔄';
    else if (status === 'done') el.textContent = '✅';
    else el.textContent = '⏳';
}

// ==================== STEP INDICATOR ====================
function setStepIndicator(currentStep) {
    for (let i = 1; i <= 4; i++) {
        const dot = document.getElementById(`step-dot-${i}`);
        if (!dot) continue;
        // Preserve nav-item class; set step state classes cleanly
        dot.className = 'nav-item step';
        if (i < currentStep) dot.classList.add('done');
        else if (i === currentStep) dot.classList.add('active');
    }

    // Update step connectors (step-line class added to nav-connectors in HTML)
    const lines = document.querySelectorAll('.step-line');
    lines.forEach((line, idx) => {
        line.classList.toggle('done', idx + 1 < currentStep);
    });
}

// ==================== ERROR ====================
function showError(msg) {
    document.getElementById('errorMsg').textContent = msg;
    document.getElementById('errorToast').classList.remove('hidden');
    setTimeout(hideError, 6000);
}

function hideError() {
    document.getElementById('errorToast').classList.add('hidden');
}

// ==================== UTILS ====================
function escapeHtml(text) {
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

// ==================== PARAM TOOLTIP ====================
function setupParamTooltip() {
    const tooltip = document.getElementById('paramTooltip');
    if (!tooltip) return;

    const groups = document.getElementById('paramGroups');

    groups.addEventListener('mouseover', (e) => {
        const row = e.target.closest('.param-row');
        if (!row) { tooltip.classList.add('hidden'); return; }

        const paramName = row.dataset.param;
        if (!paramName) return;

        const nameEl  = row.querySelector('.param-row-name');
        const valueEl = row.querySelector('.param-row-value');
        const onName  = e.target.closest('.param-row-name');
        const onValue = e.target.closest('.param-row-value');

        if (onName && nameEl && nameEl.scrollWidth > nameEl.clientWidth) {
            // Hovering truncated param name → show name only
            tooltip.textContent = paramName;
            tooltip.classList.remove('hidden');
            positionTooltip(tooltip, e);
        } else if (onValue && valueEl && valueEl.scrollWidth > valueEl.clientWidth) {
            // Hovering truncated value → show full value + unit
            const pd = _extractedData[paramName] || {};
            const valStr = String(pd.value ?? '') + (pd.unit ? '\u2002' + pd.unit : '');
            tooltip.textContent = valStr;
            tooltip.classList.remove('hidden');
            positionTooltip(tooltip, e);
        } else {
            // Not on a truncated element → hide immediately
            tooltip.classList.add('hidden');
        }
    });

    groups.addEventListener('mousemove', (e) => {
        if (!tooltip.classList.contains('hidden')) {
            positionTooltip(tooltip, e);
        }
    });

    groups.addEventListener('mouseleave', () => {
        tooltip.classList.add('hidden');
    });
}

function positionTooltip(tooltip, e) {
    const GAP = 10;
    const tw = tooltip.offsetWidth || 200;
    const th = tooltip.offsetHeight || 32;

    let x = e.clientX - tw / 2;
    let y = e.clientY - th - GAP;

    x = Math.max(8, Math.min(x, window.innerWidth - tw - 8));
    if (y < 8) y = e.clientY + GAP;

    tooltip.style.left = x + 'px';
    tooltip.style.top  = y + 'px';
}

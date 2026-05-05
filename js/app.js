// ============================================================
//  P&L PDF → Excel Converter — Client-Side Application
//  Dependencies: PDF.js (global pdfjsLib), SheetJS (global XLSX)
// ============================================================

// ── Live Clock ────────────────────────────────────────────────
(function startClock() {
    const timeEl = document.getElementById('clock-time');
    const dateEl = document.getElementById('clock-date');
    const DAY_NAMES  = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    const MON_NAMES  = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

    function tick() {
        const now = new Date();
        const h   = String(now.getHours()).padStart(2, '0');
        const m   = String(now.getMinutes()).padStart(2, '0');
        const s   = String(now.getSeconds()).padStart(2, '0');
        timeEl.textContent = `${h}:${m}:${s}`;
        dateEl.textContent = `${DAY_NAMES[now.getDay()]} ${now.getDate()} ${MON_NAMES[now.getMonth()]} ${now.getFullYear()}`;
    }

    tick();
    setInterval(tick, 1000);
})();

// Configure PDF.js worker (must match CDN version in index.html)
pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

// ── State ─────────────────────────────────────────────────────
const state = {
    files: [],         // Array of FileEntry objects
    isProcessing: false,
};

// FileEntry shape:
// { id: number, file: File, status: 'pending'|'processing'|'success'|'error', data: object|null, error: string|null }

// ── DOM References ─────────────────────────────────────────────
const $ = id => document.getElementById(id);

const dom = {
    apiKeyInput:    $('api-key-input'),
    btnToggleKey:   $('btn-toggle-key'),
    modelSelect:    $('model-select'),
    dropZone:       $('drop-zone'),
    fileInput:      $('file-input'),
    btnBrowse:      $('btn-browse'),
    fileQueue:      $('file-queue'),
    badgeCount:     $('badge-count'),
    fileList:       $('file-list'),
    btnClearAll:    $('btn-clear-all'),
    btnProcess:     $('btn-process'),
    progressBlock:  $('progress-block'),
    progressLabel:  $('progress-label'),
    progressFill:   $('progress-fill'),
    cardResults:    $('card-results'),
    statsStrip:     $('stats-strip'),
    resultsBody:    $('results-body'),
    btnDownload:    $('btn-download'),
    btnReset:       $('btn-reset'),
};

// ── Event Listeners ────────────────────────────────────────────

// Show/hide API key
dom.btnToggleKey.addEventListener('click', () => {
    dom.apiKeyInput.type = dom.apiKeyInput.type === 'password' ? 'text' : 'password';
});

// Browse button opens file dialog
dom.btnBrowse.addEventListener('click', e => {
    e.stopPropagation();
    dom.fileInput.click();
});

// Drop-zone also clickable
dom.dropZone.addEventListener('click', () => dom.fileInput.click());
dom.dropZone.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') dom.fileInput.click(); });

// File input change
dom.fileInput.addEventListener('change', e => {
    addFiles(e.target.files);
    e.target.value = '';  // allow re-selecting the same file
});

// Drag-and-drop
dom.dropZone.addEventListener('dragover', e => {
    e.preventDefault();
    dom.dropZone.classList.add('drag-over');
});
dom.dropZone.addEventListener('dragleave', e => {
    if (!dom.dropZone.contains(e.relatedTarget)) dom.dropZone.classList.remove('drag-over');
});
dom.dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dom.dropZone.classList.remove('drag-over');
    const pdfs = Array.from(e.dataTransfer.files).filter(f => f.type === 'application/pdf' || f.name.toLowerCase().endsWith('.pdf'));
    if (pdfs.length) addFiles(pdfs);
});

// Clear all files
dom.btnClearAll.addEventListener('click', () => {
    state.files = [];
    renderFileQueue();
    syncProcessBtn();
});

// Process button
dom.btnProcess.addEventListener('click', processAllFiles);

// Download button
dom.btnDownload.addEventListener('click', downloadExcel);

// Reset (process more)
dom.btnReset.addEventListener('click', () => {
    dom.cardResults.hidden = true;
    dom.progressBlock.hidden = true;
    // Reset only processed files back to pending for re-run, or allow fresh upload
    state.files = state.files.filter(f => f.status !== 'success');
    state.files.forEach(f => { f.status = 'pending'; f.error = null; });
    renderFileQueue();
    syncProcessBtn();
});

// ── File Management ────────────────────────────────────────────

function addFiles(fileList) {
    const incoming = Array.from(fileList).filter(
        f => f.type === 'application/pdf' || f.name.toLowerCase().endsWith('.pdf')
    );
    incoming.forEach(file => {
        const isDuplicate = state.files.some(
            e => e.file.name === file.name && e.file.size === file.size
        );
        if (!isDuplicate) {
            state.files.push({ id: Date.now() + Math.random(), file, status: 'pending', data: null, error: null });
        }
    });
    renderFileQueue();
    syncProcessBtn();
}

function removeFile(id) {
    state.files = state.files.filter(e => String(e.id) !== String(id));
    renderFileQueue();
    syncProcessBtn();
}

function updateEntry(id, updates) {
    const entry = state.files.find(e => String(e.id) === String(id));
    if (entry) Object.assign(entry, updates);
    renderFileQueue();
}

// ── Render: File Queue ─────────────────────────────────────────

function renderFileQueue() {
    const { files } = state;
    dom.fileQueue.hidden = files.length === 0;
    dom.badgeCount.textContent = files.length;

    dom.fileList.innerHTML = files.map(entry => {
        const statusEl = statusDotHTML(entry.status);
        const errorMsg = entry.status === 'error' && entry.error
            ? `<span class="file-error-msg" title="${escHtml(entry.error)}">&#9888; ${escHtml(entry.error)}</span>`
            : '';
        const removeBtn = (entry.status === 'pending' || entry.status === 'error') && !state.isProcessing
            ? `<button class="file-remove-btn" data-id="${entry.id}" title="Remove" aria-label="Remove ${escHtml(entry.file.name)}">&times;</button>`
            : '';
        return `
            <li class="file-item">
                ${statusEl}
                <span class="file-name" title="${escHtml(entry.file.name)}">${escHtml(entry.file.name)}</span>
                <span class="file-size">${formatSize(entry.file.size)}</span>
                ${errorMsg}
                ${removeBtn}
            </li>`;
    }).join('');

    // Attach remove button handlers
    dom.fileList.querySelectorAll('.file-remove-btn').forEach(btn => {
        btn.addEventListener('click', e => { e.stopPropagation(); removeFile(btn.dataset.id); });
    });
}

function statusDotHTML(status) {
    const labels = { pending: '●', success: '✓', error: '✕' };
    if (status === 'processing') {
        return `<span class="file-status processing" role="status" aria-label="Processing"></span>`;
    }
    return `<span class="file-status ${status}" aria-label="${status}">${labels[status] || '?'}</span>`;
}

function syncProcessBtn() {
    const hasPending = state.files.some(e => e.status === 'pending' || e.status === 'error');
    dom.btnProcess.disabled = !hasPending || state.isProcessing;
}

// ── PDF Text Extraction ────────────────────────────────────────

async function extractTextFromPDF(file) {
    const arrayBuffer = await file.arrayBuffer();
    const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
    const pdf = await loadingTask.promise;
    const pageTexts = [];
    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();
        pageTexts.push(content.items.map(item => item.str).join(' '));
    }
    return pageTexts.join('\n');
}

// ── Gemini API Call ────────────────────────────────────────────

async function callGemini(text, apiKey, model, maxRetries = 2) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    const prompt = `Return ONLY a valid JSON object. No markdown fences, no explanation, just raw JSON.
Extract all P&L (Profit & Loss) data from the document text below.

Required JSON keys:
- "company_name": string
- "period": string (e.g. "Q1 2024" or "Year ended Dec 31 2023")
- "currency": string (e.g. "USD", "EGP")
- "line_items": array of objects, each with:
    - "item_description": string
    - "amount": number (positive values only)
    - "category": "Revenue" or "Expense"

Rules:
- Exclude subtotals, grand totals, and net income/loss lines from line_items
- All amounts should be positive numbers (do not negate expenses)
- If currency cannot be determined use "N/A"

DOCUMENT TEXT:
${text.slice(0, 30000)}`;

    const payload = { contents: [{ parts: [{ text: prompt }] }] };

    for (let attempt = 0; attempt <= maxRetries; attempt++) {
        const res = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload),
        });

        if (res.ok) {
            const json = await res.json();
            const rawText = json?.candidates?.[0]?.content?.parts?.[0]?.text;
            if (!rawText) throw new Error('Empty response from Gemini API');

            // Strip markdown code fences if present, then extract JSON object
            const stripped = rawText.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
            const match = stripped.match(/\{[\s\S]*\}/);
            if (!match) throw new Error('No JSON object found in AI response');
            return JSON.parse(match[0]);
        }

        if (res.status === 429 && attempt < maxRetries) {
            setProgressLabel(`Rate limit reached — waiting 45 s before retry ${attempt + 1}/${maxRetries}…`);
            await sleep(45000);
            continue;
        }

        const errText = await res.text().catch(() => '');
        throw new Error(`Gemini API error ${res.status}: ${errText.slice(0, 200)}`);
    }
    throw new Error('Max retries exceeded');
}

// ── Main Processing ────────────────────────────────────────────

async function processAllFiles() {
    const apiKey = dom.apiKeyInput.value.trim();
    if (!apiKey) {
        alert('Please enter your Google Gemini API key in Step 1.');
        dom.apiKeyInput.focus();
        return;
    }

    const model = dom.modelSelect.value;
    const toProcess = state.files.filter(e => e.status === 'pending' || e.status === 'error');
    if (!toProcess.length) return;

    state.isProcessing = true;
    dom.btnProcess.disabled = true;
    dom.progressBlock.hidden = false;
    dom.cardResults.hidden = true;

    const total = toProcess.length;
    let done = 0;

    for (const entry of toProcess) {
        // 1. Extract text
        updateEntry(entry.id, { status: 'processing' });
        setProgressLabel(`[${done + 1}/${total}] Extracting text from "${entry.file.name}"…`);
        setProgressPct(done / total);

        let text;
        try {
            text = await extractTextFromPDF(entry.file);
        } catch (err) {
            updateEntry(entry.id, { status: 'error', error: `PDF read failed: ${err.message}` });
            done++;
            setProgressPct(done / total);
            continue;
        }

        if (!text.trim()) {
            updateEntry(entry.id, { status: 'error', error: 'No text found — PDF may be a scanned image' });
            done++;
            setProgressPct(done / total);
            continue;
        }

        // 2. AI extraction
        setProgressLabel(`[${done + 1}/${total}] Analyzing "${entry.file.name}" with AI…`);
        try {
            const data = await callGemini(text, apiKey, model);
            // Normalize line_items
            if (Array.isArray(data.line_items)) {
                data.line_items = data.line_items.map(item => ({
                    ...item,
                    amount: parseFloat(item.amount) || 0,
                    category: String(item.category || '').trim(),
                }));
            } else {
                data.line_items = [];
            }
            updateEntry(entry.id, { status: 'success', data });
        } catch (err) {
            updateEntry(entry.id, { status: 'error', error: err.message });
        }

        done++;
        setProgressPct(done / total);

        // Small delay between files to respect API rate limits
        if (done < total) {
            setProgressLabel(`Waiting before next file…`);
            await sleep(2000);
        }
    }

    setProgressLabel(`Done — processed ${done} file${done !== 1 ? 's' : ''}.`);
    setProgressPct(1);
    state.isProcessing = false;
    syncProcessBtn();
    renderResults();
}

function setProgressLabel(msg)     { dom.progressLabel.textContent = msg; }
function setProgressPct(ratio)     { dom.progressFill.style.width = `${Math.round(ratio * 100)}%`; }

// ── Render: Results Table ──────────────────────────────────────

function renderResults() {
    const successful = state.files.filter(e => e.status === 'success' && e.data);
    const failed     = state.files.filter(e => e.status === 'error');

    if (!successful.length && !failed.length) return;

    dom.cardResults.hidden = false;
    dom.resultsBody.innerHTML = '';

    let totalLineItems = 0;

    successful.forEach(entry => {
        const d = entry.data;
        const items = d.line_items || [];
        totalLineItems += items.length;

        const revenue  = sumByCategory(items, 'Revenue');
        const expenses = sumByCategory(items, 'Expense');
        const net      = revenue - expenses;

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${escHtml(d.company_name || '—')}</td>
            <td>${escHtml(d.period || '—')}</td>
            <td>${escHtml(d.currency || '—')}</td>
            <td class="col-num cell-revenue">${fmt(revenue)}</td>
            <td class="col-num">${fmt(expenses)}</td>
            <td class="col-num ${net >= 0 ? 'cell-net-pos' : 'cell-net-neg'}">${fmt(net)}</td>
            <td class="cell-filename" title="${escHtml(entry.file.name)}">${escHtml(entry.file.name)}</td>
        `;
        dom.resultsBody.appendChild(tr);
    });

    failed.forEach(entry => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td colspan="6" class="td-error">&#9888; ${escHtml(entry.file.name)}: ${escHtml(entry.error || 'Unknown error')}</td>
            <td><span class="status-pill err">Failed</span></td>
        `;
        dom.resultsBody.appendChild(tr);
    });

    // Stats strip
    dom.statsStrip.innerHTML = [
        { value: successful.length, label: 'Files Extracted' },
        { value: failed.length,     label: 'Failed' },
        { value: totalLineItems,    label: 'Line Items' },
    ].map(s => `
        <div class="stat-item">
            <span class="stat-value">${s.value}</span>
            <span class="stat-label">${s.label}</span>
        </div>
    `).join('');
}

// ── Excel Generation ───────────────────────────────────────────

function downloadExcel() {
    const successful = state.files.filter(e => e.status === 'success' && e.data);
    if (!successful.length) { alert('No successfully extracted files to export.'); return; }

    const wb = XLSX.utils.book_new();

    // ── Sheet 1: Summary ──
    const summaryData = [
        ['Company', 'Period', 'Currency', 'Total Revenue', 'Total Expenses', 'Net Income', 'Line Item Count', 'Source File']
    ];
    successful.forEach(entry => {
        const d     = entry.data;
        const items = d.line_items || [];
        const rev   = sumByCategory(items, 'Revenue');
        const exp   = sumByCategory(items, 'Expense');
        summaryData.push([
            d.company_name || '',
            d.period       || '',
            d.currency     || '',
            rev,
            exp,
            rev - exp,
            items.length,
            entry.file.name,
        ]);
    });
    const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
    summarySheet['!cols'] = [
        { wch: 30 }, { wch: 20 }, { wch: 10 }, { wch: 18 },
        { wch: 18 }, { wch: 18 }, { wch: 16 }, { wch: 40 },
    ];
    XLSX.utils.book_append_sheet(wb, summarySheet, 'Summary');

    // ── Sheet 2: All Line Items ──
    const lineData = [
        ['Company', 'Period', 'Currency', 'Item Description', 'Amount', 'Category', 'Source File']
    ];
    successful.forEach(entry => {
        const d = entry.data;
        (d.line_items || []).forEach(item => {
            lineData.push([
                d.company_name || '',
                d.period       || '',
                d.currency     || '',
                item.item_description || '',
                item.amount,
                item.category  || '',
                entry.file.name,
            ]);
        });
    });
    const lineSheet = XLSX.utils.aoa_to_sheet(lineData);
    lineSheet['!cols'] = [
        { wch: 30 }, { wch: 20 }, { wch: 10 }, { wch: 50 },
        { wch: 18 }, { wch: 12 }, { wch: 40 },
    ];
    XLSX.utils.book_append_sheet(wb, lineSheet, 'All_Line_Items');

    // ── Sheet per company (when multiple) ──
    if (successful.length > 1) {
        const usedNames = new Set(['Summary', 'All_Line_Items']);
        successful.forEach(entry => {
            const d        = entry.data;
            const baseName = (d.company_name || entry.file.name.replace(/\.pdf$/i, '')).slice(0, 28);
            let sheetName  = baseName;
            let suffix     = 2;
            while (usedNames.has(sheetName)) sheetName = `${baseName}_${suffix++}`;
            usedNames.add(sheetName);

            const rows = [
                ['Item Description', 'Amount', 'Category'],
                ...(d.line_items || []).map(i => [i.item_description || '', i.amount, i.category || ''])
            ];
            const sheet = XLSX.utils.aoa_to_sheet(rows);
            sheet['!cols'] = [{ wch: 50 }, { wch: 18 }, { wch: 12 }];
            XLSX.utils.book_append_sheet(wb, sheet, sheetName);
        });
    }

    const date     = new Date().toISOString().slice(0, 10);
    const fileName = `PnL_Report_${date}.xlsx`;
    XLSX.writeFile(wb, fileName);
}

// ── Utility Functions ──────────────────────────────────────────

function sumByCategory(items, category) {
    return items
        .filter(i => String(i.category).trim().toLowerCase() === category.toLowerCase())
        .reduce((sum, i) => sum + (parseFloat(i.amount) || 0), 0);
}

function fmt(n) {
    if (typeof n !== 'number' || isNaN(n)) return '—';
    return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatSize(bytes) {
    if (bytes < 1024)        return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1048576)).toFixed(1)} MB`;
}

function escHtml(str) {
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function sleep(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }

// ============================================================
//  CATEGORIZE DOCUMENTS — Collapsible Section
// ============================================================

// ── Classify State ─────────────────────────────────────────────
const clsState = {
    files: [],
    isProcessing: false,
};

// ── Classify DOM References ────────────────────────────────────
const clsDom = {
    toggleBtn:      $('btn-classify-toggle'),
    panel:          $('classify-panel'),
    apiKey:         $('cls-api-key'),
    btnToggleKey:   $('cls-btn-toggle-key'),
    modelSelect:    $('cls-model-select'),
    dropZone:       $('cls-drop-zone'),
    fileInput:      $('cls-file-input'),
    btnBrowse:      $('cls-btn-browse'),
    fileQueue:      $('cls-file-queue'),
    badgeCount:     $('cls-badge-count'),
    fileList:       $('cls-file-list'),
    btnClearAll:    $('cls-btn-clear-all'),
    btnProcess:     $('cls-btn-process'),
    progressBlock:  $('cls-progress-block'),
    progressLabel:  $('cls-progress-label'),
    progressFill:   $('cls-progress-fill'),
    results:        $('cls-results'),
    statsStrip:     $('cls-stats-strip'),
    resultsBody:    $('cls-results-body'),
    btnDownloadZip: $('cls-btn-download-zip'),
    btnReset:       $('cls-btn-reset'),
};

// ── Collapse / Expand ──────────────────────────────────────────
clsDom.toggleBtn.addEventListener('click', () => {
    const isOpen = clsDom.panel.classList.toggle('open');
    clsDom.toggleBtn.setAttribute('aria-expanded', String(isOpen));
});

// ── API Key Sync ───────────────────────────────────────────────
// Keeps the classify key and the P&L key in sync (both ways)
clsDom.apiKey.addEventListener('input', () => { dom.apiKeyInput.value = clsDom.apiKey.value; });
dom.apiKeyInput.addEventListener('input', () => { clsDom.apiKey.value = dom.apiKeyInput.value; });

clsDom.btnToggleKey.addEventListener('click', () => {
    clsDom.apiKey.type = clsDom.apiKey.type === 'password' ? 'text' : 'password';
});

// ── Classify File Management ───────────────────────────────────
clsDom.btnBrowse.addEventListener('click', e => { e.stopPropagation(); clsDom.fileInput.click(); });
clsDom.dropZone.addEventListener('click', () => clsDom.fileInput.click());
clsDom.dropZone.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') clsDom.fileInput.click(); });

clsDom.fileInput.addEventListener('change', e => { clsAddFiles(e.target.files); e.target.value = ''; });

clsDom.dropZone.addEventListener('dragover', e => { e.preventDefault(); clsDom.dropZone.classList.add('drag-over'); });
clsDom.dropZone.addEventListener('dragleave', e => {
    if (!clsDom.dropZone.contains(e.relatedTarget)) clsDom.dropZone.classList.remove('drag-over');
});
clsDom.dropZone.addEventListener('drop', e => {
    e.preventDefault();
    clsDom.dropZone.classList.remove('drag-over');
    const pdfs = Array.from(e.dataTransfer.files).filter(f => f.type === 'application/pdf' || f.name.toLowerCase().endsWith('.pdf'));
    if (pdfs.length) clsAddFiles(pdfs);
});

clsDom.btnClearAll.addEventListener('click', () => { clsState.files = []; clsRenderQueue(); clsSyncBtn(); });
clsDom.btnProcess.addEventListener('click', clsProcessAll);
clsDom.btnDownloadZip.addEventListener('click', clsDownloadZip);
clsDom.btnReset.addEventListener('click', () => {
    clsDom.results.hidden = true;
    clsDom.progressBlock.hidden = true;
    clsState.files = clsState.files.filter(e => e.status !== 'success');
    clsState.files.forEach(e => { e.status = 'pending'; e.error = null; e.data = null; });
    clsRenderQueue();
    clsSyncBtn();
});

function clsAddFiles(fileList) {
    Array.from(fileList)
        .filter(f => f.type === 'application/pdf' || f.name.toLowerCase().endsWith('.pdf'))
        .forEach(file => {
            if (!clsState.files.some(e => e.file.name === file.name && e.file.size === file.size)) {
                clsState.files.push({ id: Date.now() + Math.random(), file, status: 'pending', data: null, error: null });
            }
        });
    clsRenderQueue();
    clsSyncBtn();
}

function clsRemoveFile(id) {
    clsState.files = clsState.files.filter(e => String(e.id) !== String(id));
    clsRenderQueue();
    clsSyncBtn();
}

function clsUpdateEntry(id, updates) {
    const entry = clsState.files.find(e => String(e.id) === String(id));
    if (entry) Object.assign(entry, updates);
    clsRenderQueue();
}

// ── Classify Render: File Queue ────────────────────────────────
function clsRenderQueue() {
    clsDom.fileQueue.hidden = clsState.files.length === 0;
    clsDom.badgeCount.textContent = clsState.files.length;

    clsDom.fileList.innerHTML = clsState.files.map(entry => {
        const statusEl = statusDotHTML(entry.status);
        const errorMsg = entry.status === 'error' && entry.error
            ? `<span class="file-error-msg" title="${escHtml(entry.error)}">&#9888; ${escHtml(entry.error)}</span>`
            : '';
        const removeBtn = (entry.status === 'pending' || entry.status === 'error') && !clsState.isProcessing
            ? `<button class="file-remove-btn" data-id="${entry.id}" title="Remove">&times;</button>`
            : '';
        return `
            <li class="file-item">
                ${statusEl}
                <span class="file-name" title="${escHtml(entry.file.name)}">${escHtml(entry.file.name)}</span>
                <span class="file-size">${formatSize(entry.file.size)}</span>
                ${errorMsg}
                ${removeBtn}
            </li>`;
    }).join('');

    clsDom.fileList.querySelectorAll('.file-remove-btn').forEach(btn => {
        btn.addEventListener('click', e => { e.stopPropagation(); clsRemoveFile(btn.dataset.id); });
    });
}

function clsSyncBtn() {
    const hasPending = clsState.files.some(e => e.status === 'pending' || e.status === 'error');
    clsDom.btnProcess.disabled = !hasPending || clsState.isProcessing;
}

// ── Classify with Gemini ───────────────────────────────────────
async function classifyWithGemini(text, filename, apiKey, model, maxRetries = 2) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    const prompt = `You are a document classifier for DCT (Department of Culture and Tourism).
Classify the document into EXACTLY ONE of the following categories:

1. "Contracts"  — Legal agreements, service contracts, vendor agreements, event contracts,
                  purchase orders, NDAs, MoUs, terms & conditions, sponsorship agreements,
                  letters of intent, facility-use agreements
2. "P&L"        — Profit & Loss statements, financial reports, income statements,
                  revenue & expense summaries, financial performance reports, budget actuals
3. "Unrelated"  — Everything else: invitations, presentations, event programs, floor plans,
                  schedules, menus, photos or image descriptions, brochures, general
                  correspondence, press releases, social-media content, etc.

Document filename : ${escHtml(filename)}
Document text (excerpt):
${text.slice(0, 5000)}

Return ONLY valid JSON, no markdown, no explanation:
{"category": "Contracts" | "P&L" | "Unrelated", "reason": "one sentence"}`;

    const payload = { contents: [{ parts: [{ text: prompt }] }] };
    const VALID   = ['Contracts', 'P&L', 'Unrelated'];

    for (let attempt = 0; attempt <= maxRetries; attempt++) {
        const res = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload),
        });

        if (res.ok) {
            const json    = await res.json();
            const rawText = json?.candidates?.[0]?.content?.parts?.[0]?.text;
            if (!rawText) throw new Error('Empty response from Gemini');

            const stripped = rawText.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
            const match    = stripped.match(/\{[\s\S]*?\}/);
            if (!match) throw new Error('No JSON in AI response');

            const result   = JSON.parse(match[0]);
            const category = VALID.includes(result.category) ? result.category : 'Unrelated';
            return { category, reason: result.reason || '' };
        }

        if (res.status === 429 && attempt < maxRetries) {
            clsDom.progressLabel.textContent = `Rate limit — waiting 45 s (retry ${attempt + 1}/${maxRetries})…`;
            await sleep(45000);
            continue;
        }

        const errText = await res.text().catch(() => '');
        throw new Error(`API ${res.status}: ${errText.slice(0, 200)}`);
    }
    throw new Error('Max retries exceeded');
}

// ── Classify: Main Processing Loop ────────────────────────────
async function clsProcessAll() {
    const apiKey = clsDom.apiKey.value.trim() || dom.apiKeyInput.value.trim();
    if (!apiKey) {
        alert('Please enter your Google Gemini API key.');
        clsDom.apiKey.focus();
        return;
    }

    const model     = clsDom.modelSelect.value;
    const toProcess = clsState.files.filter(e => e.status === 'pending' || e.status === 'error');
    if (!toProcess.length) return;

    clsState.isProcessing = true;
    clsDom.btnProcess.disabled = true;
    clsDom.progressBlock.hidden = false;
    clsDom.results.hidden = true;

    const total = toProcess.length;
    let done = 0;

    for (const entry of toProcess) {
        clsUpdateEntry(entry.id, { status: 'processing' });
        clsDom.progressLabel.textContent = `[${done + 1}/${total}] Reading "${entry.file.name}"…`;
        clsDom.progressFill.style.width  = `${Math.round((done / total) * 100)}%`;

        let text;
        try {
            text = await extractTextFromPDF(entry.file);
        } catch (err) {
            clsUpdateEntry(entry.id, { status: 'error', error: `PDF read failed: ${err.message}` });
            done++;
            continue;
        }

        if (!text.trim()) {
            // No extractable text — mark as Unrelated with a note
            clsUpdateEntry(entry.id, {
                status: 'success',
                data: { category: 'Unrelated', reason: 'No text found (scanned image) — filed as Unrelated' },
            });
            done++;
            clsDom.progressFill.style.width = `${Math.round((done / total) * 100)}%`;
            continue;
        }

        clsDom.progressLabel.textContent = `[${done + 1}/${total}] Classifying "${entry.file.name}"…`;
        try {
            const result = await classifyWithGemini(text, entry.file.name, apiKey, model);
            clsUpdateEntry(entry.id, { status: 'success', data: result });
        } catch (err) {
            clsUpdateEntry(entry.id, { status: 'error', error: err.message });
        }

        done++;
        clsDom.progressFill.style.width = `${Math.round((done / total) * 100)}%`;

        if (done < total) {
            clsDom.progressLabel.textContent = `Waiting before next file…`;
            await sleep(1500);
        }
    }

    clsDom.progressLabel.textContent = `Done — classified ${done} file${done !== 1 ? 's' : ''}.`;
    clsDom.progressFill.style.width  = '100%';
    clsState.isProcessing = false;
    clsSyncBtn();
    clsRenderResults();
}

// ── Classify Render: Results Table ────────────────────────────
function clsRenderResults() {
    const successful = clsState.files.filter(e => e.status === 'success' && e.data);
    const failed     = clsState.files.filter(e => e.status === 'error');
    if (!successful.length && !failed.length) return;

    clsDom.results.hidden = false;
    clsDom.resultsBody.innerHTML = '';

    const counts = { Contracts: 0, 'P&L': 0, Unrelated: 0 };

    successful.forEach(entry => {
        const { category, reason } = entry.data;
        if (counts[category] !== undefined) counts[category]++;

        const badgeClass = { 'Contracts': 'contracts', 'P&L': 'pnl', 'Unrelated': 'unrelated' }[category] || 'unrelated';
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="cell-filename" title="${escHtml(entry.file.name)}">${escHtml(entry.file.name)}</td>
            <td><span class="cat-badge ${badgeClass}">${escHtml(category)}</span></td>
            <td style="font-size:.83rem;color:var(--clr-secondary)">${escHtml(reason)}</td>
            <td><span class="status-pill ok">&#10003; Done</span></td>
        `;
        clsDom.resultsBody.appendChild(tr);
    });

    failed.forEach(entry => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="cell-filename" title="${escHtml(entry.file.name)}">${escHtml(entry.file.name)}</td>
            <td colspan="2" class="td-error">&#9888; ${escHtml(entry.error || 'Unknown error')}</td>
            <td><span class="status-pill err">Failed</span></td>
        `;
        clsDom.resultsBody.appendChild(tr);
    });

    clsDom.statsStrip.innerHTML = [
        { value: successful.length,    label: 'Classified',  color: 'var(--clr-teal)' },
        { value: counts['Contracts'],  label: 'Contracts',   color: '#15803d' },
        { value: counts['P&L'],        label: 'P&amp;L',     color: 'var(--clr-primary)' },
        { value: counts['Unrelated'],  label: 'Unrelated',   color: 'var(--clr-secondary)' },
        { value: failed.length,        label: 'Failed',      color: 'var(--clr-danger)' },
    ].map(s => `
        <div class="stat-item">
            <span class="stat-value" style="color:${s.color}">${s.value}</span>
            <span class="stat-label">${s.label}</span>
        </div>
    `).join('');
}

// ── Classify: Download ZIP ─────────────────────────────────────
async function clsDownloadZip() {
    const successful = clsState.files.filter(e => e.status === 'success' && e.data);
    if (!successful.length) { alert('No classified files to export.'); return; }

    clsDom.btnDownloadZip.disabled = true;
    clsDom.btnDownloadZip.textContent = 'Building ZIP…';

    try {
        const ts   = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 15);
        const zip  = new JSZip();
        const root = zip.folder(`Output_${ts}`);

        const folders = {
            'Contracts': root.folder('Contracts'),
            'P&L':       root.folder('P&L'),
            'Unrelated': root.folder('Unrelated'),
        };

        for (const entry of successful) {
            const category = entry.data.category;
            const folder   = folders[category] || folders['Unrelated'];

            // Avoid filename collisions within a folder
            const existingNames = new Set(
                successful
                    .filter(e => e !== entry && e.data?.category === category)
                    .map(e => e.file.name)
            );
            let fname = entry.file.name;
            if (existingNames.has(fname)) {
                const [stem, ...extParts] = fname.split('.');
                const ext = extParts.length ? '.' + extParts.join('.') : '';
                let counter = 2;
                while (existingNames.has(`${stem}_${counter}${ext}`)) counter++;
                fname = `${stem}_${counter}${ext}`;
            }

            const buffer = await entry.file.arrayBuffer();
            folder.file(fname, buffer);
        }

        const blob   = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
        const url    = URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href  = url;
        anchor.download = `Output_${ts}.zip`;
        anchor.click();
        URL.revokeObjectURL(url);
    } finally {
        clsDom.btnDownloadZip.disabled = false;
        clsDom.btnDownloadZip.innerHTML = `
            <svg class="btn-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                <polyline points="7 10 12 15 17 10"/>
                <line x1="12" y1="15" x2="12" y2="3"/>
            </svg>
            Download Organised ZIP`;
    }
}

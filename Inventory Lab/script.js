// Set current date
document.addEventListener('DOMContentLoaded', async function() {
    const now = new Date();
    const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                   'July', 'August', 'September', 'October', 'November', 'December'];
    const dateStr = `${months[now.getMonth()]} ${now.getFullYear()}`;
    document.getElementById('currentDate').textContent = dateStr;

    // Toggle header (hide/show header + controls + sheet tabs) para table lang makita
    const toggleHeaderBtn = document.getElementById('toggleHeaderBtn');
    const container = document.querySelector('.container');
    if (toggleHeaderBtn && container) {
        toggleHeaderBtn.addEventListener('click', function() {
            const hidden = container.classList.toggle('header-hidden');
            toggleHeaderBtn.textContent = hidden ? 'Show header' : 'Hide header';
            toggleHeaderBtn.title = hidden ? 'Show header' : 'Hide header';
        });
    }
    
    if (isBackupMode) {
        document.body.classList.add('backup-mode');
        document.getElementById('addItemBtn').style.display = 'none';
        document.getElementById('addPCBtn').style.display = 'none';
        document.getElementById('newSheetBtn').style.display = 'none';
        document.getElementById('importExcelBtn').style.display = 'none';
        document.getElementById('importExcelInput').style.display = 'none';
        document.getElementById('clearBtn').style.display = 'none';
        document.getElementById('restoreBackupBtn').style.display = 'inline-flex';
        var sigSection = document.querySelector('.signatures-inventory');
        if (sigSection) sigSection.style.display = 'none';
        var sumSection = document.querySelector('.inventory-summary-section');
        if (sumSection) sumSection.style.display = 'none';
        (async function() {
            // Unahin ang localStorage snapshot — laging updated agad pag nag-save sa main page (add PC, add item, etc.)
            var snapshot = loadBackupSnapshot();
            if (!snapshot || !snapshot.sheets || Object.keys(snapshot.sheets).length === 0) {
                snapshot = loadBackupFromLocalStorage();
            }
            if (!snapshot || !snapshot.sheets || Object.keys(snapshot.sheets).length === 0) {
                if (typeof initSupabase === 'function') initSupabase();
                await new Promise(function(r) { setTimeout(r, 400); });
                for (var t = 0; t < 8 && !checkSupabaseConnection(); t++) {
                    await new Promise(function(r) { setTimeout(r, 300); });
                }
                if (checkSupabaseConnection()) {
                    for (var retry = 0; retry < 2; retry++) {
                        snapshot = await loadBackupFromSupabase();
                        if (snapshot && snapshot.sheets) {
                            var rows = Object.values(snapshot.sheets).reduce(function(n, s) { return n + (s.data ? s.data.length : 0); }, 0);
                            if (rows > 0) break;
                        }
                        await new Promise(function(r) { setTimeout(r, 500); });
                    }
                }
            }
        var totalRows = snapshot && snapshot.sheets ? Object.values(snapshot.sheets).reduce(function(n, s) { return n + (s.data ? s.data.length : 0); }, 0) : 0;
        if (snapshot && snapshot.sheets && Object.keys(snapshot.sheets).length > 0 && totalRows > 0) {
            sheets = snapshot.sheets;
            currentSheetId = snapshot.currentSheetId || Object.keys(sheets)[0];
            sheetCounter = snapshot.sheetCounter || 1;
            updateSheetTabs();
            var firstId = Object.keys(sheets).find(function(sid) {
                var s = sheets[sid];
                return s && s.data && s.data.length > 0;
            }) || Object.keys(sheets)[0];
            if (firstId) {
                currentSheetId = firstId;
                var sheet = sheets[firstId];
                displayData(sheet.data || [], true);
                makeTableReadOnly();
                document.getElementById('exportBtn').disabled = false;
            }
            document.querySelectorAll('.sheet-tab-menu').forEach(function(el) { el.style.display = 'none'; });
            document.querySelectorAll('.sheet-tab .close-sheet').forEach(function(el) { el.style.display = 'none'; });
        } else {
            document.getElementById('tableBody').innerHTML = '<tr class="empty-row"><td colspan="13" class="empty-message">No backup data yet. I-open muna ang main Inventory page (walang ?backup=1), hintayin mag-load ang data, tapos balik dito. O i-refresh ang main Inventory bago buksan ang Backup.</td></tr>';
        }
        document.getElementById('restoreBackupBtn').addEventListener('click', async function() {
            if (!confirm('Restore backup data to Inventory? This will replace current Inventory data.')) return;
            var snap = await loadBackupFromSupabase();
            if (!snap || !snap.sheets || Object.keys(snap.sheets).length === 0) snap = loadBackupSnapshot();
            if (!snap || !snap.sheets || Object.keys(snap.sheets).length === 0) snap = loadBackupFromLocalStorage();
            if (!snap || !snap.sheets) { alert('No backup data to restore.'); return; }
            sheets = snap.sheets;
            currentSheetId = snap.currentSheetId || Object.keys(sheets)[0];
            sheetCounter = snap.sheetCounter || 1;
            var backupStr = JSON.stringify({ sheets: sheets, currentSheetId: currentSheetId, sheetCounter: sheetCounter, savedAt: Date.now() });
            localStorage.setItem(BACKUP_KEY, backupStr);
            if (typeof syncToSupabase === 'function' && checkSupabaseConnection()) {
                syncToSupabase().then(function() {
                    window.location.href = window.location.pathname.replace(/[?].*$/, '') + '?restored=1';
                }).catch(function() {
                    window.location.href = window.location.pathname.replace(/[?].*$/, '') + '?restored=1';
                });
            } else {
                window.location.href = window.location.pathname.replace(/[?].*$/, '') + '?restored=1';
            }
        });
        })();
        return;
    }
    
    // Enable add buttons even when no data
    document.getElementById('addItemBtn').disabled = false;
    document.getElementById('addPCBtn').disabled = false;
    
    setupRowAddListeners();
    
    // Check Supabase connection after a short delay to ensure config.js is loaded
    setTimeout(async () => {
        // Try to initialize Supabase if not already done
        if (typeof initSupabase === 'function') {
            initSupabase();
        }
        
        // Wait a bit more for Supabase to initialize, with retry logic
        let initAttempts = 0;
        const maxAttempts = 10;
        const checkSupabaseInit = setInterval(async () => {
            initAttempts++;
            if (checkSupabaseConnection()) {
                console.log('✅ Supabase connected, loading data...');
                clearInterval(checkSupabaseInit);
                var restored = window.location.search.indexOf('restored=1') !== -1;
                if (restored) {
                    var rest = loadBackupFromLocalStorage();
                    if (rest && rest.sheets && Object.keys(rest.sheets).length > 0) {
                        sheets = rest.sheets;
                        currentSheetId = rest.currentSheetId || Object.keys(sheets)[0];
                        sheetCounter = rest.sheetCounter || 1;
                        updateSheetTabs();
                        var fid = Object.keys(sheets)[0];
                        if (fid) { currentSheetId = fid; switchToSheet(fid); }
                        syncToSupabase();
                        if (hasAnyData()) {
                            saveBackupToLocalStorage();
                            syncBackupToSupabase();
                        }
                        if (window.history && window.history.replaceState) {
                            window.history.replaceState({}, '', window.location.pathname + window.location.hash);
                        }
                    } else {
                        await loadFromSupabase();
                    }
                } else {
                    await loadFromSupabase();
                }
                var hasData = hasAnyData();
                if (hasData) {
                    console.log('✅ Data loaded successfully' + (restored ? ' (restored from backup)' : ' from Supabase'));
                } else {
                    console.log('ℹ️ No existing data found in Supabase (this is normal for first use)');
                }
            } else if (initAttempts >= maxAttempts) {
                console.warn('⚠️ Supabase initialization timeout. Using local storage only.');
                clearInterval(checkSupabaseInit);
            }
        }, 500); // Check every 500ms
    }, 200);
});

const BACKUP_KEY = 'inventory_lab_backup';
const BACKUP_SNAPSHOT_KEY = 'inventory_lab_backup_snapshot'; // Hindi naaapektuhan ng Clear Data; gamit sa Backup page

function saveBackupToLocalStorage() {
    if (isBackupMode) return;
    try {
        const backup = {
            sheets: sheets,
            currentSheetId: currentSheetId,
            sheetCounter: sheetCounter,
            savedAt: Date.now()
        };
        localStorage.setItem(BACKUP_KEY, JSON.stringify(backup));
        if (hasAnyData()) {
            localStorage.setItem(BACKUP_SNAPSHOT_KEY, JSON.stringify(backup));
        }
    } catch (e) {
        if (e.name === 'QuotaExceededError') {
            try {
                const slim = { sheets: {}, currentSheetId, sheetCounter, savedAt: Date.now() };
                Object.keys(sheets).forEach(sid => {
                    const s = sheets[sid];
                    slim.sheets[sid] = { id: s.id, name: s.name, data: s.data, hasData: s.hasData, highlightStates: s.highlightStates || [], pictureUrls: [] };
                });
                localStorage.setItem(BACKUP_KEY, JSON.stringify(slim));
            } catch (e2) { }
        }
    }
}

function loadBackupFromLocalStorage() {
    try {
        const raw = localStorage.getItem(BACKUP_KEY);
        if (!raw) return null;
        const backup = JSON.parse(raw);
        if (!backup || !backup.sheets) return null;
        return backup;
    } catch (e) { return null; }
}

function loadBackupSnapshot() {
    try {
        const raw = localStorage.getItem(BACKUP_SNAPSHOT_KEY);
        if (!raw) return null;
        const backup = JSON.parse(raw);
        if (!backup || !backup.sheets) return null;
        return backup;
    } catch (e) { return null; }
}

function makeTableReadOnly() {
    var tbody = document.getElementById('tableBody');
    if (!tbody) return;
    tbody.querySelectorAll('input').forEach(function(inp) { inp.disabled = true; inp.readOnly = true; });
    tbody.querySelectorAll('.actions-cell').forEach(function(cell) {
        cell.innerHTML = '';
        cell.textContent = '—';
    });
    // Keep first column visible but empty (same layout as main Inventory page)
    tbody.querySelectorAll('.row-add-cell').forEach(function(cell) {
        if (cell) { cell.innerHTML = ''; cell.style.display = ''; }
    });
    document.querySelectorAll('#inventoryTable .col-add').forEach(function(cell) {
        if (cell) cell.style.display = '';
    });
    // Hide Delete section button in PC header so it doesn't overflow into Picture column
    tbody.querySelectorAll('.pc-header-row .delete-btn').forEach(function(btn) {
        if (btn) btn.style.display = 'none';
    });
    tbody.querySelectorAll('.picture-cell button').forEach(function(btn) { btn.style.display = 'none'; });
}

// Save bago mag-close/refresh — localStorage muna (sync, instant), tapos Supabase (async, best-effort)
function flushSaveBeforeUnload() {
    saveCurrentSheetData(true);
    saveBackupToLocalStorage();
    if (syncTimeout) {
        clearTimeout(syncTimeout);
        syncTimeout = null;
    }
    if (checkSupabaseConnection() && !isSyncing) {
        syncToSupabase();
    }
}
window.addEventListener('beforeunload', function() {
    flushSaveBeforeUnload();
});

// Sheet management
let sheets = {
    'sheet-1': {
        id: 'sheet-1',
        name: 'Sheet 1',
        data: [],
        hasData: false,
        highlightStates: [],
        pictureUrls: []
    }
};
let currentSheetId = 'sheet-1';
let sheetCounter = 1;
let isLoadingFromSupabase = false; // Flag to prevent sync during load
var isBackupMode = typeof window !== 'undefined' && window.location && window.location.search.indexOf('backup=1') !== -1;
let isSyncing = false; // Flag to prevent concurrent syncs
let syncTimeout = null; // For debouncing sync calls
let originalWorkbook = null;

// Get current sheet data
function getCurrentSheet() {
    return sheets[currentSheetId];
}

// Set current sheet data
function setCurrentSheetData(data, hasDataFlag = true) {
    sheets[currentSheetId].data = data;
    sheets[currentSheetId].hasData = hasDataFlag;
}

// Apply condition-based row color
function applyConditionColor(row, condition) {
    // Remove existing condition classes
    row.classList.remove('condition-borrowed', 'condition-unserviceable');
    
    if (condition === 'Borrowed') {
        row.classList.add('condition-borrowed');
    } else if (condition === 'Unserviceable') {
        row.classList.add('condition-unserviceable');
    }
}

// Column order: must match table header and Excel export (Article/It, Description, ..., User)
const DATA_COLUMN_ORDER = ['Article/It', 'Description', 'Old Property N Assigned', 'Unit of meas', 'Unit Value', 'Quantity per Physical count', 'Location/Whereabout', 'Condition', 'Remarks', 'User'];
const UNIT_MEAS_COL = 3;
const UNIT_VALUE_COL = 4;
const USER_COL = 9;
const QUANTITY_COL = 5;

// Summary by Article/Item: Item, Unit measure, Existing, Years (2026–2029), For Disposal, Remarks
function computeSummaryData() {
    const sheet = getCurrentSheet();
    const data = (sheet && sheet.data) ? sheet.data : [];
    const map = {};
    data.forEach(function(row) {
        if (!row || row.length < 10) return;
        const articleRaw = (row[0] != null ? String(row[0]).trim() : '');
        if (!articleRaw) return;
        const item = toTitleCase(articleRaw);
        const unitMeas = (row[UNIT_MEAS_COL] != null ? String(row[UNIT_MEAS_COL]).trim() : '') || '—';
        const qty = parseInt(row[QUANTITY_COL], 10) || 0;
        if (!map[item]) {
            map[item] = { item: item, unitMeasure: unitMeas || '—', existing: 0, y2026: 0, y2027: 0, y2028: 0, y2029: 0, forDisposal: 0, remarks: '' };
        }
        map[item].existing += qty;
        if ((unitMeas || '') !== '' && (map[item].unitMeasure === '—' || !map[item].unitMeasure)) map[item].unitMeasure = unitMeas;
    });
    return Object.keys(map).sort().map(function(k) { return map[k]; });
}

function renderSummaryTable() {
    const tbody = document.getElementById('inventorySummaryBody');
    if (!tbody) return;
    saveSummaryExtra();
    const rows = computeSummaryData();
    const sheet = getCurrentSheet();
    const summaryExtra = (sheet && sheet.summaryExtra) ? sheet.summaryExtra : {};
    tbody.innerHTML = '';
    if (rows.length === 0) {
        const tr = document.createElement('tr');
        tr.innerHTML = '<td colspan="9" class="summary-empty">No data — add items or import Excel to see summary.</td>';
        tbody.appendChild(tr);
        return;
    }
    rows.forEach(function(r) {
        const extra = summaryExtra[r.item] || {};
        const tr = document.createElement('tr');
        tr.setAttribute('data-summary-item', r.item);
        const y2026 = extra.y2026 != null ? extra.y2026 : '';
        const y2027 = extra.y2027 != null ? extra.y2027 : '';
        const y2028 = extra.y2028 != null ? extra.y2028 : '';
        const y2029 = extra.y2029 != null ? extra.y2029 : '';
        const forD = extra.forDisposal != null ? extra.forDisposal : '';
        const rem = (extra.remarks != null ? extra.remarks : '') || '';
        tr.innerHTML =
            '<td>' + (r.item || '') + '</td>' +
            '<td>' + (r.unitMeasure || '—') + '</td>' +
            '<td>' + (r.existing || 0) + '</td>' +
            '<td><input type="number" class="summary-year" data-year="2026" min="0" value="' + y2026 + '" placeholder="0"></td>' +
            '<td><input type="number" class="summary-year" data-year="2027" min="0" value="' + y2027 + '" placeholder="0"></td>' +
            '<td><input type="number" class="summary-year" data-year="2028" min="0" value="' + y2028 + '" placeholder="0"></td>' +
            '<td><input type="number" class="summary-year" data-year="2029" min="0" value="' + y2029 + '" placeholder="0"></td>' +
            '<td><input type="number" class="summary-disposal" min="0" value="' + forD + '" placeholder="0"></td>' +
            '<td><input type="text" class="summary-remarks" value="' + (rem.replace(/"/g, '&quot;')) + '" placeholder="Remarks"></td>';
        tbody.appendChild(tr);
    });
    tbody.querySelectorAll('.summary-year, .summary-disposal, .summary-remarks').forEach(function(inp) {
        inp.addEventListener('change', saveSummaryExtra);
        inp.addEventListener('blur', saveSummaryExtra);
    });
}

function saveSummaryExtra() {
    const tbody = document.getElementById('inventorySummaryBody');
    const sheet = getCurrentSheet();
    if (!tbody || !sheet) return;
    sheet.summaryExtra = sheet.summaryExtra || {};
    var rows = tbody.querySelectorAll('tr[data-summary-item]');
    if (rows.length === 0) return;
    rows.forEach(function(tr) {
        const item = tr.getAttribute('data-summary-item');
        if (!item) return;
        const y2026 = tr.querySelector('.summary-year[data-year="2026"]');
        const y2027 = tr.querySelector('.summary-year[data-year="2027"]');
        const y2028 = tr.querySelector('.summary-year[data-year="2028"]');
        const y2029 = tr.querySelector('.summary-year[data-year="2029"]');
        const disposal = tr.querySelector('.summary-disposal');
        const remarks = tr.querySelector('.summary-remarks');
        sheet.summaryExtra[item] = {
            y2026: y2026 ? y2026.value : '',
            y2027: y2027 ? y2027.value : '',
            y2028: y2028 ? y2028.value : '',
            y2029: y2029 ? y2029.value : '',
            forDisposal: disposal ? disposal.value : '',
            remarks: remarks ? remarks.value : ''
        };
    });
    saveBackupToLocalStorage();
}

// Per PC section: merge Unit of meas, Unit Value, at User vertically; walang kulay ang columns na ito
function mergeUnitColumnsInTable() {
    const tbody = document.getElementById('tableBody');
    if (!tbody) return;
    const rows = Array.from(tbody.querySelectorAll('tr:not(.empty-row)'));
    const sections = [];
    let group = [];
    for (let i = 0; i < rows.length; i++) {
        if (rows[i].classList.contains('pc-header-row')) {
            if (group.length > 0) {
                sections.push(group);
                group = [];
            }
        } else {
            group.push(i);
        }
    }
    if (group.length > 0) sections.push(group);

    for (const indices of sections) {
        if (indices.length === 0) continue;
        const firstIdx = indices[0];
        const firstRow = rows[firstIdx];
        const unitMeasTd = firstRow.querySelector(`td.editable[data-column="${UNIT_MEAS_COL}"]`);
        const unitValueTd = firstRow.querySelector(`td.editable[data-column="${UNIT_VALUE_COL}"]`);
        const userTd = firstRow.querySelector(`td.editable[data-column="${USER_COL}"]`);
        if (!unitMeasTd || !unitValueTd) continue;
        const n = indices.length;
        unitMeasTd.rowSpan = n;
        unitValueTd.rowSpan = n;
        if (userTd) {
            userTd.rowSpan = n;
            userTd.classList.add('col-no-bg');
        }
        // Picture column: hindi na merged — bawat row may sariling picture cell, pwede mag-upload per row
        unitMeasTd.classList.add('col-no-bg');
        unitValueTd.classList.add('col-no-bg');
        for (let j = 1; j < n; j++) {
            const row = rows[indices[j]];
            [USER_COL, UNIT_VALUE_COL, UNIT_MEAS_COL].forEach(col => {
                const td = row.querySelector(`td.editable[data-column="${col}"]`);
                if (td) td.remove();
            });
        }
    }
}

// Create editable cell
function createEditableCell(value, isPCHeader = false, cellIndex = -1, row = null) {
    const td = document.createElement('td');
    td.classList.add('editable');
    if (cellIndex >= 0) td.setAttribute('data-column', String(cellIndex));
    if (cellIndex === UNIT_MEAS_COL || cellIndex === UNIT_VALUE_COL || cellIndex === USER_COL) td.classList.add('col-no-bg');
    
    const input = document.createElement('input');
    input.type = 'text';
    // Unit Value: strip existing ₱ so we don't double the prefix; stored value is number/text only
    if (cellIndex === UNIT_VALUE_COL) {
        input.value = (value || '').toString().replace(/^₱\s*/i, '').trim();
    } else if (cellIndex === 0) {
        input.value = toTitleCase(value || '');
    } else {
        input.value = value || '';
    }
    if (cellIndex === UNIT_MEAS_COL) input.placeholder = 'e.g., pcs, set, unit';
    if (cellIndex === UNIT_VALUE_COL) input.placeholder = '5,000.00';
    if (cellIndex === USER_COL) input.placeholder = 'e.g., MR TO ANGELINA C. PAQUIBOT';
    
    // Article/It (index 0): auto Title Case (capital each word)
    const isAutoTitleCase = cellIndex === 0;
    // Description (index 1): auto uppercase
    const isAutoUppercase = cellIndex === 1;
    if (isAutoTitleCase) {
        input.addEventListener('input', function() {
            const cursorPos = this.selectionStart;
            this.value = toTitleCase(this.value);
            this.setSelectionRange(cursorPos, cursorPos);
        });
        input.addEventListener('paste', function(e) {
            e.preventDefault();
            const pastedText = toTitleCase((e.clipboardData || window.clipboardData).getData('text'));
            const start = this.selectionStart;
            const end = this.selectionEnd;
            this.value = this.value.substring(0, start) + pastedText + this.value.substring(end);
            this.setSelectionRange(start + pastedText.length, start + pastedText.length);
        });
    }
    if (isAutoUppercase) {
        input.style.textTransform = 'uppercase';
        input.addEventListener('input', function() {
            const cursorPos = this.selectionStart;
            this.value = this.value.toUpperCase();
            this.setSelectionRange(cursorPos, cursorPos);
        });
        input.addEventListener('paste', function(e) {
            e.preventDefault();
            const pastedText = (e.clipboardData || window.clipboardData).getData('text').toUpperCase();
            const start = this.selectionStart;
            const end = this.selectionEnd;
            this.value = this.value.substring(0, start) + pastedText + this.value.substring(end);
            this.setSelectionRange(start + pastedText.length, start + pastedText.length);
        });
    }
    
    input.addEventListener('blur', function() {
        if (isAutoTitleCase) this.value = toTitleCase(this.value);
        if (isAutoUppercase) this.value = this.value.toUpperCase();
        updateDataFromTable();
        // If this is the condition cell (index 7), update row color
        if (cellIndex === 7 && row) {
            applyConditionColor(row, input.value);
        }
    });
    input.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            this.blur();
        }
    });
    
    // Unit Value: show peso sign prefix automatically
    if (cellIndex === UNIT_VALUE_COL) {
        const wrapper = document.createElement('span');
        wrapper.className = 'unit-value-wrapper';
        const prefix = document.createElement('span');
        prefix.className = 'unit-value-prefix';
        prefix.textContent = '₱ ';
        wrapper.appendChild(prefix);
        wrapper.appendChild(input);
        td.appendChild(wrapper);
    } else {
        td.appendChild(input);
    }
    
    if (isPCHeader) {
        td.style.fontWeight = 'bold';
        input.style.fontWeight = 'bold';
    }
    
    return td;
}

// Create action cell with delete and highlight buttons
function createActionCell() {
    const td = document.createElement('td');
    td.classList.add('actions-cell');
    
    const buttonContainer = document.createElement('div');
    buttonContainer.style.display = 'flex';
    buttonContainer.style.gap = '5px';
    buttonContainer.style.flexWrap = 'wrap';
    buttonContainer.style.justifyContent = 'center';
    buttonContainer.style.alignItems = 'center';
    
    // Highlight button
    const highlightBtn = document.createElement('button');
    highlightBtn.className = 'highlight-btn';
    highlightBtn.textContent = '⭐ Highlight';
    highlightBtn.addEventListener('click', function() {
        const tr = this.closest('tr');
        if (tr && !tr.classList.contains('pc-header-row') && !tr.classList.contains('empty-row')) {
            if (tr.classList.contains('highlighted-row')) {
                tr.classList.remove('highlighted-row');
                highlightBtn.textContent = '⭐ Highlight';
                highlightBtn.style.background = '#ffc107';
            } else {
                tr.classList.add('highlighted-row');
                highlightBtn.textContent = '✅ Highlighted';
                highlightBtn.style.background = '#28a745';
            }
            updateDataFromTable();
        }
    });
    
    // Delete button
    const deleteBtn = document.createElement('button');
    deleteBtn.className = 'delete-btn';
    deleteBtn.textContent = '🗑️ Delete';
    deleteBtn.addEventListener('click', function() {
        if (confirm('Are you sure you want to delete this row?')) {
            const tr = this.closest('tr');
            if (tr) {
                tr.remove();
                mergeUnitColumnsInTable();
                updateDataFromTable();
                
                // Check if table is now empty
                const tbody = document.getElementById('tableBody');
                if (tbody.children.length === 0) {
                    tbody.innerHTML = '<tr class="empty-row"><td colspan="13" class="empty-message">No data loaded. Please import an Excel file or add items manually.</td></tr>';
                    setCurrentSheetData([], false);
                    document.getElementById('exportBtn').disabled = !hasAnyData();
                    document.getElementById('clearBtn').disabled = true;
                } else {
                    updateDataFromTable();
                }
            }
        }
    });
    
    buttonContainer.appendChild(highlightBtn);
    buttonContainer.appendChild(deleteBtn);
    td.appendChild(buttonContainer);
    return td;
}

// Pending insert from row-add dropdown: { refRow, position } — used when Add Item form submits
let pendingRowInsert = null;

// Create add-cell with + button (shows when row is clicked). Click + → Above/Below → Add Item / Add PC Section.
function createAddCell(tr) {
    const td = document.createElement('td');
    td.classList.add('row-add-cell');
    const plusBtn = document.createElement('button');
    plusBtn.type = 'button';
    plusBtn.className = 'row-add-btn';
    plusBtn.innerHTML = '+';
    plusBtn.title = 'Add above or below';
    plusBtn.setAttribute('aria-label', 'Add row above or below');
    
    const dropdown = document.createElement('div');
    dropdown.className = 'row-add-dropdown';
    
    // Level 1: Above / Below
    const aboveOpt = document.createElement('button');
    aboveOpt.type = 'button';
    aboveOpt.textContent = 'Above';
    aboveOpt.className = 'row-add-option';
    const belowOpt = document.createElement('button');
    belowOpt.type = 'button';
    belowOpt.textContent = 'Below';
    belowOpt.className = 'row-add-option';
    
    // Level 2: Add Item / Add PC Section (reuse same container)
    const addItemOpt = document.createElement('button');
    addItemOpt.type = 'button';
    addItemOpt.textContent = 'Add Item';
    addItemOpt.className = 'row-add-option row-add-sub';
    addItemOpt.style.display = 'none';
    const addPCOpt = document.createElement('button');
    addPCOpt.type = 'button';
    addPCOpt.textContent = 'Add PC Section';
    addPCOpt.className = 'row-add-option row-add-sub';
    addPCOpt.style.display = 'none';
    
    function showFirstLevel() {
        aboveOpt.style.display = 'block';
        belowOpt.style.display = 'block';
        addItemOpt.style.display = 'none';
        addPCOpt.style.display = 'none';
    }
    function showSecondLevel(position) {
        aboveOpt.style.display = 'none';
        belowOpt.style.display = 'none';
        addItemOpt.style.display = 'block';
        addPCOpt.style.display = 'block';
        addItemOpt.dataset.position = position;
        addPCOpt.dataset.position = position;
    }
    
    plusBtn.addEventListener('click', function(e) {
        e.stopPropagation();
        document.querySelectorAll('.row-add-dropdown.show').forEach(d => {
            d.classList.remove('show');
            // Reset to first level (Above/Below): first 2 options visible, last 2 hidden
            const opts = d.querySelectorAll('.row-add-option');
            opts.forEach((o, i) => { o.style.display = i < 2 ? 'block' : 'none'; });
        });
        showFirstLevel();
        dropdown.classList.toggle('show');
        if (dropdown.classList.contains('show')) {
            const rect = plusBtn.getBoundingClientRect();
            dropdown.style.top = (rect.bottom + 4) + 'px';
            dropdown.style.left = rect.left + 'px';
        }
    });
    
    aboveOpt.addEventListener('click', function(e) {
        e.stopPropagation();
        showSecondLevel('above');
    });
    belowOpt.addEventListener('click', function(e) {
        e.stopPropagation();
        showSecondLevel('below');
    });
    
    addItemOpt.addEventListener('click', function(e) {
        e.stopPropagation();
        dropdown.classList.remove('show');
        pendingRowInsert = { refRow: tr, position: addItemOpt.dataset.position || 'below' };
        const modal = document.getElementById('addItemModal');
        if (modal) {
            modal.classList.add('show');
            const articleEl = document.getElementById('article');
            if (articleEl) articleEl.focus();
        }
    });
    addPCOpt.addEventListener('click', function(e) {
        e.stopPropagation();
        dropdown.classList.remove('show');
        const pos = addPCOpt.dataset.position || 'below';
        addPCSectionAt(tr, pos);
    });
    
    dropdown.appendChild(aboveOpt);
    dropdown.appendChild(belowOpt);
    dropdown.appendChild(addItemOpt);
    dropdown.appendChild(addPCOpt);
    td.appendChild(plusBtn);
    td.appendChild(dropdown);
    return td;
}

// Row click: show + for that row
function setupRowAddListeners() {
    const tbody = document.getElementById('tableBody');
    if (!tbody) return;
    tbody.addEventListener('click', function(e) {
        const tr = e.target.closest('tr');
        if (!tr || tr.classList.contains('empty-row')) return;
        if (e.target.closest('button')) return; // Don't select when clicking buttons
        document.querySelectorAll('#tableBody tr.row-selected').forEach(r => r.classList.remove('row-selected'));
        tr.classList.add('row-selected');
    });
}

// Add empty item row at position (above or below refRow)
function addItemRowAt(refRow, position) {
    const tbody = document.getElementById('tableBody');
    const emptyRow = tbody.querySelector('.empty-row');
    if (emptyRow) emptyRow.remove();
    
    const tr = document.createElement('tr');
    for (let j = 0; j < 10; j++) {
        const td = createEditableCell('', false, j, tr);
        tr.appendChild(td);
    }
    const pictureCell = createPictureCell(null);
    tr.appendChild(pictureCell);
    const actionCell = createActionCell();
    tr.appendChild(actionCell);
    
    const addCell = createAddCell(tr);
    tr.insertBefore(addCell, tr.firstChild);
    
    if (position === 'above') {
        tbody.insertBefore(tr, refRow);
    } else {
        tbody.insertBefore(tr, refRow.nextElementSibling);
    }
    mergeUnitColumnsInTable();
    setCurrentSheetData(getCurrentSheet().data, true);
    document.getElementById('exportBtn').disabled = false;
    document.getElementById('clearBtn').disabled = false;
    updateDataFromTable();
    const firstInput = tr.querySelector('td input');
    if (firstInput) firstInput.focus();
}

// Add PC section row above or below refRow
function addPCSectionAt(refRow, position) {
    const pcName = prompt('Enter PC Section Name (e.g., "PC USED BY: JOVEN T. CRUZ" or "SERVER"):');
    if (pcName === null) return;
    
    const tbody = document.getElementById('tableBody');
    const emptyRow = tbody.querySelector('.empty-row');
    if (emptyRow) emptyRow.remove();
    
    const tr = document.createElement('tr');
    tr.classList.add('pc-header-row');
    const addCell = createAddCell(tr);
    tr.appendChild(addCell);
    
    const td = document.createElement('td');
    td.colSpan = 11;
    td.className = 'pc-name-cell';
    td.style.fontWeight = 'bold';
    td.style.fontSize = '14px';
    const wrapper = document.createElement('div');
    wrapper.style.display = 'flex';
    wrapper.style.alignItems = 'center';
    wrapper.style.justifyContent = 'space-between';
    wrapper.style.gap = '10px';
    wrapper.style.width = '100%';
    const input = document.createElement('input');
    input.type = 'text';
    input.value = pcName;
    input.placeholder = 'PC name (e.g. PC 1, PC USED BY: NAME)';
    input.style.flex = '1';
    input.style.textAlign = 'center';
    input.style.fontWeight = 'bold';
    input.style.fontSize = '14px';
    input.style.border = 'none';
    input.style.background = 'transparent';
    input.addEventListener('blur', () => updateDataFromTable());
    input.addEventListener('keypress', (e) => { if (e.key === 'Enter') input.blur(); });
    const deleteSectionBtn = document.createElement('button');
    deleteSectionBtn.type = 'button';
    deleteSectionBtn.textContent = 'Delete section';
    deleteSectionBtn.className = 'delete-btn';
    deleteSectionBtn.style.flexShrink = '0';
    deleteSectionBtn.addEventListener('click', function() {
        if (confirm('Delete this PC section and all items under it?')) {
            tr.remove();
            mergeUnitColumnsInTable();
            updateDataFromTable();
            if (tbody.querySelectorAll('tr:not(.empty-row)').length === 0) {
                tbody.innerHTML = '<tr class="empty-row"><td colspan="13" class="empty-message">No data loaded. Please import an Excel file or add items manually.</td></tr>';
                setCurrentSheetData([], false);
                document.getElementById('clearBtn').disabled = true;
            }
        }
    });
    wrapper.appendChild(input);
    wrapper.appendChild(deleteSectionBtn);
    td.appendChild(wrapper);
    tr.appendChild(td);
    var actionsTd = document.createElement('td');
    actionsTd.classList.add('actions-cell');
    tr.appendChild(actionsTd);
    
    if (position === 'above') {
        tbody.insertBefore(tr, refRow);
    } else {
        tbody.insertBefore(tr, refRow.nextElementSibling);
    }
    mergeUnitColumnsInTable();
    updateDataFromTable();
    document.getElementById('exportBtn').disabled = false;
    document.getElementById('clearBtn').disabled = false;
}

// Add new item row
function addItemRow(isPCHeader = false, insertAfterRow = null) {
    const tbody = document.getElementById('tableBody');
    
    // Remove empty message if present
    const emptyRow = tbody.querySelector('.empty-row');
    if (emptyRow) {
        emptyRow.remove();
    }
    
    const tr = document.createElement('tr');
    if (isPCHeader) {
        tr.classList.add('pc-header-row');
    }
    // Add cell (for + button when row is clicked)
    const addCell = createAddCell(tr);
    tr.appendChild(addCell);
    // Create 10 editable cells
    for (let j = 0; j < 10; j++) {
        let defaultValue = '';
        if (isPCHeader && j === 0) {
            defaultValue = 'PC ' + (tbody.children.length + 1);
        }
        const td = createEditableCell(defaultValue, isPCHeader && j === 0, j, tr);
        tr.appendChild(td);
    }
    if (!isPCHeader) {
        const pictureCell = createPictureCell(null);
        tr.appendChild(pictureCell);
    }
    // Add action cell
    const actionCell = createActionCell();
    tr.appendChild(actionCell);
    
    if (insertAfterRow) {
        tbody.insertBefore(tr, insertAfterRow.nextElementSibling);
    } else {
        tbody.appendChild(tr);
    }
    if (!isPCHeader) mergeUnitColumnsInTable();
    setCurrentSheetData(getCurrentSheet().data, true);
    document.getElementById('exportBtn').disabled = false;
    document.getElementById('clearBtn').disabled = false;
    
    const firstInput = tr.querySelector('td input');
    if (firstInput) firstInput.focus();
}

// Modal functionality
const modal = document.getElementById('addItemModal');
const addItemForm = document.getElementById('addItemForm');
const closeModal = document.querySelector('.close-modal');
const cancelBtn = document.getElementById('cancelAddBtn');

// Article/Item list from INVENTORY-AS-OF-AUGUST-2026.xlsx (unique column values) - used when articles.json not loaded (e.g. file://)
const ARTICLE_ITEMS_FALLBACK = ["ACCESS POINT","AIRCON","AMPLIFIER","AUDIO","AVR","CEILING/ORBIT FAN","Celing fan","Computer set","Computer Table","Cubicle","Headset","KEVLER UM-200SDual Wireless Mic W/ Receiver","Keyboard","Monitor","Mouse","NETWORK","ORBITAL FAN","PRINTER","Router","SERVER","SMOKE DETECTOR","SOHO FD-4 Westminster 4-Drawer Lateral Filing Cabinet","SPARE","SPEAKER","SWITCH HUB","System Unit","UNSERVICEABLE SYSTEM UNIT","UPS","WEBCAM","Wireless Mic","WOOD AND STEEL TABLE"];

function toTitleCase(str) {
    if (!str || typeof str !== 'string') return str;
    return str.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
}

function loadArticleDropdown() {
    const articleSelect = document.getElementById('article');
    const articleOther = document.getElementById('articleOther');
    if (!articleSelect) return;
    function fillOptions(items) {
        if (!Array.isArray(items) || items.length === 0) items = ARTICLE_ITEMS_FALLBACK;
        const frag = document.createDocumentFragment();
        items.forEach(val => {
            const opt = document.createElement('option');
            opt.value = val;
            opt.textContent = toTitleCase(val);
            frag.appendChild(opt);
        });
        const otherOpt = document.createElement('option');
        otherOpt.value = '__OTHER__';
        otherOpt.textContent = 'Other (type below)';
        frag.appendChild(otherOpt);
        articleSelect.appendChild(frag);
        articleSelect.addEventListener('change', function() {
            if (articleOther) {
                articleOther.style.display = this.value === '__OTHER__' ? 'block' : 'none';
                articleOther.classList.toggle('form-field-blocked', this.value !== '__OTHER__');
                if (this.value !== '__OTHER__') articleOther.value = '';
            }
        });
    }
    fetch('articles.json')
        .then(res => res.ok ? res.json() : Promise.reject(new Error('Not found')))
        .then(items => { fillOptions(items); })
        .catch(() => { fillOptions(ARTICLE_ITEMS_FALLBACK); });
}

// Auto uppercase for Description in form (Article/It is now a dropdown)
document.addEventListener('DOMContentLoaded', function() {
    loadArticleDropdown();
    const descriptionInput = document.getElementById('description');
    
    if (descriptionInput) {
        descriptionInput.style.textTransform = 'uppercase';
        descriptionInput.addEventListener('input', function() {
            const cursorPos = this.selectionStart;
            this.value = this.value.toUpperCase();
            this.setSelectionRange(cursorPos, cursorPos);
        });
        descriptionInput.addEventListener('paste', function(e) {
            e.preventDefault();
            const pastedText = (e.clipboardData || window.clipboardData).getData('text').toUpperCase();
            const start = this.selectionStart;
            const end = this.selectionEnd;
            this.value = this.value.substring(0, start) + pastedText + this.value.substring(end);
            const newPos = start + pastedText.length;
            this.setSelectionRange(newPos, newPos);
        });
    }
});

// Open modal when Add Item button is clicked
document.getElementById('addItemBtn').addEventListener('click', function() {
    modal.classList.add('show');
    // Focus on first input
    document.getElementById('article').focus();
});

// Picture preview functionality
const pictureInput = document.getElementById('picture');
const picturePreview = document.getElementById('picturePreview');
let selectedImageData = null;

pictureInput.addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (file) {
        if (file.type.startsWith('image/')) {
            const reader = new FileReader();
            reader.onload = function(ev) {
                const dataUrl = ev.target.result;
                compressImage(dataUrl).then(function(compressed) {
                    selectedImageData = compressed;
                    picturePreview.innerHTML = '<img src="' + selectedImageData + '" alt="Preview">';
                    picturePreview.classList.remove('empty');
                });
            };
            reader.readAsDataURL(file);
        } else {
            alert('Please select an image file.');
            pictureInput.value = '';
            selectedImageData = null;
            picturePreview.innerHTML = '<span class="empty">No image selected</span>';
            picturePreview.classList.add('empty');
        }
    } else {
        selectedImageData = null;
        picturePreview.innerHTML = '<span class="empty">No image selected</span>';
        picturePreview.classList.add('empty');
    }
});

// Close modal
function closeModalFunc() {
    pendingRowInsert = null;
    modal.classList.remove('show');
    addItemForm.reset();
    const articleOther = document.getElementById('articleOther');
    if (articleOther) {
        articleOther.style.display = 'none';
        articleOther.value = '';
    }
    selectedImageData = null;
    picturePreview.innerHTML = '<span class="empty">No image selected</span>';
    picturePreview.classList.add('empty');
}

closeModal.addEventListener('click', closeModalFunc);
cancelBtn.addEventListener('click', closeModalFunc);

// Close modal when clicking outside
window.addEventListener('click', function(event) {
    if (event.target === modal) {
        closeModalFunc();
    }
});

// Prevent double submit when adding item (one item = one record)
let isAddingItem = false;

// Handle form submission
addItemForm.addEventListener('submit', function(e) {
    e.preventDefault();
    
    if (isAddingItem) {
        console.log('⏸️ Add item already in progress, ignoring duplicate submit');
        return;
    }
    
    // Get Article/It (dropdown or "Other" text) - allow blank
    const articleEl = document.getElementById('article');
    const articleOtherEl = document.getElementById('articleOther');
    let article = (articleEl && articleEl.value) ? articleEl.value.trim() : '';
    if (article === '__OTHER__' && articleOtherEl) {
        article = articleOtherEl.value.trim();
    }
    const description = document.getElementById('description').value.trim();
    const condition = document.getElementById('condition').value.trim();
    
    // Get form values (saving allowed even with blanks); picture from form upload, nagsesave sa row
    const formData = {
        article: toTitleCase(article),
        description: description.toUpperCase(),
        oldProperty: document.getElementById('oldProperty').value.trim(),
        unitOfMeas: document.getElementById('unitOfMeas').value.trim(),
        unitValue: document.getElementById('unitValue').value.trim(),
        quantity: document.getElementById('quantity').value.trim(),
        location: document.getElementById('location').value.trim(),
        condition: condition,
        remarks: document.getElementById('remarks').value.trim(),
        user: document.getElementById('user').value.trim(),
        picture: selectedImageData || null
    };
    
    // Add item to table (only once)
    try {
        isAddingItem = true;
        const insertRef = pendingRowInsert ? pendingRowInsert.refRow : null;
        const position = pendingRowInsert ? pendingRowInsert.position : null;
        addItemFromForm(formData, insertRef, position);
        pendingRowInsert = null;
        
        // Close modal and reset form
        closeModalFunc();
    } catch (error) {
        console.error('Error adding item:', error);
        alert('Error adding item. Please try again.');
    } finally {
        // Re-enable form after a short delay so one click = one add
        setTimeout(function() {
            isAddingItem = false;
        }, 800);
    }
});

// Convert data URL to Blob for Storage upload
function dataURLtoBlob(dataUrl) {
    const arr = dataUrl.split(',');
    const mime = (arr[0].match(/:(.*?);/) || [])[1] || 'image/jpeg';
    const bstr = atob(arr[1]);
    let n = bstr.length;
    const u8arr = new Uint8Array(n);
    while (n--) u8arr[n] = bstr.charCodeAt(n);
    return new Blob([u8arr], { type: mime });
}

// Upload data URL to Supabase Storage; return public URL so PC Location link can show picture
var _storageBucketChecked = false;
var _storageBucketMissing = false;
async function uploadDataUrlToStorage(dataUrl, sheetId, rowIndex) {
    if (!window.supabaseClient || !dataUrl || !dataUrl.startsWith('data:image/')) return null;
    if (_storageBucketMissing) return null;
    try {
        const blob = dataURLtoBlob(dataUrl);
        const path = `${(sheetId || 'sheet-1')}/${rowIndex}/${Date.now()}.jpg`;
        const { error } = await window.supabaseClient.storage.from('inventory-pictures').upload(path, blob, { contentType: 'image/jpeg', upsert: true });
        if (error) {
            _storageBucketChecked = true;
            const msg = error.message || '';
            if (msg.includes('Bucket not found') || msg.includes('not found')) {
                _storageBucketMissing = true;
                console.warn('Storage bucket "inventory-pictures" not found. Create it in Supabase: Dashboard → Storage → New bucket → name: inventory-pictures, Public: ON.');
            } else if (msg.includes('row-level security') || msg.includes('violates')) {
                _storageBucketMissing = true;
                console.warn('Storage RLS: Run database/storage-policies.sql in Supabase SQL Editor to allow uploads to inventory-pictures. See database/README.md.');
            } else {
                console.warn('Storage upload failed:', msg);
            }
            return null;
        }
        const { data: { publicUrl } } = window.supabaseClient.storage.from('inventory-pictures').getPublicUrl(path);
        return publicUrl;
    } catch (e) {
        if (!_storageBucketChecked) console.warn('Storage upload error:', e);
        return null;
    }
}

// Compress image so it saves to Supabase and appears in Excel (max 800px, JPEG 0.8)
function compressImage(dataUrl) {
    return new Promise(function(resolve) {
        const img = new Image();
        img.onload = function() {
            const maxSize = 800;
            let w = img.width, h = img.height;
            if (w > maxSize || h > maxSize) {
                if (w > h) { h = (h * maxSize) / w; w = maxSize; }
                else { w = (w * maxSize) / h; h = maxSize; }
            }
            const canvas = document.createElement('canvas');
            canvas.width = w;
            canvas.height = h;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0, w, h);
            try {
                const compressed = canvas.toDataURL('image/jpeg', 0.8);
                resolve(compressed);
            } catch (e) {
                resolve(dataUrl); // fallback to original
            }
        };
        img.onerror = function() { resolve(dataUrl); };
        img.src = dataUrl;
    });
}

// Create picture cell
function createPictureCell(imageData = null) {
    const td = document.createElement('td');
    td.classList.add('picture-cell');
    
    if (imageData) {
        const img = document.createElement('img');
        img.src = imageData;
        img.alt = 'Item Picture';
        img.addEventListener('click', function() {
            showImageModal(img.src, td);
        });
        td.appendChild(img);
        
        // Add change picture button
        const changeBtn = document.createElement('button');
        changeBtn.className = 'upload-picture-btn';
        changeBtn.textContent = '📷 Change';
        changeBtn.addEventListener('click', function(e) {
            e.stopPropagation();
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = 'image/*';
            input.addEventListener('change', function(e) {
                const file = e.target.files[0];
                if (file) {
                    const reader = new FileReader();
                    reader.onload = function(e) {
                        compressImage(e.target.result).then(function(compressed) {
                            img.src = compressed;
                            updateDataFromTable();
                        });
                    };
                    reader.readAsDataURL(file);
                }
            });
            input.click();
        });
        td.appendChild(changeBtn);
    } else {
        const noImage = document.createElement('span');
        noImage.className = 'no-image';
        noImage.textContent = 'No image';
        td.appendChild(noImage);
        
        const uploadBtn = document.createElement('button');
        uploadBtn.className = 'upload-picture-btn';
        uploadBtn.textContent = '📷 Upload';
        uploadBtn.addEventListener('click', function() {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = 'image/*';
            input.addEventListener('change', function(e) {
                const file = e.target.files[0];
                if (file) {
                    const reader = new FileReader();
                    reader.onload = function(e) {
                        compressImage(e.target.result).then(function(compressed) {
                            td.innerHTML = '';
                            const img = document.createElement('img');
                            img.src = compressed;
                            img.alt = 'Item Picture';
                            img.addEventListener('click', function() {
                                showImageModal(img.src, td);
                            });
                            td.appendChild(img);
                            
                            const changeBtn = document.createElement('button');
                            changeBtn.className = 'upload-picture-btn';
                            changeBtn.textContent = '📷 Change';
                            changeBtn.addEventListener('click', function(e) {
                                e.stopPropagation();
                                const changeInput = document.createElement('input');
                                changeInput.type = 'file';
                                changeInput.accept = 'image/*';
                                changeInput.addEventListener('change', function(e) {
                                    const file = e.target.files[0];
                                    if (file) {
                                        const reader2 = new FileReader();
                                        reader2.onload = function(e) {
                                            compressImage(e.target.result).then(function(c) {
                                                img.src = c;
                                                updateDataFromTable();
                                            });
                                        };
                                        reader2.readAsDataURL(file);
                                    }
                                });
                                changeInput.click();
                            });
                            td.appendChild(changeBtn);
                            updateDataFromTable();
                        });
                    };
                    reader.readAsDataURL(file);
                }
            });
            input.click();
        });
        td.appendChild(uploadBtn);
    }
    
    return td;
}

// Reset picture cell to "No image" + Upload button (same as empty state in createPictureCell)
function setPictureCellEmpty(td) {
    if (!td || !td.classList.contains('picture-cell')) return;
    td.innerHTML = '';
    const noImage = document.createElement('span');
    noImage.className = 'no-image';
    noImage.textContent = 'No image';
    td.appendChild(noImage);
    const uploadBtn = document.createElement('button');
    uploadBtn.className = 'upload-picture-btn';
    uploadBtn.textContent = '📷 Upload';
    uploadBtn.addEventListener('click', function() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'image/*';
        input.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = function(e) {
                    compressImage(e.target.result).then(function(compressed) {
                        td.innerHTML = '';
                        const img = document.createElement('img');
                        img.src = compressed;
                        img.alt = 'Item Picture';
                        img.addEventListener('click', function() {
                            showImageModal(img.src, td);
                        });
                        td.appendChild(img);
                        const changeBtn = document.createElement('button');
                        changeBtn.className = 'upload-picture-btn';
                        changeBtn.textContent = '📷 Change';
                        changeBtn.addEventListener('click', function(e) {
                            e.stopPropagation();
                            const changeInput = document.createElement('input');
                            changeInput.type = 'file';
                            changeInput.accept = 'image/*';
                            changeInput.addEventListener('change', function(e) {
                                const file = e.target.files[0];
                                if (file) {
                                    const reader2 = new FileReader();
                                    reader2.onload = function(e) {
                                        compressImage(e.target.result).then(function(c) {
                                            img.src = c;
                                            updateDataFromTable();
                                        });
                                    };
                                    reader2.readAsDataURL(file);
                                }
                            });
                            changeInput.click();
                        });
                        td.appendChild(changeBtn);
                        updateDataFromTable();
                    });
                };
                reader.readAsDataURL(file);
            }
        });
        input.click();
    });
    td.appendChild(uploadBtn);
}

// Show image in modal; optional pictureCell = td.picture-cell — if provided, shows "Delete picture" button
function showImageModal(imageSrc, pictureCell) {
    let imageModal = document.getElementById('imageModal');
    if (!imageModal) {
        imageModal = document.createElement('div');
        imageModal.id = 'imageModal';
        imageModal.className = 'image-modal';
        const inner = document.createElement('div');
        inner.className = 'image-modal-inner';
        inner.addEventListener('click', function(e) { e.stopPropagation(); });
        const img = document.createElement('img');
        inner.appendChild(img);
        const deletePicBtn = document.createElement('button');
        deletePicBtn.type = 'button';
        deletePicBtn.className = 'image-modal-delete-btn';
        deletePicBtn.textContent = '🗑️ Delete picture';
        inner.appendChild(deletePicBtn);
        imageModal.appendChild(inner);
        document.body.appendChild(imageModal);
        imageModal.addEventListener('click', function() {
            imageModal.classList.remove('show');
        });
    }
    const imgEl = imageModal.querySelector('img');
    const deleteBtn = imageModal.querySelector('.image-modal-delete-btn');
    imgEl.src = imageSrc;
    if (pictureCell) {
        deleteBtn.style.display = 'inline-block';
        deleteBtn.onclick = function(e) {
            e.stopPropagation();
            setPictureCellEmpty(pictureCell);
            updateDataFromTable();
            imageModal.classList.remove('show');
        };
    } else {
        deleteBtn.style.display = 'none';
    }
    imageModal.classList.add('show');
}

// Add item from form data (optional insertRef + position for row-add dropdown)
function addItemFromForm(formData, insertRef, position) {
    const tbody = document.getElementById('tableBody');
    
    // Remove empty message if present
    const emptyRow = tbody.querySelector('.empty-row');
    if (emptyRow) {
        emptyRow.remove();
    }
    
    const tr = document.createElement('tr');
    const addCell = createAddCell(tr);
    tr.appendChild(addCell);
    
    // Create cells with form data
    const cells = [
        formData.article,
        formData.description,
        formData.oldProperty,
        formData.unitOfMeas,
        formData.unitValue,
        formData.quantity,
        formData.location,
        formData.condition,
        formData.remarks,
        formData.user
    ];
    
    cells.forEach((value, index) => {
        const td = createEditableCell(value, false, index, tr);
        tr.appendChild(td);
    });
    
    // Apply condition-based color
    applyConditionColor(tr, formData.condition);
    
    // Add picture cell
    const pictureCell = createPictureCell(formData.picture);
    tr.appendChild(pictureCell);
    
    // Add action cell
    const actionCell = createActionCell();
    tr.appendChild(actionCell);
    
    if (insertRef && (position === 'above' || position === 'below')) {
        if (position === 'above') {
            tbody.insertBefore(tr, insertRef);
        } else {
            tbody.insertBefore(tr, insertRef.nextElementSibling);
        }
    } else {
        tbody.appendChild(tr);
    }
    mergeUnitColumnsInTable();
    document.getElementById('exportBtn').disabled = false;
    document.getElementById('clearBtn').disabled = false;
    
    updateDataFromTable();
}

// Add PC Section button
document.getElementById('addPCBtn').addEventListener('click', function() {
    const pcName = prompt('Enter PC Section Name (e.g., "PC USED BY: JOVEN T. CRUZ" or "SERVER"):');
    if (pcName !== null) {
        const tbody = document.getElementById('tableBody');
        const emptyRow = tbody.querySelector('.empty-row');
        if (emptyRow) {
            emptyRow.remove();
        }
        
        const tr = document.createElement('tr');
        tr.classList.add('pc-header-row');
        const addCell = createAddCell(tr);
        tr.appendChild(addCell);
        // Isang merged cell mula Article/It hanggang Picture (11 columns), kulay gray
        const td = document.createElement('td');
        td.colSpan = 11;
        td.className = 'pc-name-cell';
        td.style.fontWeight = 'bold';
        td.style.fontSize = '14px';
        
        const wrapper = document.createElement('div');
        wrapper.style.display = 'flex';
        wrapper.style.alignItems = 'center';
        wrapper.style.justifyContent = 'space-between';
        wrapper.style.gap = '10px';
        wrapper.style.width = '100%';
        
        const input = document.createElement('input');
        input.type = 'text';
        input.value = pcName;
        input.placeholder = 'PC name (e.g. PC 1, PC USED BY: NAME)';
        input.style.flex = '1';
        input.style.textAlign = 'center';
        input.style.fontWeight = 'bold';
        input.style.fontSize = '14px';
        input.style.border = 'none';
        input.style.background = 'transparent';
        input.addEventListener('blur', function() {
            updateDataFromTable();
        });
        input.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') this.blur();
        });
        
        const deleteSectionBtn = document.createElement('button');
        deleteSectionBtn.type = 'button';
        deleteSectionBtn.textContent = '🗑️ Delete section';
        deleteSectionBtn.className = 'delete-btn';
        deleteSectionBtn.style.flexShrink = '0';
        deleteSectionBtn.addEventListener('click', function() {
            if (confirm('Delete this PC section and all items under it?')) {
                tr.remove();
                mergeUnitColumnsInTable();
                updateDataFromTable();
                const remaining = tbody.querySelectorAll('tr:not(.empty-row)').length;
                if (remaining === 0) {
                    tbody.innerHTML = '<tr class="empty-row"><td colspan="13" class="empty-message">No data loaded. Please import an Excel file or add items manually.</td></tr>';
                    setCurrentSheetData([], false);
                    document.getElementById('clearBtn').disabled = true;
                }
            }
        });
        
        wrapper.appendChild(input);
        wrapper.appendChild(deleteSectionBtn);
        td.appendChild(wrapper);
        tr.appendChild(td);
        tr.appendChild(document.createElement('td')); // Actions column (blank)
        
        tbody.appendChild(tr);
        mergeUnitColumnsInTable();
        updateDataFromTable(); // Save and sync to Supabase
        document.getElementById('exportBtn').disabled = false;
        document.getElementById('clearBtn').disabled = false;
    }
});

// Display data in table (readOnlyForBackup = true kapag backup page para tamang empty message)
function displayData(data, readOnlyForBackup) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    
    console.log(`🖥️ Displaying data: ${data ? data.length : 0} row(s)`);
    
    if (!data || data.length === 0) {
        var msg = readOnlyForBackup
            ? 'No backup data yet. I-open muna ang main Inventory page (walang ?backup=1), hintayin mag-load ang data, tapos balik dito.'
            : 'No data found in the Excel file.';
        tbody.innerHTML = '<tr class="empty-row"><td colspan="13" class="empty-message">' + msg + '</td></tr>';
        console.log('⚠️ No data to display');
        renderSummaryTable();
        return;
    }
    
    // Find header row
    let headerRowIndex = 0;
    for (let i = 0; i < Math.min(20, data.length); i++) {
        if (data[i] && data[i].length > 0) {
            const rowText = data[i].map(cell => cell ? cell.toString().toLowerCase() : '').join(' ');
            if (rowText.includes('article') && rowText.includes('description')) {
                headerRowIndex = i;
                break;
            }
        }
    }
    
    // Process all rows
    let dataRowIndex = 0; // Track index for highlight states (excluding PC headers)
    const sheet = getCurrentSheet();
    
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) {
            continue;
        }
        // Check if this is a header row
        const rowText = row.map(cell => cell ? cell.toString().toLowerCase() : '').join(' ');
        const isHeaderRow = rowText.includes('article') && rowText.includes('description');
        
        if (isHeaderRow) {
            continue; // Skip header row as it's already in thead
        }
        // Check if this is a PC header row (first meaningful cell may be in row[2] for exported PC row)
        let firstCell = (row.find(c => c != null && String(c).trim() !== '') || row[0] || '').toString().trim();
        const isPCHeader = firstCell && (
            row.length === 1 || // Single cell = PC header
            firstCell.toUpperCase().includes('PC USED BY') ||
            firstCell.toUpperCase() === 'SERVER' ||
            firstCell.toUpperCase().startsWith('PC ') ||
            firstCell.toUpperCase() === 'PC'
        );
        // Kung na-save dati ang text ng dropdown (+AboveBelowAdd Item...) gamitin default at gawing editable
        if (isPCHeader && /above|below|add\s*item|add\s*pc\s*section/i.test(firstCell)) {
            firstCell = 'PC Section';
        }
        
        const tr = document.createElement('tr');
        if (isPCHeader) {
            tr.classList.add('pc-header-row');
            const addCell = createAddCell(tr);
            tr.appendChild(addCell);
            // Isang merged cell mula Article/It hanggang Picture (11 columns), kulay gray — may class para sigurado sa save
            const td = document.createElement('td');
            td.colSpan = 11;
            td.className = 'pc-name-cell';
            td.style.fontWeight = 'bold';
            td.style.fontSize = '14px';
            
            const wrapper = document.createElement('div');
            wrapper.style.display = 'flex';
            wrapper.style.alignItems = 'center';
            wrapper.style.justifyContent = 'space-between';
            wrapper.style.gap = '10px';
            wrapper.style.width = '100%';
            
            const input = document.createElement('input');
            input.type = 'text';
            input.value = firstCell;
            input.placeholder = 'PC name (e.g. PC 1, PC USED BY: NAME)';
            input.style.flex = '1';
            input.style.textAlign = 'center';
            input.style.fontWeight = 'bold';
            input.style.fontSize = '14px';
            input.style.border = 'none';
            input.style.background = 'transparent';
            input.addEventListener('blur', function() {
                updateDataFromTable();
            });
            input.addEventListener('keypress', function(e) {
                if (e.key === 'Enter') this.blur();
            });
            
            const deleteSectionBtn = document.createElement('button');
            deleteSectionBtn.type = 'button';
            deleteSectionBtn.textContent = '🗑️ Delete section';
            deleteSectionBtn.className = 'delete-btn';
            deleteSectionBtn.style.flexShrink = '0';
            deleteSectionBtn.addEventListener('click', function() {
                const tbody = document.getElementById('tableBody');
                if (confirm('Delete this PC section and all items under it?')) {
                    tr.remove();
                    mergeUnitColumnsInTable();
                    updateDataFromTable();
                    if (tbody.querySelectorAll('tr:not(.empty-row)').length === 0) {
                        tbody.innerHTML = '<tr class="empty-row"><td colspan="13" class="empty-message">No data loaded. Please import an Excel file or add items manually.</td></tr>';
                        setCurrentSheetData([], false);
                        document.getElementById('clearBtn').disabled = true;
                    }
                }
            });
            
            wrapper.appendChild(input);
            wrapper.appendChild(deleteSectionBtn);
            td.appendChild(wrapper);
            tr.appendChild(td);
            var actionsTd = document.createElement('td');
            actionsTd.classList.add('actions-cell');
            tr.appendChild(actionsTd);
        } else {
            // Legacy: kung na-save dati with # column (11 elements, first is number), strip it
            let dataRow = row;
            if (Array.isArray(row) && row.length === 11 && /^\d+$/.test(String(row[0] || '').trim())) {
                dataRow = row.slice(1);
            }
            const addCell = createAddCell(tr);
            tr.appendChild(addCell);
            // Create 10 editable cells for regular rows
            for (let j = 0; j < 10; j++) {
                const cellValue = dataRow[j] !== undefined && dataRow[j] !== null ? dataRow[j].toString() : '';
                const td = createEditableCell(cellValue, false, j, tr);
                tr.appendChild(td);
            }
            
            // Apply condition-based color (condition is at index 7)
            const conditionValue = dataRow[7] ? dataRow[7].toString().trim() : '';
            applyConditionColor(tr, conditionValue);
            
            // Get picture URL from sheet metadata (pictureUrls is aligned with sheet.data by row index i)
            let pictureData = null;
            if (sheet.pictureUrls && sheet.pictureUrls[i]) {
                pictureData = sheet.pictureUrls[i];
            }
            
            // Add picture cell
            const pictureCell = createPictureCell(pictureData);
            tr.appendChild(pictureCell);
            
            // Add action cell
            const actionCell = createActionCell();
            tr.appendChild(actionCell);
            
            // Restore highlight state if exists
            if (sheet.highlightStates && sheet.highlightStates[dataRowIndex] === true) {
                tr.classList.add('highlighted-row');
                const highlightBtn = tr.querySelector('.highlight-btn');
                if (highlightBtn) {
                    highlightBtn.textContent = '✅ Highlighted';
                    highlightBtn.style.background = '#28a745';
                }
            }
            
            dataRowIndex++;
        }
        
        tbody.appendChild(tr);
    }
    // Run merge after DOM is fully updated so Picture (and User, Unit of meas, Unit Value) merge correctly
    setTimeout(function() { mergeUnitColumnsInTable(); }, 0);
}

// 3-dot menu on sheet tab: Rename / Delete
function setupSheetTabMenu(tab, sheetId) {
    const menuBtn = tab.querySelector('.sheet-tab-menu');
    const dropdown = tab.querySelector('.sheet-tab-dropdown');
    const nameSpan = tab.querySelector('.sheet-name');
    if (!menuBtn || !dropdown) return;
    
    menuBtn.addEventListener('click', function(e) {
        e.preventDefault();
        e.stopPropagation();
        document.querySelectorAll('.sheet-tab-dropdown.show').forEach(d => {
            d.classList.remove('show');
            d.removeAttribute('style');
        });
        const isOpening = !dropdown.classList.contains('show');
        dropdown.classList.toggle('show');
        if (isOpening) {
            const rect = menuBtn.getBoundingClientRect();
            dropdown.style.top = (rect.bottom + 4) + 'px';
            dropdown.style.left = Math.max(8, rect.right - 120) + 'px';
        } else {
            dropdown.removeAttribute('style');
        }
    });
    
    dropdown.querySelector('.rename-option')?.addEventListener('click', async function(e) {
        e.stopPropagation();
        dropdown.classList.remove('show');
        dropdown.removeAttribute('style');
        const sheet = sheets[sheetId];
        if (!sheet) return;
        const newName = prompt('Rename sheet:', sheet.name);
        if (newName !== null && newName.trim() !== '') {
            sheet.name = newName.trim();
            if (nameSpan) nameSpan.textContent = sheet.name;
            if (checkSupabaseConnection()) {
                try {
                    await window.supabaseClient.from('sheets').update({ name: sheet.name }).eq('id', sheetId);
                } catch (err) {
                    console.warn('Supabase sheet rename sync failed:', err);
                }
            }
        }
    });
    
    dropdown.querySelector('.delete-option')?.addEventListener('click', function(e) {
        e.stopPropagation();
        dropdown.classList.remove('show');
        dropdown.removeAttribute('style');
        deleteSheet(sheetId);
    });
}

document.addEventListener('click', function() {
        document.querySelectorAll('.sheet-tab-dropdown.show').forEach(d => {
            d.classList.remove('show');
            d.removeAttribute('style');
        });
        document.querySelectorAll('.row-add-dropdown.show').forEach(d => {
            d.classList.remove('show');
            const opts = d.querySelectorAll('.row-add-option');
            opts.forEach((o, i) => { o.style.display = i < 2 ? 'block' : 'none'; });
        });
    });

// Sheet Management Functions
function createNewSheet(name = null, data = null) {
    sheetCounter++;
    const sheetId = `sheet-${sheetCounter}`;
    const sheetName = name || `Sheet ${sheetCounter}`;
    
    sheets[sheetId] = {
        id: sheetId,
        name: sheetName,
        data: data || [],
        hasData: data ? true : false,
        highlightStates: [], // Initialize highlight states
        pictureUrls: [] // Initialize picture URLs
    };
    
    // Create tab
    const sheetTabs = document.getElementById('sheetTabs');
    const tab = document.createElement('div');
    tab.className = 'sheet-tab';
    tab.setAttribute('data-sheet-id', sheetId);
    tab.innerHTML = `
        <span class="sheet-name">${sheetName}</span>
        <span class="sheet-tab-menu" data-sheet-id="${sheetId}" title="More">⋮</span>
        <div class="sheet-tab-dropdown">
            <button type="button" class="rename-option">Rename</button>
            <button type="button" class="delete-option">Delete</button>
        </div>
        <span class="close-sheet" data-sheet-id="${sheetId}">×</span>
    `;
    
    tab.addEventListener('click', function(e) {
        if (!e.target.closest('.close-sheet') && !e.target.closest('.sheet-tab-menu') && !e.target.closest('.sheet-tab-dropdown')) {
            switchToSheet(sheetId);
        }
    });
    
    const closeBtn = tab.querySelector('.close-sheet');
    closeBtn.addEventListener('click', function(e) {
        e.stopPropagation();
        deleteSheet(sheetId);
    });
    
    setupSheetTabMenu(tab, sheetId);
    sheetTabs.appendChild(tab);
    switchToSheet(sheetId);
    
    return sheetId;
}

function switchToSheet(sheetId) {
    if (!sheets[sheetId]) return;
    
    // Save current sheet data from table only when NOT loading from Supabase (and not in backup read-only mode)
    if (!isLoadingFromSupabase && !isBackupMode) {
        saveCurrentSheetData();
    }
    
    // Switch to new sheet
    currentSheetId = sheetId;
    
    // Update tabs
    document.querySelectorAll('.sheet-tab').forEach(tab => {
        tab.classList.remove('active');
    });
    const activeTab = document.querySelector(`[data-sheet-id="${sheetId}"]`);
    if (activeTab) activeTab.classList.add('active');
    
    // Load sheet data
    const sheet = sheets[sheetId];
    if (!sheet) {
        console.error(`❌ Sheet ${sheetId} not found in sheets object`);
        return;
    }
    
    console.log(`📋 switchToSheet: Sheet "${sheet.name}" has ${sheet.data ? sheet.data.length : 0} row(s) in memory`);
    console.log(`📋 switchToSheet: Sheet data:`, sheet.data);
    
    displayData(sheet.data);
    if (isBackupMode) makeTableReadOnly();
    
    // Update buttons
    document.getElementById('exportBtn').disabled = !hasAnyData();
    if (!isBackupMode) document.getElementById('clearBtn').disabled = !sheet.hasData;
    renderSummaryTable();
}

function saveCurrentSheetData(skipSync) {
    const tbody = document.getElementById('tableBody');
    const rows = tbody.querySelectorAll('tr:not(.empty-row)');
    
    const sheetData = [];
    const highlightStates = [];
    const pictureUrls = []; // Store picture URLs (per row; merged section uses first row's picture)
    let sectionUnitMeas = '';
    let sectionUnitValue = '';
    let sectionUser = '';
    let sectionPictureUrl = null;

    rows.forEach((row, index) => {
        if (row.classList.contains('pc-header-row')) {
            sectionUnitMeas = '';
            sectionUnitValue = '';
            sectionUser = '';
            sectionPictureUrl = null;
            const firstCell = row.querySelector('td.pc-name-cell') || row.querySelector('td[colspan="11"]');
            if (firstCell) {
                const input = firstCell.querySelector('input');
                const pcName = input ? input.value.trim() : firstCell.textContent.trim();
                sheetData.push([pcName]);
                highlightStates.push(false);
                pictureUrls.push(null);
            }
        } else {
            const cells = Array.from(row.querySelectorAll('td.editable'));
            cells.sort((a, b) => (parseInt(a.getAttribute('data-column') || '0', 10) - parseInt(b.getAttribute('data-column') || '0', 10)));
            const rowData = [];
            cells.forEach(cell => {
                const input = cell.querySelector('input');
                rowData.push(input ? input.value : '');
            });
            if (rowData.length === 7) {
                rowData.splice(3, 0, sectionUnitMeas, sectionUnitValue);
                rowData.splice(9, 0, sectionUser);
            } else if (rowData.length === 8) {
                rowData.splice(3, 0, sectionUnitMeas, sectionUnitValue);
            } else if (rowData.length >= 10) {
                sectionUnitMeas = rowData[UNIT_MEAS_COL] || '';
                sectionUnitValue = rowData[UNIT_VALUE_COL] || '';
                sectionUser = rowData[USER_COL] || '';
            }
            if (rowData.length > 0) {
                sheetData.push(rowData);
                highlightStates.push(row.classList.contains('highlighted-row'));
                const pictureCell = row.querySelector('.picture-cell');
                sectionPictureUrl = null;
                if (pictureCell) {
                    const img = pictureCell.querySelector('img');
                    if (img && img.src) sectionPictureUrl = img.src;
                }
                pictureUrls.push(sectionPictureUrl);
            }
        }
    });
    
    setCurrentSheetData(sheetData, sheetData.length > 0);
    // Store highlight states and picture URLs in sheet metadata
    sheets[currentSheetId].highlightStates = highlightStates;
    sheets[currentSheetId].pictureUrls = pictureUrls;
    
    renderSummaryTable();
    // Laging i-save sa localStorage agad — kahit mag-refresh hindi mawawala
    saveBackupToLocalStorage();
    
    // Don't sync if we're loading from Supabase or caller asked to skip (e.g. when preserving table during load)
    if (skipSync || isLoadingFromSupabase) {
        if (isLoadingFromSupabase) console.log('⏸️ Skipping sync during load from Supabase');
        return;
    }
    // Debounce sync — mas mabilis (100ms) para hindi mawala sa refresh
    if (checkSupabaseConnection()) {
        if (syncTimeout) clearTimeout(syncTimeout);
        syncTimeout = setTimeout(() => {
            syncTimeout = null;
            syncToSupabase();
            if (hasAnyData()) syncBackupToSupabase();
        }, 100);
    }
}

// Retry async op on "Failed to fetch" / network errors (common on Vercel)
// Supabase returns { data, error } and does NOT throw on fetch failure — so we must throw when error is fetch-related to trigger retry
function isFetchError(err) {
    if (!err) return false;
    const msg = (err.message || '').toString();
    return msg === 'Failed to fetch' || msg.includes('fetch') || err.name === 'TypeError';
}
async function withRetry(fn, maxAttempts = 4, delayMs = 1500) {
    let lastErr;
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        try {
            const result = await fn();
            // Supabase client resolves with { error } on network failure; throw so we retry
            if (result && result.error && isFetchError(result.error)) {
                throw result.error;
            }
            return result;
        } catch (e) {
            lastErr = e;
            if (attempt < maxAttempts && isFetchError(e)) {
                console.warn(`⚠️ Attempt ${attempt} failed (${e.message}), retrying in ${delayMs}ms...`);
                await new Promise(r => setTimeout(r, delayMs));
            } else {
                throw e;
            }
        }
    }
    throw lastErr;
}

// Supabase Sync Functions
async function syncToSupabase() {
    if (!checkSupabaseConnection()) {
        console.warn('⚠️ Cannot sync: Supabase not connected');
        return;
    }
    
    // Prevent concurrent syncs
    if (isSyncing) {
        console.log('⏸️ Sync already in progress, skipping duplicate call');
        return;
    }
    
    isSyncing = true;
    const statusEl = document.getElementById('saveStatus');
    if (statusEl) {
        statusEl.textContent = 'Saving…';
        statusEl.className = 'save-status saving';
    }
    
    try {
        const sheet = getCurrentSheet();
        console.log(`📤 Starting sync for sheet: ${sheet.name} (${currentSheetId})`);
        
        // Save/update sheet in Supabase (with retry for Failed to fetch)
        const sheetResult = await withRetry(async () => {
            return await window.supabaseClient
                .from('sheets')
                .upsert({
                    id: currentSheetId,
                    name: sheet.name,
                    updated_at: new Date().toISOString()
                }, { onConflict: 'id' });
        });
        
        if (sheetResult.error) {
            const sheetError = sheetResult.error;
            console.error('❌ Error saving sheet to Supabase:', sheetError);
            if (statusEl) { statusEl.textContent = 'Error saving'; statusEl.className = 'save-status error'; }
            alert(`Error saving to database: ${sheetError.message}`);
            return;
        }
        
        console.log(`✅ Sheet "${sheet.name}" saved to Supabase`);
        
        // Delete existing items for this sheet (with retry)
        const deleteResult = await withRetry(async () => {
            return await window.supabaseClient
                .from('inventory_items')
                .delete()
                .eq('sheet_id', currentSheetId);
        });
        
        if (deleteResult.error) {
            const deleteError = deleteResult.error;
            console.error('❌ Error deleting old items:', deleteError);
            if (statusEl) { statusEl.textContent = 'Error saving'; statusEl.className = 'save-status error'; }
            alert(`Error updating database: ${deleteError.message}`);
            return;
        }
        
        console.log('🗑️ Old items deleted, preparing new items...');
        
        // Prepare items to insert — use same logic as saveCurrentSheetData so merged rows (7–8 cells) are included
        const itemsToInsert = [];
        const tbody = document.getElementById('tableBody');
        const rows = tbody.querySelectorAll('tr:not(.empty-row)');
        let sectionUnitMeas = '';
        let sectionUnitValue = '';
        let sectionUser = '';
        let sectionPictureUrl = null;
        
        console.log(`📋 Found ${rows.length} row(s) in table`);
        
        let lastPCHeaderName = null;
        let lastDataRowKey = null; // para i-skip ang magkasunod na duplicate data row
        rows.forEach((row, rowIndex) => {
            if (row.classList.contains('pc-header-row')) {
                lastDataRowKey = null;
                sectionUnitMeas = '';
                sectionUnitValue = '';
                sectionUser = '';
                sectionPictureUrl = null;
                const firstCell = row.querySelector('td.pc-name-cell') || row.querySelector('td[colspan="11"]');
                if (firstCell) {
                    const input = firstCell.querySelector('input');
                    const pcName = (input ? input.value.trim() : firstCell.textContent.trim()) || '';
                    // Huwag mag-save ng magkasunod na duplicate PC header (e.g. dalawang "PC 1")
                    if (pcName && pcName === lastPCHeaderName) return;
                    lastPCHeaderName = pcName;
                    itemsToInsert.push({
                        sheet_id: currentSheetId,
                        sheet_name: sheet.name,
                        row_index: rowIndex,
                        article: pcName,
                        is_pc_header: true,
                        is_highlighted: false
                    });
                }
            } else {
                const cells = Array.from(row.querySelectorAll('td.editable'));
                cells.sort((a, b) => (parseInt(a.getAttribute('data-column') || '0', 10) - parseInt(b.getAttribute('data-column') || '0', 10)));
                const vals = cells.map(c => (c.querySelector('input')?.value || '').trim());
                // Rebuild 10 columns when row has merged cells (7 or 8 editable cells)
                let rowData = vals;
                if (vals.length === 7) {
                    rowData = [
                        vals[0], vals[1], vals[2],
                        sectionUnitMeas, sectionUnitValue,
                        vals[3], vals[4], vals[5], vals[6],
                        sectionUser
                    ];
                } else if (vals.length === 8) {
                    rowData = [
                        vals[0], vals[1], vals[2],
                        sectionUnitMeas, sectionUnitValue,
                        vals[3], vals[4], vals[5], vals[6], vals[7]
                    ];
                } else if (vals.length >= 10) {
                    sectionUnitMeas = vals[UNIT_MEAS_COL] || '';
                    sectionUnitValue = vals[UNIT_VALUE_COL] || '';
                    sectionUser = vals[USER_COL] || '';
                }
                // Huwag mag-save ng magkasunod na duplicate data row (same 10 columns)
                const dataKey = (rowData.slice(0, 10) || []).map(v => (v || '').toString().trim()).join('\t');
                if (dataKey && dataKey === lastDataRowKey) return;
                lastDataRowKey = dataKey;
                lastPCHeaderName = null;
                sectionPictureUrl = null;
                const pictureCell = row.querySelector('.picture-cell');
                if (pictureCell) {
                    const img = pictureCell.querySelector('img');
                    if (img && img.src) {
                        const src = String(img.src);
                        if (!src.startsWith('data:')) {
                            sectionPictureUrl = src;
                        } else if (src.length <= 800000) {
                            sectionPictureUrl = src;
                        }
                    }
                }
                itemsToInsert.push({
                    sheet_id: currentSheetId,
                    sheet_name: sheet.name,
                    row_index: rowIndex,
                    article: rowData[0] || '',
                    description: rowData[1] || '',
                    old_property_n_assigned: rowData[2] || '',
                    unit_of_meas: rowData[3] || '',
                    unit_value: rowData[4] || '',
                    quantity: rowData[5] || '',
                    location: rowData[6] || '',
                    condition: rowData[7] || '',
                    remarks: rowData[8] || '',
                    user: rowData[9] || '',
                    picture_url: sectionPictureUrl,
                    is_pc_header: false,
                    is_highlighted: row.classList.contains('highlighted-row')
                });
            }
        });
        
        // Upload data URL pictures to Storage so we get https URL — para lumabas ang picture sa PC Location link
        const rowIndexToPublicUrl = {};
        for (let i = 0; i < itemsToInsert.length; i++) {
            const item = itemsToInsert[i];
            if (item.picture_url && String(item.picture_url).startsWith('data:image/') && !item.is_pc_header) {
                const publicUrl = await uploadDataUrlToStorage(item.picture_url, currentSheetId, item.row_index);
                if (publicUrl) {
                    item.picture_url = publicUrl;
                    rowIndexToPublicUrl[item.row_index] = publicUrl;
                }
            }
        }
        
        console.log(`📦 Prepared ${itemsToInsert.length} item(s) to insert`);
        
        // Insert all items (with retry for Failed to fetch)
        if (itemsToInsert.length > 0) {
            let insertResult;
            try {
                insertResult = await withRetry(async () => {
                    return await window.supabaseClient
                        .from('inventory_items')
                        .insert(itemsToInsert)
                        .select();
                });
            } catch (e) {
                insertResult = { data: null, error: e };
            }
            
            if (insertResult.error) {
                const insertError = insertResult.error;
                console.error('❌ Error inserting items to Supabase:', insertError);
                if (statusEl) { statusEl.textContent = 'Error saving'; statusEl.className = 'save-status error'; }
                alert(`Error saving items to database: ${insertError.message}\n\nKung "Failed to fetch", subukan ulit mamaya o check kung naka-pause ang Supabase project (free tier).`);
            } else {
                const insertedData = insertResult.data;
                console.log(`✅ Successfully saved ${itemsToInsert.length} item(s) to Supabase`);
                // I-update ang DOM at pictureUrls sa rows na na-upload ang picture — para lumabas ang image sa PC Location link
                const pictureUrls = sheets[currentSheetId].pictureUrls || [];
                for (const [idx, publicUrl] of Object.entries(rowIndexToPublicUrl)) {
                    const i = parseInt(idx, 10);
                    if (pictureUrls[i] !== undefined) pictureUrls[i] = publicUrl;
                    const row = rows[i];
                    if (row) {
                        const pictureCell = row.querySelector('.picture-cell img');
                        if (pictureCell) pictureCell.src = publicUrl;
                    }
                }
                sheets[currentSheetId].pictureUrls = pictureUrls;
            }
        } else {
            console.log('⚠️ No items to sync to Supabase (all rows are empty)');
        }
        if (statusEl && !statusEl.classList.contains('error')) {
            statusEl.textContent = 'Saved';
            statusEl.className = 'save-status saved';
            setTimeout(() => { statusEl.textContent = ''; statusEl.className = 'save-status'; }, 2500);
        }
    } catch (error) {
        console.error('❌ Unexpected error syncing to Supabase:', error);
        if (statusEl) {
            statusEl.textContent = 'Error saving';
            statusEl.className = 'save-status error';
        }
        alert(`Unexpected error: ${error.message}\n\nPlease check the browser console for more details.`);
    } finally {
        isSyncing = false;
    }
}

async function syncBackupToSupabase() {
    if (!checkSupabaseConnection() || isBackupMode) return;
    try {
        var payload = {
            sheets: sheets,
            currentSheetId: currentSheetId,
            sheetCounter: sheetCounter,
            savedAt: Date.now()
        };
        var res = await window.supabaseClient
            .from('inventory_backup_snapshot')
            .upsert({ id: 1, data: payload, updated_at: new Date().toISOString() }, { onConflict: 'id' });
        if (res.error) console.warn('Backup sync: upsert failed', res.error);
        else console.log('Inventory backup saved to Supabase (inventory_backup_snapshot)');
    } catch (e) {
        console.warn('Backup sync error', e);
    }
}

async function loadBackupFromSupabase() {
    if (!checkSupabaseConnection()) return null;
    try {
        var res = await window.supabaseClient
            .from('inventory_backup_snapshot')
            .select('data')
            .eq('id', 1)
            .maybeSingle();
        if (res.error || !res.data || !res.data.data) return null;
        var data = res.data.data;
        if (!data.sheets || typeof data.sheets !== 'object' || Object.keys(data.sheets).length === 0) return null;
        return {
            sheets: data.sheets,
            currentSheetId: data.currentSheetId || Object.keys(data.sheets)[0],
            sheetCounter: data.sheetCounter != null ? data.sheetCounter : 1
        };
    } catch (e) {
        console.warn('Load backup from Supabase error', e);
        return null;
    }
}

// Load data from Supabase
async function loadFromSupabase() {
    if (!checkSupabaseConnection()) return;
    
    // Set loading flag to prevent sync during load
    isLoadingFromSupabase = true;
    
    try {
        // Load all sheets
        const { data: sheetsData, error: sheetsError } = await window.supabaseClient
            .from('sheets')
            .select('*')
            .order('created_at', { ascending: true });
        
        if (sheetsError) {
            console.error('Error loading sheets from Supabase:', sheetsError);
            return;
        }
        
        const oldSheets = { ...sheets };
        sheets = {};
        sheetCounter = 0;
        console.log(`🗑️ Cleared ${Object.keys(oldSheets).length} existing sheet(s) before loading from Supabase`);
        
        // If no rows in sheets table, try loading from inventory_items (fallback so data still shows after refresh)
        let sheetsToLoad = sheetsData && sheetsData.length > 0 ? sheetsData : [];
        if (sheetsToLoad.length === 0) {
            console.log('No rows in sheets table; loading from inventory_items as fallback...');
            const { data: allItems, error: itemsErr } = await window.supabaseClient
                .from('inventory_items')
                .select('*')
                .order('sheet_id', { ascending: true })
                .order('row_index', { ascending: true });
            if (itemsErr || !allItems || allItems.length === 0) {
                console.log('No inventory_items found in Supabase; restoring from localStorage if any');
                var backup = loadBackupFromLocalStorage();
                var backupRows = backup && backup.sheets ? Object.values(backup.sheets).reduce(function(n, s) { return n + (s.data ? s.data.length : 0); }, 0) : 0;
                if (backupRows > 0) {
                    console.log('📂 Restoring ' + backupRows + ' row(s) from localStorage backup');
                    sheets = backup.sheets || {};
                    currentSheetId = backup.currentSheetId || 'sheet-1';
                    sheetCounter = backup.sheetCounter || 1;
                    updateSheetTabs();
                    var firstId = Object.keys(sheets)[0];
                    if (firstId) {
                        currentSheetId = firstId;
                        switchToSheet(firstId);
                        if (hasAnyData()) {
                            saveBackupToLocalStorage();
                            syncBackupToSupabase();
                        }
                    }
                } else if (Object.keys(sheets).length === 0) {
                    sheets['sheet-1'] = { id: 'sheet-1', name: 'Sheet 1', data: [], hasData: false, highlightStates: [], pictureUrls: [] };
                    currentSheetId = 'sheet-1';
                    updateSheetTabs();
                    switchToSheet('sheet-1');
                }
                return;
            }
            const bySheet = {};
            allItems.forEach(item => {
                const sid = item.sheet_id || 'sheet-1';
                if (!bySheet[sid]) bySheet[sid] = { id: sid, name: item.sheet_name || 'Sheet 1', items: [] };
                bySheet[sid].items.push(item);
            });
            Object.values(bySheet).forEach(s => s.items.sort((a, b) => (a.row_index ?? 0) - (b.row_index ?? 0)));
            sheetsToLoad = Object.values(bySheet);
        }
        
        // Load each sheet and its items
        for (const sheetData of sheetsToLoad) {
            let itemsData = sheetData.items; // from fallback
            if (!itemsData) {
                const res = await window.supabaseClient
                    .from('inventory_items')
                    .select('*')
                    .eq('sheet_id', sheetData.id)
                    .order('row_index', { ascending: true });
                if (res.error) {
                    console.error('Error loading items for sheet:', res.error);
                    continue;
                }
                itemsData = res.data || [];
            }
            
            // Convert items back to row data format
            const rowData = [];
            const highlightStates = [];
            const pictureUrls = []; // Store picture URLs separately
            
            console.log(`📥 Loading ${itemsData.length} item(s) for sheet "${sheetData.name}"`);
            
            itemsData.forEach((item, index) => {
                if (item.is_pc_header) {
                    rowData.push([item.article || '']);
                    highlightStates.push(false);
                    pictureUrls.push(null); // PC headers don't have pictures
                    console.log(`  [${index}] PC Header: ${item.article || '(empty)'}`);
                } else {
                    const row = [
                        item.article || '',
                        item.description || '',
                        item.old_property_n_assigned || '',
                        item.unit_of_meas || '',
                        item.unit_value || '',
                        item.quantity || '',
                        item.location || '',
                        item.condition || '',
                        item.remarks || '',
                        item.user || ''
                    ];
                    rowData.push(row);
                    highlightStates.push(item.is_highlighted || false);
                    pictureUrls.push(item.picture_url || null); // Store picture URL
                    console.log(`  [${index}] Item: ${item.article || '(empty)'} - ${item.description || '(empty)'}`);
                }
            });
            
            // Remove consecutive duplicate PC headers and duplicate data rows (so table hindi magulo)
            const deduped = [];
            const dedupedHighlight = [];
            const dedupedPictures = [];
            function rowEqual(a, b) {
                if (!a || !b || a.length !== b.length) return false;
                for (let k = 0; k < a.length; k++) {
                    if ((a[k] || '').toString().trim() !== (b[k] || '').toString().trim()) return false;
                }
                return true;
            }
            for (let i = 0; i < rowData.length; i++) {
                const r = rowData[i];
                const isPC = r && r.length === 1 && (r[0] || '').toString().trim() !== '';
                const prev = deduped.length > 0 ? deduped[deduped.length - 1] : null;
                const prevIsPC = prev && prev.length === 1;
                if (isPC && prevIsPC && (r[0] || '').toString().trim() === (prev[0] || '').toString().trim()) {
                    continue; // skip duplicate consecutive PC header
                }
                // Skip consecutive duplicate data row (same 10 columns)
                if (r && r.length >= 10 && prev && prev.length >= 10 && rowEqual(r, prev)) {
                    continue;
                }
                deduped.push(r);
                dedupedHighlight.push(highlightStates[i]);
                dedupedPictures.push(pictureUrls[i]);
            }
            if (deduped.length !== rowData.length) {
                console.log(`  🧹 Removed ${rowData.length - deduped.length} duplicate row(s)`);
                rowData.length = 0;
                rowData.push(...deduped);
                highlightStates.length = 0;
                highlightStates.push(...dedupedHighlight);
                pictureUrls.length = 0;
                pictureUrls.push(...dedupedPictures);
            }
            
            console.log(`✅ Converted ${rowData.length} row(s) for sheet "${sheetData.name}"`);
            
            // Create sheet
            sheets[sheetData.id] = {
                id: sheetData.id,
                name: sheetData.name,
                data: rowData,
                hasData: rowData.length > 0,
                highlightStates: highlightStates,
                pictureUrls: pictureUrls // Store picture URLs
            };
            
            console.log(`✅ Sheet "${sheetData.name}" created with ${sheets[sheetData.id].data.length} row(s) in memory`);
            
            // Update sheet counter
            const sheetNum = parseInt(sheetData.id.replace('sheet-', ''));
            if (sheetNum > sheetCounter) {
                sheetCounter = sheetNum;
            }
        }
        
        // Update sheet tabs first so tab elements exist before switching
        updateSheetTabs();
        
        // Kung walang data mula Supabase, gamitin localStorage backup — hindi dapat mawala ang data kahit mag-refresh
        if (!hasAnyData()) {
            const backup = loadBackupFromLocalStorage();
            const backupRows = backup && backup.sheets ? Object.values(backup.sheets).reduce((n, s) => n + (s.data ? s.data.length : 0), 0) : 0;
            if (backupRows > 0) {
                console.log(`📂 Supabase empty, restoring ${backupRows} row(s) from localStorage backup`);
                sheets = backup.sheets || {};
                currentSheetId = backup.currentSheetId || 'sheet-1';
                sheetCounter = backup.sheetCounter || 1;
                updateSheetTabs();
            }
        }
        
        const firstSheetId = Object.keys(sheets)[0];
        if (firstSheetId) {
            currentSheetId = firstSheetId;
            const sheet = sheets[firstSheetId];
            // If user added rows during load, keep table content
            const tbody = document.getElementById('tableBody');
            const tableRowCount = tbody ? tbody.querySelectorAll('tr:not(.empty-row)').length : 0;
            if (tableRowCount > 0 && sheet.data && sheet.data.length < tableRowCount) {
                saveCurrentSheetData(true);
            }
            console.log(`📋 Switching to sheet "${sheet.name}" with ${sheet.data ? sheet.data.length : 0} row(s)`);
            
            switchToSheet(firstSheetId);
        } else {
            console.log('⚠️ No sheets found to display');
        }
        
        console.log(`✅ Data loaded: ${Object.keys(sheets).length} sheet(s)`);
        
        // Log summary of loaded data
        Object.values(sheets).forEach(sheet => {
            console.log(`  - ${sheet.name}: ${sheet.data.length} row(s)`);
        });
        
        // I-update ang backup snapshot at Supabase para makita rin sa Backup page ang laman ng Inventory
        if (hasAnyData()) {
            saveBackupToLocalStorage();
            syncBackupToSupabase();
        }
    } catch (error) {
        console.error('Error loading from Supabase:', error);
        if (!hasAnyData()) {
            var backup = loadBackupFromLocalStorage();
            var backupRows = backup && backup.sheets ? Object.values(backup.sheets).reduce(function(n, s) { return n + (s.data ? s.data.length : 0); }, 0) : 0;
            if (backupRows > 0) {
                console.log('📂 Error during load; restoring ' + backupRows + ' row(s) from localStorage');
                sheets = backup.sheets || {};
                currentSheetId = backup.currentSheetId || 'sheet-1';
                updateSheetTabs();
                var firstId = Object.keys(sheets)[0];
                if (firstId) { currentSheetId = firstId; switchToSheet(firstId); }
            } else if (Object.keys(sheets).length === 0) {
                sheets['sheet-1'] = { id: 'sheet-1', name: 'Sheet 1', data: [], hasData: false, highlightStates: [], pictureUrls: [] };
                currentSheetId = 'sheet-1';
                updateSheetTabs();
                switchToSheet('sheet-1');
            }
        }
    } finally {
        // Clear loading flag after a short delay to ensure display is complete
        setTimeout(() => {
            isLoadingFromSupabase = false;
            console.log('✅ Loading complete, sync enabled');
        }, 1000);
    }
}

// Update sheet tabs UI
function updateSheetTabs() {
    const sheetTabs = document.getElementById('sheetTabs');
    sheetTabs.innerHTML = '';
    
    Object.values(sheets).forEach(sheet => {
        const tab = document.createElement('div');
        tab.className = 'sheet-tab' + (sheet.id === currentSheetId ? ' active' : '');
        tab.setAttribute('data-sheet-id', sheet.id);
        tab.innerHTML = `
            <span class="sheet-name">${sheet.name}</span>
            <span class="sheet-tab-menu" data-sheet-id="${sheet.id}" title="More">⋮</span>
            <div class="sheet-tab-dropdown">
                <button type="button" class="rename-option">Rename</button>
                <button type="button" class="delete-option">Delete</button>
            </div>
            <span class="close-sheet" data-sheet-id="${sheet.id}">×</span>
        `;
        
        tab.addEventListener('click', function(e) {
            if (!e.target.closest('.close-sheet') && !e.target.closest('.sheet-tab-menu') && !e.target.closest('.sheet-tab-dropdown')) {
                switchToSheet(sheet.id);
            }
        });
        
        const closeBtn = tab.querySelector('.close-sheet');
        closeBtn.addEventListener('click', function(e) {
            e.stopPropagation();
            deleteSheet(sheet.id);
        });
        
        setupSheetTabMenu(tab, sheet.id);
        sheetTabs.appendChild(tab);
    });
}

async function deleteSheet(sheetId) {
    if (Object.keys(sheets).length <= 1) {
        alert('Cannot delete the last sheet. At least one sheet must exist.');
        return;
    }
    
    if (confirm(`Are you sure you want to delete "${sheets[sheetId].name}"?`)) {
        // Delete from Supabase if connected
        if (checkSupabaseConnection()) {
            try {
                // Delete items first
                const { error: itemsError } = await window.supabaseClient
                    .from('inventory_items')
                    .delete()
                    .eq('sheet_id', sheetId);
                
                if (itemsError) {
                    console.error('Error deleting items from Supabase:', itemsError);
                }
                
                // Delete sheet
                const { error: sheetError } = await window.supabaseClient
                    .from('sheets')
                    .delete()
                    .eq('id', sheetId);
                
                if (sheetError) {
                    console.error('Error deleting sheet from Supabase:', sheetError);
                }
            } catch (error) {
                console.error('Error deleting from Supabase:', error);
            }
        }
        
        delete sheets[sheetId];
        
        // Remove tab
        const tab = document.querySelector(`[data-sheet-id="${sheetId}"]`);
        if (tab) tab.remove();
        
        // Switch to another sheet if current was deleted
        if (currentSheetId === sheetId) {
            const remainingSheets = Object.keys(sheets);
            if (remainingSheets.length > 0) {
                switchToSheet(remainingSheets[0]);
            }
        }
    }
}

function hasAnyData() {
    return Object.values(sheets).some(sheet => sheet.hasData);
}

// Update inventoryData from table
function updateDataFromTable() {
    saveCurrentSheetData();
}

// Convert blob: or https: image URL to data URL for Excel embed
function urlToDataUrl(url) {
    if (!url || typeof url !== 'string') return Promise.resolve(null);
    if (url.startsWith('data:image/')) return Promise.resolve(url);
    return fetch(url)
        .then(res => res.blob())
        .then(blob => new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = reject;
            reader.readAsDataURL(blob);
        }))
        .catch(err => {
            console.warn('Could not convert image URL to data URL:', err);
            return null;
        });
}

// Import Excel: read .xlsx/.xls and load into current sheet
document.getElementById('importExcelBtn').addEventListener('click', function() {
    document.getElementById('importExcelInput').click();
});
document.getElementById('importExcelInput').addEventListener('change', function(e) {
    var file = e.target.files && e.target.files[0];
    if (!file) return;
    e.target.value = '';
    var reader = new FileReader();
    reader.onload = async function(ev) {
        try {
            if (typeof XLSX === 'undefined') {
                alert('XLSX library not loaded. Please refresh and try again.');
                return;
            }
            var data = ev.target.result;
            var workbook = XLSX.read(new Uint8Array(data), { type: 'array', cellDates: true });
            var firstSheetName = workbook.SheetNames[0];
            if (!firstSheetName) {
                alert('No sheet found in the Excel file.');
                return;
            }
            var worksheet = workbook.Sheets[firstSheetName];
            var rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            if (!rows || rows.length === 0) {
                alert('No data found in the Excel file.');
                return;
            }
            // Find header row (contains "Article" and "Description")
            var headerRowIndex = -1;
            for (var i = 0; i < Math.min(30, rows.length); i++) {
                var row = rows[i];
                if (!row || !Array.isArray(row)) continue;
                var rowText = row.map(function(c) { return (c != null ? String(c) : '').toLowerCase(); }).join(' ');
                if (rowText.indexOf('article') !== -1 && rowText.indexOf('description') !== -1) {
                    headerRowIndex = i;
                    break;
                }
            }
            if (headerRowIndex < 0) headerRowIndex = 0;
            // Map column indices from header (export uses A=empty, B=Article, C=Description, ... K=User)
            var headerRow = rows[headerRowIndex];
            var colMap = { article: 1, description: 2, oldProperty: 3, unitMeas: 4, unitValue: 5, qty: 6, location: 7, condition: 8, remarks: 9, user: 10 };
            if (headerRow && headerRow.length > 0) {
                for (var c = 0; c < headerRow.length; c++) {
                    var h = (headerRow[c] != null ? String(headerRow[c]) : '').toLowerCase();
                    if (h.indexOf('article') !== -1 && (h.indexOf('item') !== -1 || h.indexOf('/') !== -1)) colMap.article = c;
                    else if (h.indexOf('description') !== -1) colMap.description = c;
                    else if (h.indexOf('old property') !== -1 || h.indexOf('property n') !== -1) colMap.oldProperty = c;
                    else if (h.indexOf('unit of meas') !== -1) colMap.unitMeas = c;
                    else if (h.indexOf('unit value') !== -1) colMap.unitValue = c;
                    else if (h.indexOf('quantity') !== -1 || h.indexOf('physical count') !== -1) colMap.qty = c;
                    else if (h.indexOf('location') !== -1 || h.indexOf('whereabout') !== -1) colMap.location = c;
                    else if (h.indexOf('condition') !== -1) colMap.condition = c;
                    else if (h.indexOf('remarks') !== -1) colMap.remarks = c;
                    else if (h === 'user') colMap.user = c;
                }
            }
            var sheetData = [];
            var highlightStates = [];
            var pictureUrls = [];
            for (var r = headerRowIndex + 1; r < rows.length; r++) {
                var row = rows[r];
                if (!row || !Array.isArray(row)) continue;
                var firstVal = (row.find(function(c) { return c != null && String(c).trim() !== ''; }) || row[0] || '').toString().trim();
                var firstUpper = firstVal.toUpperCase();
                var isPCHeader = firstVal && (
                    row.length === 1 ||
                    firstUpper === 'SERVER' || firstUpper === 'PC' ||
                    firstUpper.indexOf('PC USED BY') !== -1 ||
                    /^PC\s*\d*$/.test(firstUpper) || firstUpper.indexOf('PC ') === 0
                );
                if (isPCHeader) {
                    sheetData.push([firstVal]);
                    highlightStates.push(false);
                    pictureUrls.push(null);
                } else {
                    var art = row[colMap.article] != null ? String(row[colMap.article]).trim() : '';
                    var desc = row[colMap.description] != null ? String(row[colMap.description]).trim() : '';
                    var oldP = row[colMap.oldProperty] != null ? String(row[colMap.oldProperty]).trim() : '';
                    var uom = row[colMap.unitMeas] != null ? String(row[colMap.unitMeas]).trim() : '';
                    var uv = row[colMap.unitValue] != null ? String(row[colMap.unitValue]).trim() : '';
                    var qty = row[colMap.qty] != null ? String(row[colMap.qty]).trim() : '';
                    var loc = row[colMap.location] != null ? String(row[colMap.location]).trim() : '';
                    var cond = row[colMap.condition] != null ? String(row[colMap.condition]).trim() : '';
                    var rem = row[colMap.remarks] != null ? String(row[colMap.remarks]).trim() : '';
                    var user = row[colMap.user] != null ? String(row[colMap.user]).trim() : '';
                    sheetData.push([art, desc, oldP, uom, uv, qty, loc, cond, rem, user]);
                    highlightStates.push(false);
                    pictureUrls.push(null);
                }
            }
            if (sheetData.length === 0) {
                alert('No data rows found after the header in the Excel file.');
                return;
            }
            setCurrentSheetData(sheetData, true);
            var sheet = getCurrentSheet();
            sheet.highlightStates = highlightStates;
            sheet.pictureUrls = pictureUrls;
            displayData(sheetData);
            document.getElementById('exportBtn').disabled = false;
            document.getElementById('clearBtn').disabled = false;
            saveBackupToLocalStorage();
            if (checkSupabaseConnection()) {
                if (syncTimeout) clearTimeout(syncTimeout);
                syncTimeout = null;
                var statusEl = document.getElementById('saveStatus');
                if (statusEl) { statusEl.textContent = 'Saving to Supabase…'; statusEl.className = 'save-status saving'; }
                try {
                    await syncToSupabase();
                    if (hasAnyData()) await syncBackupToSupabase();
                    if (statusEl) { statusEl.textContent = 'Saved'; statusEl.className = 'save-status saved'; }
                } catch (err) {
                    console.warn('Import sync error', err);
                    if (statusEl) { statusEl.textContent = 'Import saved locally; sync failed'; statusEl.className = 'save-status error'; }
                }
            }
            alert('Imported ' + sheetData.length + ' row(s) from Excel.');
        } catch (err) {
            console.error('Import Excel error:', err);
            alert('Could not read the Excel file. Make sure it is a valid .xlsx or .xls file.');
        }
    };
    reader.readAsArrayBuffer(file);
});

// Export to Excel with ExcelJS (supports images and styling)
document.getElementById('exportBtn').addEventListener('click', async function() {
    updateDataFromTable();
    
    if (!hasAnyData()) {
        alert('No data to export. Please add items or import an Excel file first.');
        return;
    }
    
    // Check if ExcelJS is available
    if (typeof ExcelJS === 'undefined') {
        alert('ExcelJS library is loading. Please wait a moment and try again.');
        return;
    }
    
    try {
        const workbook = new ExcelJS.Workbook();
        
        // Export all sheets
        for (const sheet of Object.values(sheets)) {
            if (!sheet.hasData || sheet.data.length === 0) continue;
            
            const worksheet = workbook.addWorksheet(sheet.name);
            
            // Get current date
            const now = new Date();
            const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                           'July', 'August', 'September', 'October', 'November', 'December'];
            const currentDate = `${months[now.getMonth()]} ${now.getFullYear()}`;
            
            // Logo: 4 cm, naka-center sa kaliwang bahagi (col C)
            const logoSizeCm = 4;
            const logoSizePx = Math.round(logoSizeCm * 37.8);
            try {
                const logoResponse = await fetch('../images/omsc.png');
                const logoBuffer = await logoResponse.arrayBuffer();
                const imageId = workbook.addImage({
                    buffer: logoBuffer,
                    extension: 'png',
                });
                worksheet.addImage(imageId, {
                    tl: { col: 0, row: 0 },
                    ext: { width: logoSizePx, height: logoSizePx }
                });
            } catch (logoError) {
                console.warn('Could not load logo image:', logoError);
            }
            
            // Title block: walang unang empty column — naka-center (merge A–F, cols 1–6)
            const titleStartCol = 1;
            const titleEndCol = 6;
            const row1 = worksheet.addRow(['OCCIDENTAL MINDORO STATE COLLEGE']);
            worksheet.mergeCells(1, titleStartCol, 1, titleEndCol);
            const c1 = worksheet.getRow(1).getCell(titleStartCol);
            c1.value = 'OCCIDENTAL MINDORO STATE COLLEGE';
            c1.font = { bold: true, size: 18 };
            c1.alignment = { horizontal: 'center', vertical: 'middle' };
            
            const row2 = worksheet.addRow(['Multimedia and Speech Laboratory']);
            worksheet.mergeCells(2, titleStartCol, 2, titleEndCol);
            const c2 = worksheet.getRow(2).getCell(titleStartCol);
            c2.value = 'Multimedia and Speech Laboratory';
            c2.font = { bold: true, size: 14 };
            c2.alignment = { horizontal: 'center', vertical: 'middle' };
            
            const row3 = worksheet.addRow(['ICT Equipment, Devices & Accessories']);
            worksheet.mergeCells(3, titleStartCol, 3, titleEndCol);
            const c3 = worksheet.getRow(3).getCell(titleStartCol);
            c3.value = 'ICT Equipment, Devices & Accessories';
            c3.font = { bold: true, size: 12 };
            c3.alignment = { horizontal: 'center', vertical: 'middle' };
            
            const row4 = worksheet.addRow([`AS OF ${currentDate}`]);
            worksheet.mergeCells(4, titleStartCol, 4, titleEndCol);
            const c4 = worksheet.getRow(4).getCell(titleStartCol);
            c4.value = `AS OF ${currentDate}`;
            c4.font = { bold: true, size: 11 };
            c4.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            worksheet.getRow(4).commit();
            
            // Dalawang blank rows para sa space sa pagitan ng title at table
            worksheet.addRow([]);
            worksheet.addRow([]);
            
            // Header row: 11 columns (walang unang empty column)
            const headerLabels = [
                'Article/It',           // A
                'Description',          // B
                'Old Property N Assigned', // C
                'Unit of meas',         // D
                'Unit Value',           // E
                'Quantity per Physical count', // F
                'Location/Whereabout',  // G
                'Condition',            // H
                'Remarks',              // I
                'User',                 // J
                'Picture'               // K
            ];
            const headerRowNum = 7;
            const headerRow = worksheet.getRow(headerRowNum);
            headerRow.height = 22;
            const blackBorder = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
            for (let c = 1; c <= 11; c++) {
                const cell = headerRow.getCell(c);
                cell.value = headerLabels[c - 1];
                cell.font = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF495057' } };
                cell.border = blackBorder;
            }
            headerRow.commit();
            
            // Data rows simula row 8 (2 blank rows after title)
            let currentRow = 8;
            let dataRowIndex = 0;
            let exportSectionStart = null;
            let sectionPictureUrl = null;
            let sectionFirstRowData = null;
            let sectionName = '';
            const exportSections = [];   // E, F, K merge per section
            const dataRowsForPicture = []; // bawat row: { worksheetRow, rowIndex, rowData, sectionName } para sa Picture column (hindi na merged)
            
            sheet.data.forEach((rowData, rowIndex) => {
                // Skip empty rows
                if (!rowData || rowData.length === 0) {
                    return;
                }
                
                const firstCell = rowData[0] ? rowData[0].toString().trim() : '';
                
                // Check if this is a PC header row (export only PC name in one merged row, no other columns)
                // Match: [pcName] only, or first cell is "PC", "PC 1", "PC1", "PC USED BY...", "SERVER"
                const firstUpper = (firstCell || '').toUpperCase();
                const isPCHeader = firstCell && (
                    rowData.length === 1 ||
                    firstUpper === 'SERVER' ||
                    firstUpper === 'PC' ||
                    /^PC\s*\d*$/.test(firstUpper) ||           // PC, PC1, PC 2, etc.
                    /^PC\s+USED\s+BY/i.test(firstCell) ||
                    firstUpper.startsWith('PC ')
                );
                
                if (isPCHeader) {
                    if (exportSectionStart !== null) {
                        exportSections.push({ start: exportSectionStart, end: currentRow - 1, pictureUrl: sectionPictureUrl || null, rowData: sectionFirstRowData, sectionName: sectionName });
                        exportSectionStart = null;
                        sectionPictureUrl = null;
                        sectionFirstRowData = null;
                    }
                    const pcNameOnly = firstCell;
                    sectionName = pcNameOnly;
                    const pcRowValues = [pcNameOnly, '', '', '', '', '', '', '', '', '', ''];
                    const pcRow = worksheet.addRow(pcRowValues);
                    const grayFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD3D3D3' } };
                    const blackBorder = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
                    for (let col = 1; col <= 11; col++) {
                        const c = pcRow.getCell(col);
                        c.fill = grayFill;
                        c.border = blackBorder;
                        if (col >= 1 && col <= 11) {
                            c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                            c.font = { bold: true, size: 12 };
                        }
                    }
                    pcRow.getCell(1).value = pcNameOnly;
                    worksheet.mergeCells(currentRow, 1, currentRow, 11);
                } else if (rowData.length >= 10) {
                    const isFirstRowOfSection = (exportSectionStart === null);
                    if (exportSectionStart === null) exportSectionStart = currentRow;
                    const toStr = (val) => (val != null && String(val).trim() !== '' ? String(val).trim() : '');
                    if (isFirstRowOfSection && (sheet.pictureUrls && sheet.pictureUrls[rowIndex])) sectionPictureUrl = sheet.pictureUrls[rowIndex];
                    if (isFirstRowOfSection) sectionFirstRowData = rowData.slice(0, 10);
                    dataRowsForPicture.push({ worksheetRow: currentRow, rowIndex, rowData: rowData.slice(0, 10), sectionName });
                    const exportRow = [
                        toStr(toTitleCase(rowData[0])),  // A Article/It — Title Case
                        toStr(rowData[1]),  // B Description
                        toStr(rowData[2]),  // C Old Property N Assigned
                        toStr(rowData[3]),  // D Unit of meas
                        toStr(rowData[4]),  // E Unit Value
                        toStr(rowData[5]),  // F Quantity per Physical count
                        toStr(rowData[6]),  // G Location/Whereabout
                        toStr(rowData[7]),  // H Condition
                        toStr(rowData[8]),  // I Remarks
                        toStr(rowData[9]),  // J User
                        ''                 // K Picture (link/image added after)
                    ];
                    const dataRow = worksheet.addRow(exportRow);
                    dataRow.height = 30;   // Sapat para sa wrapped text — walang putol
                    
                    const conditionValue = (toStr(rowData[7]) || '').trim(); // Condition
                    // Kulay: Unserviceable = red, Borrowed = yellow lang; lahat ng iba (Serviceable, etc.) = puti
                    let fillColor = { argb: 'FFFFFFFF' }; // default puti
                    if (conditionValue === 'Borrowed') {
                        fillColor = { argb: 'FFFFFF3D' }; // Yellow
                    } else if (conditionValue === 'Unserviceable') {
                        fillColor = { argb: 'FFF8D7DA' }; // Light red
                    }
                    const cellFill = { type: 'pattern', pattern: 'solid', fgColor: fillColor };
                    const whiteFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
                    const blackBorder = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
                    for (let col = 1; col <= 11; col++) {
                        const cell = dataRow.getCell(col);
                        cell.border = blackBorder;
                        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                        if (cell.value) cell.font = { size: 9 };
                        cell.fill = (col === 4 || col === 5 || col === 10 || col === 11) ? whiteFill : cellFill; // D,E,J,K walang kulay
                    }
                    
                    dataRowIndex++;
                }
                currentRow++;
            });
            
            if (exportSectionStart !== null) {
                exportSections.push({ start: exportSectionStart, end: currentRow - 1, pictureUrl: sectionPictureUrl || null, rowData: sectionFirstRowData, sectionName: sectionName });
            }
            exportSections.forEach(s => {
                if (s.end >= s.start) {
                    worksheet.mergeCells(s.start, 4, s.end, 4);  // D Unit of meas
                    worksheet.mergeCells(s.start, 5, s.end, 5);  // E Unit Value
                    worksheet.mergeCells(s.start, 10, s.end, 10); // J User
                    // K Picture: hindi na merged — bawat row may sariling picture
                }
            });
            const baseUrl = window.location.origin + window.location.pathname.replace(/[^/]*$/, '');
            const pcLocationFormUrl = baseUrl + '../pc location/view.html';
            const pictureCol = 11;
            const toStr = (val) => (val != null && String(val).trim() !== '' ? String(val).trim() : '');
            // Bawat row: link at optional embed sa Picture column (per row)
            for (const r of dataRowsForPicture) {
                let picUrl = (sheet.pictureUrls && sheet.pictureUrls[r.rowIndex]) ? String(sheet.pictureUrls[r.rowIndex]).trim() : '';
                if (picUrl && picUrl.startsWith('blob:')) picUrl = await urlToDataUrl(picUrl);
                // Upload data URL to Storage during export so PC Location link gets https URL and picture shows
                if (picUrl && picUrl.startsWith('data:image/')) {
                    try {
                        const publicUrl = await uploadDataUrlToStorage(picUrl, currentSheetId, r.rowIndex);
                        if (publicUrl) picUrl = publicUrl;
                    } catch (e) { console.warn('Export: could not upload picture for link', e); }
                }
                const params = new URLSearchParams({
                    pcSection: toStr(r.sectionName),
                    article: toStr(toTitleCase(r.rowData[0])),
                    description: toStr(r.rowData[1]),
                    oldProperty: toStr(r.rowData[2]),
                    unitMeas: toStr(r.rowData[3]),
                    unitValue: toStr(r.rowData[4]),
                    qty: toStr(r.rowData[5]),
                    location: toStr(r.rowData[6]),
                    condition: toStr(r.rowData[7]),
                    remarks: toStr(r.rowData[8]),
                    user: toStr(r.rowData[9])
                });
                if (picUrl && (picUrl.startsWith('http://') || picUrl.startsWith('https://'))) params.set('image', picUrl);
                const linkUrl = pcLocationFormUrl + '?' + params.toString();
                const cell = worksheet.getRow(r.worksheetRow).getCell(pictureCol);
                cell.value = { text: linkUrl, hyperlink: linkUrl };
                cell.font = { size: 9, color: { argb: 'FF0066CC' }, underline: true };
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                if (picUrl && picUrl.startsWith('data:image/')) {
                    try {
                        const match = picUrl.match(/^data:image\/(\w+);base64,(.+)$/);
                        if (match) {
                            const ext = match[1].toLowerCase() === 'jpeg' ? 'jpeg' : match[1].toLowerCase();
                            const base64 = match[2];
                            const imageId = workbook.addImage({ base64: base64, extension: ext === 'jpg' ? 'jpeg' : ext });
                            const col0 = pictureCol - 1;
                            const row0 = r.worksheetRow - 1;
                            worksheet.addImage(imageId, {
                                tl: { col: col0 + 0.05, row: row0 + 0.05 },
                                br: { col: col0 + 0.95, row: row0 + 0.95 },
                                editAs: 'oneCell'
                            });
                        }
                    } catch (imgErr) {
                        console.warn('Could not embed picture in Excel:', imgErr);
                    }
                }
            }
            
            // Summary by Article/Item — sa baba ng data table (Item, Unit Measure, Existing, 2026–2029, For Disposal, Remarks)
            saveSummaryExtra();
            const summaryRows = computeSummaryData();
            const summaryExtra = (sheet && sheet.summaryExtra) ? sheet.summaryExtra : {};
            if (summaryRows.length > 0) {
                worksheet.addRow([]);
                const sumTitleRow = worksheet.addRow(['Summary by Article/Item']);
                sumTitleRow.getCell(1).font = { bold: true, size: 10 };
                worksheet.mergeCells(worksheet.rowCount, 1, worksheet.rowCount, 9);
                const sumHeaderRow = worksheet.addRow(['Item Categories/Names', 'Unit Measure', 'Existing', '2026', '2027', '2028', '2029', 'For Disposal', 'Remarks']);
                for (let c = 1; c <= 9; c++) sumHeaderRow.getCell(c).font = { bold: true };
                summaryRows.forEach(function(r) {
                    const ex = summaryExtra[r.item] || {};
                    const row = worksheet.addRow([
                        r.item || '',
                        r.unitMeasure || '—',
                        r.existing || 0,
                        ex.y2026 != null && ex.y2026 !== '' ? ex.y2026 : 0,
                        ex.y2027 != null && ex.y2027 !== '' ? ex.y2027 : 0,
                        ex.y2028 != null && ex.y2028 !== '' ? ex.y2028 : 0,
                        ex.y2029 != null && ex.y2029 !== '' ? ex.y2029 : 0,
                        ex.forDisposal != null && ex.forDisposal !== '' ? ex.forDisposal : 0,
                        ex.remarks != null ? ex.remarks : ''
                    ]);
                    for (let c = 1; c <= 9; c++) row.getCell(c).alignment = { horizontal: c === 3 || (c >= 4 && c <= 8) ? 'center' : 'left', vertical: 'middle' };
                });
            }
            
            // Prepared by / Noted by — sa baba ng table (name bold, role sa bagong linya, naka-center)
            worksheet.addRow([]);
            worksheet.addRow([]);
            const sigRow = worksheet.rowCount;
            const parseNameRole = (val) => {
                if (!val || !String(val).trim()) return { name: '', role: '' };
                const s = String(val).trim();
                const i = s.indexOf(',');
                if (i < 0) return { name: s, role: '' };
                return { name: s.substring(0, i).trim(), role: s.substring(i + 1).trim() };
            };
            const preparedEl = document.getElementById('preparedBy');
            const notedEl = document.getElementById('notedBy');
            const prepared = parseNameRole(preparedEl ? preparedEl.value : '');
            const noted = parseNameRole(notedEl ? notedEl.value : '');
            const fontSig = 9;
            worksheet.mergeCells(sigRow, 1, sigRow, 4);
            const preparedCell = worksheet.getRow(sigRow).getCell(1);
            preparedCell.value = {
                richText: [
                    { text: 'Prepared by:\n', font: { size: fontSig } },
                    { text: (prepared.name || '') + (prepared.role ? '\n' : ''), font: { size: fontSig, bold: true } },
                    { text: prepared.role || '', font: { size: fontSig } }
                ]
            };
            preparedCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            worksheet.mergeCells(sigRow, 7, sigRow, 11);
            const notedCell = worksheet.getRow(sigRow).getCell(7);
            notedCell.value = {
                richText: [
                    { text: 'Noted by:\n', font: { size: fontSig } },
                    { text: (noted.name || '') + (noted.role ? '\n' : ''), font: { size: fontSig, bold: true } },
                    { text: noted.role || '', font: { size: fontSig } }
                ]
            };
            notedCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            worksheet.getRow(sigRow).height = 36;
            
            // Column widths — 11 columns, mas compact para magkasya sa isang long bond
            const maxColumnWidths = [10, 12, 12, 8, 9, 8, 12, 8, 10, 10, 10];
            const columnHeaders = ['Article/It', 'Description', 'Old Property N Assigned', 'Unit of meas', 
                                   'Unit Value', 'Quantity per Physical count', 'Location/Whereabout', 
                                   'Condition', 'Remarks', 'User', 'Picture'];
            
            worksheet.eachRow((row, rowNumber) => {
                row.eachCell((cell, colNumber) => {
                    if (colNumber <= 11) {
                        const idx = colNumber - 1;
                        const cellValue = cell.value ? cell.value.toString() : '';
                        const cellLength = cellValue.length;
                        const headerLen = (columnHeaders[idx] || '').length;
                        let cap = (colNumber === 6) ? 8 : 14;
                        const estimatedWidth = Math.min(Math.max(cellLength * 1.2, headerLen * 1.2), cap);
                        if (estimatedWidth > maxColumnWidths[idx]) {
                            maxColumnWidths[idx] = Math.min(estimatedWidth, cap);
                        }
                    }
                });
            });
            
            worksheet.columns = maxColumnWidths.map((width, index) => ({
                width: Math.max(width, (columnHeaders[index] || '').length * 1.2)
            }));
            
            // Set row heights: title/header same as data rows (compact, like reference)
            worksheet.getRow(1).height = 20;
            worksheet.getRow(2).height = 18;
            worksheet.getRow(3).height = 18;
            worksheet.getRow(4).height = 18;
            worksheet.getRow(5).height = 15;
            worksheet.getRow(6).height = 15;
            worksheet.getRow(7).height = 20; // Column headers
            
            // Page setup — long bond landscape, fit sa isang bond paper
            worksheet.pageSetup = {
                paperSize: 14, // Long bond (8.5" x 13")
                orientation: 'landscape',
                fitToPage: true,
                fitToWidth: 1,
                fitToHeight: 1,
                margins: {
                    left: 0.25,
                    right: 0.25,
                    top: 0.35,
                    bottom: 0.35,
                    header: 0.2,
                    footer: 0.2
                },
                horizontalCentered: true,
                verticalCentered: false
            };
            
        }
        
        // Generate filename with current date
        const now = new Date();
        const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                       'July', 'August', 'September', 'October', 'November', 'December'];
        const filename = `INVENTORY-AS-OF-${months[now.getMonth()].toUpperCase()}-${now.getFullYear()}.xlsx`;
        
        // Write file
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        link.click();
        window.URL.revokeObjectURL(url);
        
        alert('Excel file exported successfully with ' + Object.values(sheets).filter(s => s.hasData).length + ' sheet(s)!');
    } catch (error) {
        alert('Error exporting Excel file: ' + error.message);
        console.error(error);
    }
});

// New Sheet button
document.getElementById('newSheetBtn').addEventListener('click', function() {
    const sheetName = prompt('Enter sheet name:', `Sheet ${sheetCounter + 1}`);
    if (sheetName !== null && sheetName.trim() !== '') {
        createNewSheet(sheetName.trim());
    }
});

// Initialize sheet tab click handlers (run after DOM is ready)
setTimeout(function() {
    const firstTab = document.querySelector('.sheet-tab');
    if (firstTab) {
        firstTab.addEventListener('click', function(e) {
            if (!e.target.closest('.close-sheet') && !e.target.closest('.sheet-tab-menu') && !e.target.closest('.sheet-tab-dropdown')) {
                switchToSheet('sheet-1');
            }
        });
        
        const closeBtn = firstTab.querySelector('.close-sheet');
        if (closeBtn) {
            closeBtn.addEventListener('click', function(e) {
                e.stopPropagation();
                deleteSheet('sheet-1');
            });
        }
        setupSheetTabMenu(firstTab, 'sheet-1');
    }
}, 100);

// Clear data
document.getElementById('clearBtn').addEventListener('click', function() {
    const currentSheet = getCurrentSheet();
    if (confirm(`Are you sure you want to clear all data in "${currentSheet.name}"?`)) {
        setCurrentSheetData([], false);
        document.getElementById('tableBody').innerHTML = 
            '<tr class="empty-row"><td colspan="13" class="empty-message">No data loaded. Please import an Excel file or add items manually.</td></tr>';
        document.getElementById('exportBtn').disabled = !hasAnyData();
        document.getElementById('clearBtn').disabled = true;
    }
});

(function() {
    const DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    var nextSheetId = 1;
    var sheets = [{ id: 1, name: 'COMPUTER LABORATORY', entries: [] }];
    var activeSheetId = 1;
    var isLoadingFromSupabase = false;
    var syncTimeout = null;
    var isBackupMode = typeof window !== 'undefined' && window.location && window.location.search.indexOf('backup=1') !== -1;
    var BACKUP_SNAPSHOT_KEY = 'class_schedule_backup_snapshot';
    var RESTORE_PAYLOAD_KEY = 'class_schedule_restore_payload';
    var categoryColors = { SCOA: '#ffe4c4', CBAM: '#b8d4e8', HM: '#fce4ec', CAST: '#b8d4e8', VACANT: '#d4edda' };
    var categoryToCssClass = { SCOA: 'cell-bg-pink', CBAM: 'cell-bg-light-blue', HM: 'cell-bg-light-pink', CAST: 'cell-bg-blue', VACANT: 'cell-bg-green' };

    function checkSupabaseConnection() {
        return typeof window !== 'undefined' && window.supabaseClient != null;
    }

    function getActiveSheet() {
        for (var i = 0; i < sheets.length; i++) if (sheets[i].id === activeSheetId) return sheets[i];
        return sheets[0] || null;
    }
    function getScheduleEntries() {
        var s = getActiveSheet();
        return s ? s.entries : [];
    }

    /* 4 rows before lunch, 5 after; 1-hour slots so 9-11 = 2 rows (9-10, 10-11) */
    var ROW_SLOTS = [
        { label: '7:30 - 9:00', start: 7*60+30, end: 9*60+0 },
        { label: '9:00 - 10:00', start: 9*60+0, end: 10*60+0 },
        { label: '10:00 - 11:00', start: 10*60+0, end: 11*60+0 },
        { label: '11:00 - 12:15', start: 11*60+0, end: 12*60+15 },
        { label: '12:15 - 1:00', start: 12*60+15, end: 13*60+0, lunch: true },
        { label: '1:00 - 2:30', start: 13*60+0, end: 14*60+30 },
        { label: '2:30 - 3:00', start: 14*60+30, end: 15*60+0 },
        { label: '3:00 - 4:00', start: 15*60+0, end: 16*60+0 },
        { label: '4:00 - 5:30', start: 16*60+0, end: 17*60+30 },
        { label: '5:30 - 8:30', start: 17*60+30, end: 20*60+30 }
    ];

    function parseTimeSlotToMinutes(str) {
        if (!str || typeof str !== 'string') return 0;
        var s = str.trim();
        if (s.toUpperCase() === 'LUNCH BREAK') return 12 * 60 + 15;
        var m = s.match(/(\d{1,2}):(\d{2})/);
        if (m) return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
        return 0;
    }

    function parseTimeRange(str) {
        if (!str || typeof str !== 'string') return { start: 0, end: 0 };
        var s = str.trim();
        var all = s.match(/\d{1,2}:\d{2}/g);
        if (!all || all.length === 0) return { start: 0, end: 0 };
        var minutes = all.map(function(t) {
            var p = t.split(':');
            return parseInt(p[0], 10) * 60 + parseInt(p[1], 10);
        });
        var start = minutes[0];
        var end = minutes[minutes.length - 1];
        if (end <= start) end = start + 60;
        return { start: start, end: end };
    }

    function entryOverlapsRow(entryStart, entryEnd, rowStart, rowEnd) {
        return entryStart < rowEnd && entryEnd > rowStart;
    }

    function getEntryDurationHours(entry) {
        var range = parseTimeRange(entry.timeSlot);
        return (range.end - range.start) / 60;
    }

    function getMaxRowsForEntry(entry) {
        var hours = getEntryDurationHours(entry);
        if (hours <= 1) return 1;
        if (hours <= 2) return 2;
        return Math.min(3, Math.ceil(hours)); // 3 hours = 3 rows, 9:00-12:00 merged
    }

    function getSegmentsForDay(day, entries) {
        if (entries == null) entries = getScheduleEntries();
        entries = entries.filter(function(e) { return e.day === day; });
        var rowCount = ROW_SLOTS.length;
        var rowAssignment = [];
        var entryRowsUsed = {};
        for (var r = 0; r < rowCount; r++) {
            var row = ROW_SLOTS[r];
            if (row.lunch) {
                rowAssignment[r] = { type: 'lunch' };
                continue;
            }
            var found = null;
            for (var i = 0; i < entries.length; i++) {
                var range = parseTimeRange(entries[i].timeSlot);
                if (entryOverlapsRow(range.start, range.end, row.start, row.end)) {
                    found = entries[i];
                    break;
                }
            }
            if (found) {
                var key = found.day + '|' + found.timeSlot;
                var used = entryRowsUsed[key] || 0;
                var maxRows = getMaxRowsForEntry(found);
                if (used < maxRows) {
                    entryRowsUsed[key] = used + 1;
                    rowAssignment[r] = { type: 'entry', entry: found };
                } else {
                    rowAssignment[r] = null;
                }
            } else {
                rowAssignment[r] = null;
            }
        }
        var segments = [];
        var r = 0;
        while (r < rowCount) {
            if (rowAssignment[r] && rowAssignment[r].type === 'entry') {
                var entry = rowAssignment[r].entry;
                var startR = r;
                while (r < rowCount && rowAssignment[r] && rowAssignment[r].type === 'entry' && rowAssignment[r].entry === entry) r++;
                segments.push({ type: 'entry', entry: entry, startRow: startR, endRow: r - 1, rowspan: r - startR });
            } else {
                if (rowAssignment[r] && rowAssignment[r].type === 'lunch') {
                    segments.push({ type: 'lunch', startRow: r, endRow: r, rowspan: 1 });
                }
                r++;
            }
        }
        return segments;
    }

    function getSegmentAtRow(segments, rowIndex) {
        for (var i = 0; i < segments.length; i++) {
            var s = segments[i];
            if (rowIndex >= s.startRow && rowIndex <= s.endRow) return { segment: s, isFirst: rowIndex === s.startRow };
        }
        return null;
    }

    function getCategoryFromCode(code) {
        if (!code || !String(code).trim()) return null;
        var first = String(code).trim().split(/\s+/)[0] || '';
        if (first === 'BSAIS' || first === 'BSOA') return 'SCOA';
        if (first === 'BSBA' || first === 'OM' || first === 'BPA') return 'CBAM';
        if (first === 'HM') return 'HM';
        if (first === 'BSDC') return 'CAST';
        return null;
    }

    function getCourseColorClass(entry) {
        if (entry.type === 'VACANT' || entry.type === 'BREAK' || entry.type === 'LUNCH BREAK' || entry.type === 'VACANT/LAB MAINTENANCE') return 'cell-bg-green';
        var cat = getCategoryFromCode(entry.code);
        return cat && categoryToCssClass[cat] ? categoryToCssClass[cat] : '';
    }

    function hexToArgb(hex) {
        if (!hex) return '';
        var h = hex.replace(/^#/, '');
        if (h.length === 6) return 'FF' + h.toUpperCase();
        if (h.length === 8) return h.toUpperCase();
        return '';
    }

    function getCourseColorArgb(cell) {
        if (cell.type === 'VACANT' || cell.type === 'BREAK' || cell.type === 'LUNCH BREAK' || cell.type === 'VACANT/LAB MAINTENANCE') {
            var v = categoryColors.VACANT;
            return v ? hexToArgb(v) : 'FFD4EDDA';
        }
        var cat = getCategoryFromCode(cell.code);
        var hex = cat && categoryColors[cat] ? categoryColors[cat] : null;
        return hex ? hexToArgb(hex) : null;
    }

    function applyCategoryColorsStyle() {
        var sel = document.getElementById('legend-category-colors');
        if (!sel) {
            sel = document.createElement('style');
            sel.id = 'legend-category-colors';
            document.head.appendChild(sel);
        }
        var s = '';
        ['SCOA', 'CBAM', 'HM', 'CAST', 'VACANT'].forEach(function(cat) {
            var hex = categoryColors[cat] || '#fff';
            var cls = categoryToCssClass[cat];
            if (cls) s += '.schedule-grid .cell-slot.' + cls + ' { background: ' + hex + ' !important; }\n';
            s += '#legendTable tbody tr[data-category="' + cat + '"] td { background: ' + hex + ' !important; }\n';
            s += '#programsTable tbody tr[data-category="' + cat + '"] td { background: ' + hex + ' !important; }\n';
        });
        sel.textContent = s;
    }

    function getDisplayTimeSlots() {
        return ROW_SLOTS.map(function(r) { return r.label; });
    }

    function findEntry(day, timeSlot, entries) {
        if (entries == null) entries = getScheduleEntries();
        for (var i = 0; i < entries.length; i++) {
            if (entries[i].day === day && entries[i].timeSlot === timeSlot) return entries[i];
        }
        return null;
    }

    function setEntry(day, timeSlot, type, instructor, course, code) {
        var ent = getActiveSheet();
        if (!ent) return;
        var entries = ent.entries;
        var existing = findEntry(day, timeSlot, entries);
        var entry = { day: day, timeSlot: timeSlot, type: type || '', instructor: instructor || '', course: course || '', code: code || '' };
        if (existing) {
            var idx = entries.indexOf(existing);
            entries[idx] = entry;
        } else {
            entries.push(entry);
        }
        debouncedSyncActiveSheet();
    }

    function deleteEntry(day, timeSlot) {
        var ent = getActiveSheet();
        if (!ent) return;
        ent.entries = ent.entries.filter(function(e) { return e.day !== day || e.timeSlot !== timeSlot; });
        debouncedSyncActiveSheet();
    }

    function updateEntryTime(day, oldTimeSlot, newTimeSlot) {
        var t = (newTimeSlot || '').trim();
        if (!t) return;
        var entry = findEntry(day, oldTimeSlot);
        if (entry) entry.timeSlot = t;
        debouncedSyncActiveSheet();
    }

    function renderScheduleGrid() {
        var tbody = document.getElementById('scheduleBody');
        if (!tbody) {
            console.warn('Schedule: #scheduleBody not found.');
            return;
        }
        tbody.innerHTML = '';
        var daySegments = DAYS.map(function(day) { return getSegmentsForDay(day); });
        var skipLeft = DAYS.map(function() { return 0; });
        var skipTimeLeft = DAYS.map(function() { return 0; });
        for (var rowIndex = 0; rowIndex < ROW_SLOTS.length; rowIndex++) {
            var row = ROW_SLOTS[rowIndex];
            var timeLabel = row.label;
            var tr = document.createElement('tr');
            if (row.lunch) {
                var lunchTd = document.createElement('td');
                lunchTd.setAttribute('colspan', String(DAYS.length * 2));
                lunchTd.className = 'cell-lunch-break';
                lunchTd.textContent = 'LUNCH BREAK';
                tr.appendChild(lunchTd);
            } else {
                DAYS.forEach(function(day, colIndex) {
                    if (skipLeft[colIndex] > 0) {
                        skipLeft[colIndex]--;
                    } else {
                        var info = getSegmentAtRow(daySegments[colIndex], rowIndex);
                        var contentTd = document.createElement('td');
                        contentTd.className = 'cell-slot cell-display';
                        if (info && info.isFirst && info.segment.type === 'entry') {
                            var entry = info.segment.entry;
                            contentTd.dataset.day = day;
                            contentTd.dataset.timeslot = entry.timeSlot;
                            if (info.segment.rowspan > 1) contentTd.setAttribute('rowspan', String(info.segment.rowspan));
                            if (entry.type) {
                                var specialTypes = ['VACANT', 'VACANT/LAB MAINTENANCE', 'BREAK', 'LUNCH BREAK'];
                                var hasDetails = (entry.instructor && entry.instructor.trim()) || (entry.course && entry.course.trim()) || (entry.code && entry.code.trim());
                                var isSpecialTypeWithDetails = specialTypes.indexOf(entry.type) >= 0 && hasDetails;
                                if (isSpecialTypeWithDetails) {
                                    contentTd.classList.add('cell-class');
                                    var colorClass = getCourseColorClass(entry);
                                    if (colorClass) contentTd.classList.add(colorClass);
                                    var wrap = document.createElement('div');
                                    wrap.className = 'cell-class-content';
                                    var typeLine = document.createElement('div');
                                    typeLine.className = 'cell-line-type-vacant';
                                    typeLine.textContent = entry.type;
                                    wrap.appendChild(typeLine);
                                    var line1 = document.createElement('div');
                                    line1.className = 'cell-line-instructor';
                                    line1.textContent = entry.instructor || '';
                                    var line2 = document.createElement('div');
                                    line2.className = 'cell-line-subject';
                                    line2.textContent = entry.course || '';
                                    var line3 = document.createElement('div');
                                    line3.className = 'cell-line-code';
                                    line3.textContent = entry.code || '';
                                    wrap.appendChild(line1);
                                    wrap.appendChild(line2);
                                    wrap.appendChild(line3);
                                    contentTd.appendChild(wrap);
                                } else {
                                    contentTd.textContent = entry.type;
                                    var colorClass = getCourseColorClass(entry);
                                    if (colorClass) contentTd.classList.add(colorClass);
                                }
                            } else {
                                contentTd.classList.add('cell-class');
                                var colorClass = getCourseColorClass(entry);
                                if (colorClass) contentTd.classList.add(colorClass);
                                var wrap = document.createElement('div');
                                wrap.className = 'cell-class-content';
                                var line1 = document.createElement('div');
                                line1.className = 'cell-line-instructor';
                                line1.textContent = entry.instructor || '';
                                var line2 = document.createElement('div');
                                line2.className = 'cell-line-subject';
                                line2.textContent = entry.course || '';
                                var line3 = document.createElement('div');
                                line3.className = 'cell-line-code';
                                line3.textContent = entry.code || '';
                                wrap.appendChild(line1);
                                wrap.appendChild(line2);
                                wrap.appendChild(line3);
                                contentTd.appendChild(wrap);
                            }
                            if (!isBackupMode) {
                                var delBtn = document.createElement('button');
                                delBtn.type = 'button';
                                delBtn.className = 'cell-delete';
                                delBtn.textContent = '\u00D7';
                                delBtn.title = 'Remove';
                                delBtn.addEventListener('click', function(ev) {
                                    ev.preventDefault();
                                    deleteEntry(day, entry.timeSlot);
                                    renderScheduleGrid();
                                });
                                contentTd.appendChild(delBtn);
                            }
                            skipLeft[colIndex] = info.segment.rowspan - 1;
                        } else {
                            contentTd.className = 'cell-slot cell-display cell-empty';
                        }
                        tr.appendChild(contentTd);
                    }
                    var timeTd = null;
                    if (skipTimeLeft[colIndex] > 0) {
                        skipTimeLeft[colIndex]--;
                    } else {
                        timeTd = document.createElement('td');
                        timeTd.className = 'cell-time';
                        var timeInfo = getSegmentAtRow(daySegments[colIndex], rowIndex);
                        if (timeInfo && timeInfo.segment.type === 'entry') {
                            if (timeInfo.isFirst && timeInfo.segment.rowspan > 1) {
                                timeTd.setAttribute('rowspan', String(timeInfo.segment.rowspan));
                                timeTd.textContent = timeInfo.segment.entry.timeSlot || timeLabel;
                                skipTimeLeft[colIndex] = timeInfo.segment.rowspan - 1;
                            } else {
                                timeTd.textContent = timeInfo.isFirst ? (timeInfo.segment.entry.timeSlot || timeLabel) : timeLabel;
                            }
                            if (timeInfo.isFirst && !isBackupMode) {
                            var entry = timeInfo.segment.entry;
                            var timeEditBtn = document.createElement('button');
                            timeEditBtn.type = 'button';
                            timeEditBtn.className = 'cell-time-menu';
                            timeEditBtn.textContent = '\u22EE';
                            timeEditBtn.title = 'Edit time';
                            timeEditBtn.setAttribute('aria-label', 'Edit time');
                            timeEditBtn.addEventListener('click', function(ev) {
                                ev.preventDefault();
                                ev.stopPropagation();
                                var newVal = prompt('Edit time slot (e.g. 9:00 - 11:00):', entry.timeSlot || '');
                                if (newVal != null) {
                                    var trimmed = newVal.trim();
                                    if (trimmed && trimmed !== entry.timeSlot) {
                                        updateEntryTime(day, entry.timeSlot, trimmed);
                                        renderScheduleGrid();
                                    }
                                }
                            });
                            timeTd.appendChild(timeEditBtn);
                            }
                        } else {
                            timeTd.textContent = '';
                        }
                        tr.appendChild(timeTd);
                    }
                });
            }
            tbody.appendChild(tr);
        }
    }

    function getGridData(entries) {
        if (entries == null) entries = getScheduleEntries();
        var data = [];
        var daySegments = DAYS.map(function(day) { return getSegmentsForDay(day, entries); });
        for (var r = 0; r < ROW_SLOTS.length; r++) {
            var row = [];
            DAYS.forEach(function(day, colIndex) {
                var info = getSegmentAtRow(daySegments[colIndex], r);
                if (info && info.segment.type === 'lunch') {
                    row.push({ type: 'LUNCH BREAK' });
                } else if (info && info.segment.type === 'entry') {
                    var entry = info.segment.entry;
                    var specialTypes = ['VACANT', 'VACANT/LAB MAINTENANCE', 'BREAK', 'LUNCH BREAK'];
                    var hasDetails = (entry.instructor && entry.instructor.trim()) || (entry.course && entry.course.trim()) || (entry.code && entry.code.trim());
                    var specialWithDetails = entry.type && specialTypes.indexOf(entry.type) >= 0 && hasDetails;
                    if (entry.type && !specialWithDetails) row.push({ type: entry.type });
                    else row.push({ type: entry.type || null, instructor: entry.instructor || '', course: entry.course || '', code: entry.code || '' });
                } else {
                    row.push({ type: '', instructor: '', course: '', code: '' });
                }
            });
            data.push(row);
        }
        return data;
    }

    function getTableData(tableId, hasHeader) {
        const table = document.getElementById(tableId);
        if (!table) return [];
        const rows = [];
        const trs = table.querySelectorAll('tbody tr');
        trs.forEach(function(tr) {
            const row = [];
            tr.querySelectorAll('td input, td').forEach(function(cell) {
                const inp = cell.querySelector && cell.querySelector('input');
                row.push(inp ? inp.value : (cell.textContent || '').trim());
            });
            rows.push(row);
        });
        return rows;
    }

    function fillOneSheet(workbook, sheet) {
        const ws = workbook.addWorksheet(sheet.name);
        var thinBorder = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        var isComputerLab = (sheet.name === 'COMPUTER LABORATORY');
        var fontMain = 8;
        var fontHead = 9;
        const headerText = {
            line1: 'Republic of the Philippines',
            line2: 'OCCIDENTAL MINDORO STATE COLLEGE',
            line3: 'COLLEGE OF ARTS, SCIENCES AND TECHNOLOGY',
            line4: 'San Jose, Occidental Mindoro',
            line5: 'Website: www.omsc.edu.ph  Email address: omsc_9747@yahoo.com',
            line6: 'Tele Fax No.: (043) 491-1460',
            lab: isComputerLab ? 'COMPUTER LABORATORY - MAIN CAMPUS' : sheet.name,
            semester: '2nd Semester 2025 - 2026'
        };

        const logoCol = 1;
        const textStartCol = 2;
        const textEndCol = 7;
        let row = 1;

        ws.getRow(1).height = 12;
        ws.getRow(2).height = 12;
        ws.getRow(3).height = 12;
        ws.getRow(4).height = 12;
        ws.getRow(5).height = 12;
        ws.getRow(6).height = 12;

        ws.mergeCells(row, textStartCol, row + 5, textEndCol);
        const headerCell = ws.getCell(row, textStartCol);
        headerCell.value = {
            richText: [
                { text: headerText.line1 + '\n', font: { size: fontMain } },
                { text: headerText.line2 + '\n', font: { size: fontMain, bold: true } },
                { text: headerText.line3 + '\n', font: { size: fontMain, bold: true } },
                { text: headerText.line4 + '\n', font: { size: fontMain } },
                { text: headerText.line5 + '\n', font: { size: fontMain } },
                { text: headerText.line6, font: { size: fontMain } }
            ]
        };
        headerCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

        var logoImg = document.querySelector('.page-header .header-logo');
        if (logoImg && logoImg.complete && logoImg.naturalWidth) {
            try {
                var canvas = document.createElement('canvas');
                canvas.width = logoImg.naturalWidth;
                canvas.height = logoImg.naturalHeight;
                var ctx = canvas.getContext('2d');
                ctx.drawImage(logoImg, 0, 0);
                var dataUrl = canvas.toDataURL('image/png');
                var base64 = dataUrl.replace(/^data:image\/\w+;base64,/, '');
                var imageId = workbook.addImage({ base64: base64, extension: 'png' });
                ws.addImage(imageId, {
                    tl: { nativeCol: logoCol - 1, nativeRow: row - 1 },
                    ext: { width: 90, height: 90 }
                });
            } catch (e) { /* skip logo on CORS/canvas error */ }
        }

        row += 7;
        ws.getRow(row).height = 18;
        ws.getCell(row, 1).value = headerText.lab;
        ws.getCell(row, 1).font = { bold: true, size: fontHead, underline: true };
        ws.getCell(row, 1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        ws.getCell(row, 1).border = thinBorder;
        ws.mergeCells(row, 1, row, 12);
        for (var bc = 2; bc <= 12; bc++) { ws.getCell(row, bc).border = thinBorder; }
        row++;
        ws.getRow(row).height = 12;
        ws.getCell(row, 1).value = headerText.semester;
        ws.getCell(row, 1).font = isComputerLab ? { bold: true, size: fontMain } : { size: fontMain };
        ws.getCell(row, 1).alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getCell(row, 1).border = thinBorder;
        ws.mergeCells(row, 1, row, 12);
        for (var bc = 2; bc <= 12; bc++) { ws.getCell(row, bc).border = thinBorder; }
        row += 1;

        const headerRow = row;
        const dayNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
        dayNames.forEach(function(d, i) {
            ws.mergeCells(row, i * 2 + 1, row, i * 2 + 2);
            const c = ws.getCell(row, i * 2 + 1);
            c.value = d;
            c.font = { bold: true, size: fontMain, color: { argb: 'FFFFFFFF' } };
            c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4A90A4' } };
            c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            c.border = thinBorder;
            ws.getCell(row, i * 2 + 2).border = thinBorder;
        });
        ws.getRow(row).height = 14;
        row++;
        dayNames.forEach(function(_, i) {
            ws.getCell(row, i * 2 + 1).value = 'Schedule';
            ws.getCell(row, i * 2 + 1).font = { bold: true, size: fontMain, color: { argb: 'FFFFFFFF' } };
            ws.getCell(row, i * 2 + 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF5A6C8A' } };
            ws.getCell(row, i * 2 + 1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            ws.getCell(row, i * 2 + 1).border = thinBorder;
            ws.getCell(row, i * 2 + 2).value = 'Time';
            ws.getCell(row, i * 2 + 2).font = { bold: true, size: fontMain, color: { argb: 'FFFFFFFF' } };
            ws.getCell(row, i * 2 + 2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF5A6C8A' } };
            ws.getCell(row, i * 2 + 2).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            ws.getCell(row, i * 2 + 2).border = thinBorder;
        });
        ws.getRow(row).height = 16;
        row++;

        const gridData = getGridData(sheet.entries);
        const timeLabels = ROW_SLOTS.map(function(r) { return r.lunch ? '12:15 - 1:00' : r.label; });
        const yellowFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFF3CD' } };
        const pureYellowFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
        const greenFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD4EDDA' } };
        var scheduleStartRow = row;
        var lunchRowIndex = -1;
        ROW_SLOTS.forEach(function(r, i) { if (r.lunch) lunchRowIndex = i; });
        gridData.forEach(function(rowData, r) {
            if (r === lunchRowIndex) {
                ws.mergeCells(row, 1, row, 12);
                var lunchCell = ws.getCell(row, 1);
                lunchCell.value = 'LUNCH BREAK';
                lunchCell.fill = pureYellowFill;
                lunchCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                lunchCell.border = thinBorder;
                lunchCell.font = { bold: true, size: fontMain };
                ws.getRow(row).height = 18;
                row++;
                return;
            }
            rowData.forEach(function(cell, c) {
                const contentCol = c * 2 + 1;
                const timeCol = c * 2 + 2;
                const contentCell = ws.getCell(row, contentCol);
                contentCell.border = thinBorder;
                contentCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                var specialTypes = ['VACANT', 'VACANT/LAB MAINTENANCE', 'BREAK', 'LUNCH BREAK'];
                var cellHasDetails = (cell.instructor && cell.instructor.trim()) || (cell.course && cell.course.trim()) || (cell.code && cell.code.trim());
                var specialWithDetails = cell.type && specialTypes.indexOf(cell.type) >= 0 && cellHasDetails;
                if (cell.type && !specialWithDetails) {
                    contentCell.value = cell.type;
                    contentCell.font = { size: fontMain };
                    if (cell.type === 'VACANT' || cell.type === 'VACANT/LAB MAINTENANCE') contentCell.fill = greenFill;
                    else contentCell.fill = yellowFill;
                } else {
                    var parts = [];
                    if (cell.type && specialTypes.indexOf(cell.type) >= 0) {
                        parts.push({ font: { bold: true, size: fontMain }, text: cell.type + '\n' });
                    }
                    if (cell.instructor) { parts.push({ font: { bold: true, size: fontMain }, text: cell.instructor + '\n' }); }
                    if (cell.course) { parts.push({ font: { size: fontMain }, text: cell.course + '\n' }); }
                    if (cell.code) { parts.push({ font: { bold: true, size: fontMain }, text: cell.code }); }
                    if (parts.length) contentCell.value = { richText: parts }; else contentCell.value = cell.type || '';
                    var fillArgb = getCourseColorArgb(cell);
                    if (fillArgb) contentCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillArgb } };
                }
                const timeCell = ws.getCell(row, timeCol);
                timeCell.value = timeLabels[r];
                timeCell.font = { size: fontMain };
                timeCell.border = thinBorder;
                timeCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            });
            ws.getRow(row).height = 32;
            row++;
        });
        function getCellDisplayText(val) {
            if (val == null) return '';
            if (typeof val === 'string') return val;
            if (val.richText && Array.isArray(val.richText)) return val.richText.map(function(r) { return r.text; }).join('');
            return String(val);
        }
        for (var contentCol = 1; contentCol <= 11; contentCol += 2) {
            var runStart = scheduleStartRow;
            for (var ri = scheduleStartRow + 1; ri < row; ri++) {
                if (ri === lunchRowIndex + scheduleStartRow) {
                    if (ri - 1 >= runStart) ws.mergeCells(runStart, contentCol, ri - 1, contentCol);
                    runStart = ri + 1;
                    continue;
                }
                if (ri - 1 === lunchRowIndex + scheduleStartRow) { runStart = ri; continue; }
                var vPrev = ws.getCell(ri - 1, contentCol).value;
                var vCur = ws.getCell(ri, contentCol).value;
                var same = getCellDisplayText(vPrev) === getCellDisplayText(vCur);
                if (!same) {
                    if (ri - 1 > runStart) ws.mergeCells(runStart, contentCol, ri - 1, contentCol);
                    runStart = ri;
                }
            }
            if (row - 1 > runStart) ws.mergeCells(runStart, contentCol, row - 1, contentCol);
        }

        row += 1;
        var bottomStartRow = row;
        var thin = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        var legendColors = {};
        ['SCOA', 'CBAM', 'HM', 'CAST', 'VACANT'].forEach(function(c) { legendColors[c] = hexToArgb(categoryColors[c] || '#fff'); });
        var progColors = {};
        ['SCOA', 'CBAM', 'HM', 'CAST'].forEach(function(c) { progColors[c] = hexToArgb(categoryColors[c] || '#fff'); });

        var legendTable = document.getElementById('legendTable');
        if (legendTable) {
            ws.getCell(row, 1).value = 'LEGEND';
            ws.getCell(row, 1).font = { bold: true, size: fontMain };
            ws.getCell(row, 1).border = thin;
            ws.getCell(row, 2).border = thin;
            ws.getCell(row, 3).border = thin;
            ws.getCell(row, 1).alignment = { horizontal: 'center', vertical: 'middle' };
            row++;
            legendTable.querySelectorAll('tbody tr').forEach(function(tr) {
                var cat = (tr.getAttribute('data-category') || '').toUpperCase();
                var fill = legendColors[cat] ? { type: 'pattern', pattern: 'solid', fgColor: { argb: legendColors[cat] } } : {};
                var cells = tr.querySelectorAll('td');
                var c1 = (cells[0] && cells[0].querySelector('input')) ? cells[0].querySelector('input').value : cells[0].textContent.trim();
                var c2 = (cells[1] && cells[1].querySelector('input')) ? cells[1].querySelector('input').value : cells[1].textContent.trim();
                ws.getCell(row, 1).value = c1;
                ws.getCell(row, 1).border = thin;
                ws.getCell(row, 1).font = { size: fontMain };
                ws.getCell(row, 1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                if (fill.fgColor) ws.getCell(row, 1).fill = fill;
                ws.mergeCells(row, 2, row, 3);
                var nameCell = ws.getCell(row, 2);
                nameCell.value = c2;
                nameCell.border = thin;
                nameCell.font = { size: 7 };
                nameCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                ws.getCell(row, 3).border = thin;
                ws.getRow(row).height = 24;
                row++;
            });
        }

        row = bottomStartRow;
        ws.getCell(row, 4).value = 'Number of Existing PC';
        ws.getCell(row, 4).font = { bold: true, size: fontMain };
        [4,5,6,7,8].forEach(function(col) { ws.getCell(row, col).border = thin; ws.getCell(row, col).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; });
        row++;
        var pcTable = document.getElementById('pcTable');
        if (pcTable) {
            var pcHead = pcTable.querySelector('thead tr');
            if (pcHead) {
                pcHead.querySelectorAll('th').forEach(function(th, ci) {
                    var c = ws.getCell(row, 4 + ci);
                    var inp = th.querySelector('input.pc-year-header');
                    c.value = (inp ? inp.value : th.textContent || '').trim();
                    c.font = { bold: true, size: fontMain, color: { argb: 'FFFFFFFF' } };
                    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF495057' } };
                    c.border = thin;
                    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                });
                row++;
            }
            var pcRows = pcTable.querySelectorAll('tbody tr');
            pcRows.forEach(function(tr, ri) {
                var isOverall = tr.classList.contains('overall-row');
                var cells = tr.querySelectorAll('td');
                for (var ci = 0; ci < cells.length; ci++) {
                    var cell = ws.getCell(row, 4 + ci);
                    var td = cells[ci];
                    var inp = td.querySelector('input');
                    cell.value = inp ? (inp.value || '') : td.textContent.trim();
                    cell.border = thin;
                    cell.font = { size: fontMain };
                    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                    if (isOverall) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
                }
                ws.getRow(row).height = 16;
                row++;
            });
        }

        row = bottomStartRow;
        ws.getCell(row, 10).value = 'COMPUTER PROGRAMS USED:';
        ws.getCell(row, 10).font = { bold: true, size: fontMain };
        ws.getCell(row, 10).border = thin;
        ws.getCell(row, 11).border = thin;
        ws.getCell(row, 10).alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getCell(row, 11).alignment = { horizontal: 'center', vertical: 'middle' };
        row++;
        var progTable = document.getElementById('programsTable');
        if (progTable) {
            var progHead = progTable.querySelector('thead tr');
            if (progHead) {
                progHead.querySelectorAll('th').forEach(function(th, ci) {
                    var c = ws.getCell(row, 10 + ci);
                    c.value = th.textContent.trim();
                    c.font = { bold: true, size: fontMain, color: { argb: 'FFFFFFFF' } };
                    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF495057' } };
                    c.border = thin;
                    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                });
                row++;
            }
            progTable.querySelectorAll('tbody tr').forEach(function(tr) {
                var cat = (tr.getAttribute('data-category') || '').toUpperCase();
                var fill = progColors[cat] ? { type: 'pattern', pattern: 'solid', fgColor: { argb: progColors[cat] } } : {};
                var cells = tr.querySelectorAll('td');
                var c1 = (cells[0] && cells[0].querySelector('input')) ? cells[0].querySelector('input').value : cells[0].textContent.trim();
                var c2 = (cells[1] && cells[1].querySelector('input')) ? cells[1].querySelector('input').value : cells[1].textContent.trim();
                ws.getCell(row, 10).value = c1;
                ws.getCell(row, 10).border = thin;
                ws.getCell(row, 10).font = { size: fontMain };
                ws.getCell(row, 10).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                if (fill.fgColor) ws.getCell(row, 10).fill = fill;
                ws.getCell(row, 11).value = c2;
                ws.getCell(row, 11).border = thin;
                ws.getCell(row, 11).font = { size: fontMain };
                ws.getCell(row, 11).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                if (fill.fgColor) ws.getCell(row, 11).fill = fill;
                ws.getRow(row).height = 16;
                row++;
            });
        }

        var lastTableRow = Math.max(legendTable ? bottomStartRow + 1 + (legendTable.querySelectorAll('tbody tr').length) : 0,
            pcTable ? bottomStartRow + 2 + (pcTable.querySelectorAll('tbody tr').length) : 0,
            progTable ? bottomStartRow + 2 + (progTable.querySelectorAll('tbody tr').length) : 0);
        row = lastTableRow + 1;
        var preparedEl = document.getElementById('preparedBy');
        var notedEl = document.getElementById('notedBy');
        function parseNameRole(val) {
            if (!val || !val.trim()) return { name: '', role: '' };
            var s = val.trim();
            var i = s.indexOf(',');
            if (i < 0) return { name: s, role: '' };
            return { name: s.substring(0, i).trim(), role: s.substring(i + 1).trim() };
        }
        var prepared = parseNameRole(preparedEl ? preparedEl.value : '');
        var noted = parseNameRole(notedEl ? notedEl.value : '');
        var preparedCell = ws.getCell(row, 1);
        preparedCell.value = {
            richText: [
                { text: 'Prepared by:\n', font: { size: fontMain } },
                { text: (prepared.name || '') + (prepared.role ? '\n' : ''), font: { size: fontMain, bold: true } },
                { text: prepared.role || '', font: { size: fontMain } }
            ]
        };
        preparedCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        ws.mergeCells(row, 1, row, 5);
        var notedCell = ws.getCell(row, 9);
        notedCell.value = {
            richText: [
                { text: 'Noted by:\n', font: { size: fontMain } },
                { text: (noted.name || '') + (noted.role ? '\n' : ''), font: { size: fontMain, bold: true } },
                { text: noted.role || '', font: { size: fontMain } }
            ]
        };
        notedCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        ws.mergeCells(row, 9, row, 12);
        ws.getRow(row).height = 36;
        row++;

        const colWidths = [24, 10, 24, 10, 24, 10, 24, 10, 24, 10, 24, 10];
        colWidths.forEach(function(w, i) { ws.getColumn(i + 1).width = w; });

        ws.pageSetup = {
            orientation: 'landscape',
            paperSize: 14, // Long bond (8.5" x 13") - isang bondpaper landscape
            fitToPage: true,
            fitToWidth: 1,
            fitToHeight: 1,
            margins: { left: 0.25, right: 0.25, top: 0.25, bottom: 0.25, header: 0.1, footer: 0.1 },
            horizontalCentered: true,
            verticalCentered: false,
            printArea: 'A1:L' + row
        };

        return ws;
    }

    function renderSheetTabs() {
        var bar = document.getElementById('sheetTabsBar');
        if (!bar) return;
        bar.innerHTML = '';
        sheets.forEach(function(sheet) {
            var tab = document.createElement('div');
            tab.className = 'sheet-tab' + (sheet.id === activeSheetId ? ' active' : '');
            tab.setAttribute('data-sheet-id', String(sheet.id));
            var label = document.createElement('span');
            label.className = 'sheet-tab-label';
            label.textContent = sheet.name;
            tab.appendChild(label);
            var closeBtn = document.createElement('button');
            closeBtn.type = 'button';
            closeBtn.className = 'sheet-tab-close';
            closeBtn.innerHTML = '&times;';
            closeBtn.setAttribute('aria-label', 'Close sheet');
            closeBtn.addEventListener('click', function(ev) {
                ev.stopPropagation();
                if (sheets.length <= 1) return;
                var idx = sheets.findIndex(function(s) { return s.id === sheet.id; });
                if (idx < 0) return;
                if (activeSheetId === sheet.id) {
                    if (idx > 0) activeSheetId = sheets[idx - 1].id;
                    else if (sheets.length > 1) activeSheetId = sheets[1].id;
                }
                var sheetIdToDelete = sheet.id;
                sheets.splice(idx, 1);
                renderSheetTabs();
                renderScheduleGrid();
                if (checkSupabaseConnection()) {
                    window.supabaseClient.from('class_schedule_sheets').delete().eq('id', sheetIdToDelete).then(function() {}).catch(function(err) { console.error('Delete sheet error:', err); });
                }
            });
            tab.appendChild(closeBtn);
            tab.addEventListener('click', function(ev) {
                if (ev.target === closeBtn) return;
                activeSheetId = sheet.id;
                renderSheetTabs();
                renderScheduleGrid();
            });
            bar.appendChild(tab);
        });
    }

    function getCellText(val) {
        if (val == null || val === '') return '';
        if (typeof val === 'string') return val.trim();
        return String(val).trim();
    }

    function parseContentCell(val) {
        var text = getCellText(val);
        if (!text) return { type: '', instructor: '', course: '', code: '' };
        if (text.toUpperCase() === 'LUNCH BREAK') return null;
        var lines = text.split(/\r?\n/).map(function(s) { return s.trim(); }).filter(Boolean);
        var type = lines[0] || '';
        var instructor = lines[1] || '';
        var course = lines[2] || '';
        var code = lines[3] || '';
        return { type: type, instructor: instructor, course: course, code: code };
    }

    function exportToExcel() {
        if (typeof ExcelJS === 'undefined') {
            alert('ExcelJS is loading. Please wait and try again.');
            return;
        }
        const workbook = new ExcelJS.Workbook();
        for (var i = 0; i < sheets.length; i++) {
            fillOneSheet(workbook, sheets[i]);
        }
        const headerText = { semester: '2nd Semester 2025 - 2026' };
        workbook.xlsx.writeBuffer().then(function(buffer) {
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'LAB-SCHED-' + (headerText.semester.replace(/\s+/g, '-') || 'Schedule') + '.xlsx';
            a.click();
            URL.revokeObjectURL(url);
            alert('Excel file downloaded.');
        }).catch(function(err) {
            alert('Export failed: ' + (err && err.message));
        });
    }

    function importFromExcel(file) {
        if (typeof XLSX === 'undefined') {
            alert('XLSX library is loading. Please wait and try again.');
            return;
        }
        var reader = new FileReader();
        reader.onload = async function(ev) {
            try {
                var data = new Uint8Array(ev.target.result);
                var workbook = XLSX.read(data, { type: 'array', cellDates: true });
                var sheetName = workbook.SheetNames[0];
                if (!sheetName) { alert('No sheet in Excel file.'); return; }
                var ws = workbook.Sheets[sheetName];
                var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
                if (!rows || rows.length < 5) { alert('No data found in Excel.'); return; }
                var headerRowIndex = -1;
                for (var r = 0; r < Math.min(20, rows.length); r++) {
                    var row = rows[r];
                    if (!Array.isArray(row)) continue;
                    var first = (row[0] != null ? String(row[0]) : '').toLowerCase();
                    if (first.indexOf('monday') !== -1 || (row[1] != null && String(row[1]).toLowerCase().indexOf('tuesday') !== -1)) {
                        headerRowIndex = r;
                        break;
                    }
                }
                if (headerRowIndex < 0) { alert('Could not find schedule header (Monday, Tuesday...) in Excel.'); return; }
                var dataStartRow = headerRowIndex + 2;
                var entries = [];
                var lunchRowIndex = 4;
                var slotIndex = 0;
                for (var r = 0; r < 10; r++) {
                    var excelRow = rows[dataStartRow + r];
                    if (!excelRow || !Array.isArray(excelRow)) continue;
                    var firstCell = getCellText(excelRow[0]);
                    if (firstCell.toUpperCase().indexOf('LUNCH') !== -1) {
                        slotIndex++;
                        continue;
                    }
                    var timeLabel = r < ROW_SLOTS.length ? ROW_SLOTS[r].label : '';
                    for (var dayIdx = 0; dayIdx < 6; dayIdx++) {
                        var contentCol = dayIdx * 2;
                        var timeCol = dayIdx * 2 + 1;
                        var contentVal = excelRow[contentCol];
                        var parsed = parseContentCell(contentVal);
                        if (!parsed) continue;
                        if (!parsed.type && !parsed.instructor && !parsed.course && !parsed.code) continue;
                        var timeSlot = getCellText(excelRow[timeCol]) || timeLabel;
                        entries.push({
                            day: DAYS[dayIdx],
                            timeSlot: timeSlot,
                            type: parsed.type || '',
                            instructor: parsed.instructor || '',
                            course: parsed.course || '',
                            code: parsed.code || ''
                        });
                    }
                    slotIndex++;
                }
                var sheet = getActiveSheet();
                if (!sheet) { alert('No active sheet.'); return; }
                sheet.entries = entries;
                renderScheduleGrid();
                saveBackupSnapshot();
                if (checkSupabaseConnection()) {
                    try {
                        await syncEntriesForSheet(activeSheetId);
                        await syncBackupToSupabase();
                    } catch (err) {
                        console.warn('Import sync error', err);
                    }
                }
                alert('Imported ' + entries.length + ' schedule entry/entries from Excel.');
            } catch (err) {
                console.error('Import Excel error', err);
                alert('Could not read Excel file. Make sure it is a valid .xlsx or .xls from this app\'s Export.');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function syncEntriesForSheet(sheetId) {
        if (!checkSupabaseConnection()) return Promise.resolve();
        var sheet = sheets.filter(function(s) { return s.id === sheetId; })[0];
        if (!sheet) return Promise.resolve();
        var supabase = window.supabaseClient;
        return supabase.from('class_schedule_entries').delete().eq('sheet_id', sheetId).then(function() {
            var rows = (sheet.entries || []).map(function(e) {
                return { sheet_id: sheetId, day: e.day, time_slot: e.timeSlot || e.time_slot || '', type: e.type || '', instructor: e.instructor || '', course: e.course || '', code: e.code || '' };
            });
            if (rows.length === 0) return Promise.resolve();
            return supabase.from('class_schedule_entries').insert(rows);
        }).then(function() {
            if (!isBackupMode) {
                saveBackupSnapshot();
                syncBackupToSupabase();
            }
        }).catch(function(err) { console.error('Sync entries error:', err); });
    }

    function saveBackupSnapshot() {
        if (isBackupMode) return;
        try {
            var snapshot = { sheets: sheets, activeSheetId: activeSheetId, nextSheetId: nextSheetId, savedAt: Date.now() };
            localStorage.setItem(BACKUP_SNAPSHOT_KEY, JSON.stringify(snapshot));
        } catch (e) { }
    }

    function loadBackupSnapshot() {
        try {
            var raw = localStorage.getItem(BACKUP_SNAPSHOT_KEY);
            if (!raw) return null;
            var snap = JSON.parse(raw);
            if (!snap || !Array.isArray(snap.sheets) || snap.sheets.length === 0) return null;
            return snap;
        } catch (e) { return null; }
    }

    async function syncBackupToSupabase() {
        if (isBackupMode || !checkSupabaseConnection()) return;
        try {
            var payload = { sheets: sheets, activeSheetId: activeSheetId, nextSheetId: nextSheetId, savedAt: Date.now() };
            var res = await window.supabaseClient
                .from('class_schedule_backup_snapshot')
                .upsert({ id: 1, data: payload, updated_at: new Date().toISOString() }, { onConflict: 'id' });
            if (res.error) console.warn('Schedule backup sync failed', res.error);
            else console.log('Class Schedule backup saved to Supabase (class_schedule_backup_snapshot)');
        } catch (e) { console.warn('Schedule backup sync error', e); }
    }

    async function loadBackupFromSupabase() {
        if (!checkSupabaseConnection()) return null;
        try {
            var res = await window.supabaseClient
                .from('class_schedule_backup_snapshot')
                .select('data')
                .eq('id', 1)
                .maybeSingle();
            if (res.error || !res.data || !res.data.data) return null;
            var data = res.data.data;
            if (!data.sheets || !Array.isArray(data.sheets) || data.sheets.length === 0) return null;
            return {
                sheets: data.sheets,
                activeSheetId: data.activeSheetId != null ? data.activeSheetId : (data.sheets[0] && data.sheets[0].id),
                nextSheetId: data.nextSheetId != null ? data.nextSheetId : 1
            };
        } catch (e) { console.warn('Load schedule backup from Supabase error', e); return null; }
    }

    function debouncedSyncActiveSheet() {
        if (syncTimeout) clearTimeout(syncTimeout);
        syncTimeout = setTimeout(function() {
            syncTimeout = null;
            if (getActiveSheet()) syncEntriesForSheet(activeSheetId);
        }, 300);
    }

    async function loadFromSupabase() {
        if (!checkSupabaseConnection() || isLoadingFromSupabase) return;
        isLoadingFromSupabase = true;
        try {
            var supabase = window.supabaseClient;
            var res = await supabase.from('class_schedule_sheets').select('id, name').order('id', { ascending: true });
            if (res.error) { console.error('Load sheets error:', res.error); return; }
            var sheetsData = res.data || [];
            if (sheetsData.length === 0) {
                await supabase.from('class_schedule_sheets').insert({ name: 'COMPUTER LABORATORY' });
                res = await supabase.from('class_schedule_sheets').select('id, name').order('id', { ascending: true });
                sheetsData = (res.data || []);
            }
            if (sheetsData.length === 0) {
                isLoadingFromSupabase = false;
                return;
            }
            sheets = [];
            var maxId = 0;
            for (var i = 0; i < sheetsData.length; i++) {
                var row = sheetsData[i];
                var sid = row.id;
                if (sid > maxId) maxId = sid;
                var entriesRes = await supabase.from('class_schedule_entries').select('*').eq('sheet_id', sid).order('id', { ascending: true });
                var entriesRows = (entriesRes.data || []).map(function(r) {
                    return { day: r.day, timeSlot: r.time_slot, type: r.type || '', instructor: r.instructor || '', course: r.course || '', code: r.code || '' };
                });
                sheets.push({ id: sid, name: row.name, entries: entriesRows });
            }
            nextSheetId = maxId + 1;
            activeSheetId = sheets[0].id;
            renderSheetTabs();
            renderScheduleGrid();
            if (!isBackupMode) {
                saveBackupSnapshot();
                syncBackupToSupabase();
            }
            console.log('Class Schedule: loaded ' + sheets.length + ' sheet(s) from Supabase');
        } catch (e) {
            console.error('Load from Supabase:', e);
        } finally {
            isLoadingFromSupabase = false;
        }
    }

    function makeScheduleReadOnly() {
        var addScheduleBtn = document.getElementById('addScheduleBtn');
        var addScheduleMenu = document.getElementById('addScheduleMenu');
        var addSheetBtn = document.getElementById('addSheetBtn');
        var addSheetMenu = document.getElementById('addSheetMenu');
        var addFormPanel = document.getElementById('addFormPanel');
        if (addScheduleBtn) { addScheduleBtn.style.display = 'none'; addScheduleBtn.setAttribute('aria-hidden', 'true'); }
        if (addScheduleMenu) addScheduleMenu.hidden = true;
        if (addSheetBtn) { addSheetBtn.style.display = 'none'; addSheetBtn.setAttribute('aria-hidden', 'true'); }
        if (addSheetMenu) addSheetMenu.hidden = true;
        if (addFormPanel) addFormPanel.hidden = true;
        var importBtn = document.getElementById('importExcelBtn');
        if (importBtn) { importBtn.style.display = 'none'; }
        var importInput = document.getElementById('importExcelInput');
        if (importInput) { importInput.style.display = 'none'; }
        document.querySelectorAll('.cell-delete').forEach(function(btn) { btn.style.display = 'none'; });
        document.querySelectorAll('.cell-time-menu').forEach(function(btn) { btn.style.display = 'none'; });
        document.querySelectorAll('#legendTable input, #pcTable input, #programsTable input, #preparedBy, #notedBy').forEach(function(inp) {
            if (inp) { inp.disabled = true; inp.readOnly = true; }
        });
        var restoreBtn = document.getElementById('restoreScheduleBtn');
        if (restoreBtn) restoreBtn.style.display = 'inline-block';
    }

    function init() {
        var exportBtn = document.getElementById('exportExcelBtn');
        if (exportBtn) exportBtn.addEventListener('click', exportToExcel);
        var importBtn = document.getElementById('importExcelBtn');
        var importInput = document.getElementById('importExcelInput');
        if (importBtn && importInput) {
            importBtn.addEventListener('click', function() { importInput.click(); });
            importInput.addEventListener('change', function(e) {
                var file = e.target.files && e.target.files[0];
                if (file) {
                    importFromExcel(file);
                    e.target.value = '';
                }
            });
        }
        document.querySelectorAll('.legend-color-picker').forEach(function(inp) {
            var cat = inp.getAttribute('data-category');
            if (cat && categoryColors.hasOwnProperty(cat)) {
                categoryColors[cat] = inp.value || categoryColors[cat];
                inp.addEventListener('input', function() {
                    categoryColors[cat] = inp.value;
                    applyCategoryColorsStyle();
                });
                inp.addEventListener('change', function() {
                    categoryColors[cat] = inp.value;
                    applyCategoryColorsStyle();
                });
            }
        });
        applyCategoryColorsStyle();
        var restoreJustApplied = false;

        if (isBackupMode) {
            var restoreBtn = document.getElementById('restoreScheduleBtn');
            if (restoreBtn) {
                restoreBtn.addEventListener('click', function() {
                    if (!confirm('I-restore ang backup na ito sa main Class Schedule page? Mapapalitan ang laman ng main page.')) return;
                    try {
                        var payload = { sheets: sheets, activeSheetId: activeSheetId, nextSheetId: nextSheetId };
                        localStorage.setItem(RESTORE_PAYLOAD_KEY, JSON.stringify(payload));
                        window.location.href = 'index.html';
                    } catch (e) {
                        alert('Restore failed: ' + (e && e.message));
                    }
                });
            }
        } else {
            var restorePayload = null;
            try {
                var raw = localStorage.getItem(RESTORE_PAYLOAD_KEY);
                if (raw) {
                    restorePayload = JSON.parse(raw);
                    localStorage.removeItem(RESTORE_PAYLOAD_KEY);
                }
            } catch (e) { }
            if (restorePayload && restorePayload.sheets && restorePayload.sheets.length > 0) {
                sheets = restorePayload.sheets;
                activeSheetId = restorePayload.activeSheetId != null ? restorePayload.activeSheetId : (sheets[0] && sheets[0].id);
                if (restorePayload.nextSheetId != null) nextSheetId = restorePayload.nextSheetId;
                renderSheetTabs();
                renderScheduleGrid();
                if (checkSupabaseConnection()) {
                    sheets.forEach(function(s) { syncEntriesForSheet(s.id); });
                }
                saveBackupSnapshot();
                syncBackupToSupabase();
                if (document.getElementById('restoreScheduleBtn')) document.getElementById('restoreScheduleBtn').style.display = 'none';
                restoreJustApplied = true;
            }
        }

        var addSheetBtn = document.getElementById('addSheetBtn');
        var addSheetMenu = document.getElementById('addSheetMenu');
        function toggleSheetMenu() {
            if (addSheetMenu) {
                addSheetMenu.hidden = !addSheetMenu.hidden;
                if (addSheetBtn) addSheetBtn.setAttribute('aria-expanded', addSheetMenu.hidden ? 'false' : 'true');
            }
        }
        function closeSheetMenu() {
            if (addSheetMenu) {
                addSheetMenu.hidden = true;
                if (addSheetBtn) addSheetBtn.setAttribute('aria-expanded', 'false');
            }
        }
        if (addSheetBtn && addSheetMenu) {
            addSheetBtn.addEventListener('click', function(ev) {
                ev.stopPropagation();
                toggleSheetMenu();
            });
            document.addEventListener('click', function() { closeSheetMenu(); });
            addSheetMenu.addEventListener('click', function(ev) { ev.stopPropagation(); });
        }
        document.querySelectorAll('.sheet-option-btn').forEach(function(btn) {
            btn.addEventListener('click', function() {
                var name = btn.getAttribute('data-sheet');
                if (!name) return;
                function addSheetWithId(id) {
                    var newSheet = { id: id, name: name, entries: [] };
                    sheets.push(newSheet);
                    if (id >= nextSheetId) nextSheetId = id + 1;
                    activeSheetId = newSheet.id;
                    renderSheetTabs();
                    renderScheduleGrid();
                    closeSheetMenu();
                }
                if (checkSupabaseConnection()) {
                    window.supabaseClient.from('class_schedule_sheets').insert({ name: name }).select('id').single()
                        .then(function(res) {
                            if (res.data && res.data.id != null) addSheetWithId(res.data.id);
                            else { nextSheetId++; addSheetWithId(nextSheetId); }
                        })
                        .catch(function(err) {
                            console.error('Insert sheet error:', err);
                            nextSheetId++;
                            addSheetWithId(nextSheetId);
                        });
                } else {
                    nextSheetId++;
                    addSheetWithId(nextSheetId);
                }
            });
        });
        if (isBackupMode) {
            var snap = loadBackupSnapshot();
            if (snap && snap.sheets && snap.sheets.length > 0) {
                sheets = snap.sheets;
                activeSheetId = snap.activeSheetId != null ? snap.activeSheetId : (sheets[0] && sheets[0].id);
                if (snap.nextSheetId != null) nextSheetId = snap.nextSheetId;
                renderSheetTabs();
                renderScheduleGrid();
                makeScheduleReadOnly();
            } else {
                (function waitThenLoad(attempt) {
                    if (attempt > 25) { makeScheduleReadOnly(); return; }
                    if (checkSupabaseConnection()) {
                        loadBackupFromSupabase().then(function(snapFromSupabase) {
                            if (snapFromSupabase && snapFromSupabase.sheets && snapFromSupabase.sheets.length > 0) {
                                sheets = snapFromSupabase.sheets;
                                activeSheetId = snapFromSupabase.activeSheetId != null ? snapFromSupabase.activeSheetId : (sheets[0] && sheets[0].id);
                                if (snapFromSupabase.nextSheetId != null) nextSheetId = snapFromSupabase.nextSheetId;
                                renderSheetTabs();
                                renderScheduleGrid();
                            }
                            makeScheduleReadOnly();
                        }).catch(function() { makeScheduleReadOnly(); });
                    } else {
                        setTimeout(function() { waitThenLoad(attempt + 1); }, 300);
                    }
                })(0);
            }
        } else if (!restoreJustApplied && checkSupabaseConnection()) {
            loadFromSupabase();
        } else if (!restoreJustApplied) {
            renderSheetTabs();
            renderScheduleGrid();
        }

        window.addEventListener('beforeunload', function() {
            if (syncTimeout) { clearTimeout(syncTimeout); syncTimeout = null; }
            if (checkSupabaseConnection() && getActiveSheet()) {
                syncEntriesForSheet(activeSheetId);
                if (!isBackupMode) { saveBackupSnapshot(); syncBackupToSupabase(); }
            }
        });

        var addFormPanel = document.getElementById('addFormPanel');
        var selectedDayLabel = document.getElementById('selectedDayLabel');
        var formTimeSlot = document.getElementById('formTimeSlot');
        var formType = document.getElementById('formType');
        var formClassFields = document.getElementById('formClassFields');
        var formInstructor = document.getElementById('formInstructor');
        var formCourse = document.getElementById('formCourse');
        var formCode = document.getElementById('formCode');
        var formYear = document.getElementById('formYear');
        var formSection = document.getElementById('formSection');
        var formStudentCount = document.getElementById('formStudentCount');
        var formYearRow = document.getElementById('formYearRow');
        var formSectionRow = document.getElementById('formSectionRow');
        var formStudentCountRow = document.getElementById('formStudentCountRow');

        function showForm(day) {
            if (selectedDayLabel) selectedDayLabel.textContent = day;
            if (addFormPanel) {
                addFormPanel.hidden = false;
                if (formClassFields) formClassFields.style.display = 'block';
            }
        }
        function hideForm() {
            if (addFormPanel) addFormPanel.hidden = true;
            if (formInstructor) formInstructor.value = '';
            if (formCourse) formCourse.value = '';
            if (formCode) formCode.value = '';
            if (formYear) { formYear.value = ''; if (formYearRow) formYearRow.style.display = 'none'; }
            if (formSection) { formSection.value = ''; if (formSectionRow) formSectionRow.style.display = 'none'; }
            if (formStudentCount) { formStudentCount.value = ''; if (formStudentCountRow) formStudentCountRow.style.display = 'none'; }
            if (formType) formType.value = '';
            if (formTimeSlot) formTimeSlot.value = '';
        }
        function onFormSave() {
            var day = selectedDayLabel ? selectedDayLabel.textContent : '';
            var timeSlot = formTimeSlot ? formTimeSlot.value.trim() : '';
            if (!day || !timeSlot || !formType) {
                if (!timeSlot) alert('Please enter a time slot (e.g. 7:30 - 9:30).');
                return;
            }
            var type = formType.value || '';
            var instructor = formInstructor ? formInstructor.value.trim() : '';
            var course = formCourse ? formCourse.value.trim() : '';
            var codeCourse = formCode ? formCode.value.trim() : '';
            var codeYear = formYear ? formYear.value : '';
            var codeSection = formSection ? formSection.value.trim() : '';
            var codeCount = formStudentCount ? formStudentCount.value.trim() : '';
            var code = (codeCourse && codeYear && codeSection && codeCount)
                ? (codeCourse + ' ' + codeYear + '-' + codeSection + ' ' + codeCount + ' Student')
                : (codeCourse || '');
            setEntry(day, timeSlot, type, instructor, course, code);
            if (checkSupabaseConnection() && getActiveSheet()) {
                setTimeout(function() { syncEntriesForSheet(activeSheetId); }, 150);
            }
            renderScheduleGrid();
            hideForm();
        }

        var addScheduleBtn = document.getElementById('addScheduleBtn');
        var addScheduleMenu = document.getElementById('addScheduleMenu');
        function toggleMenu() {
            if (addScheduleMenu) {
                addScheduleMenu.hidden = !addScheduleMenu.hidden;
                if (addScheduleBtn) addScheduleBtn.setAttribute('aria-expanded', addScheduleMenu.hidden ? 'false' : 'true');
            }
        }
        function closeMenu() {
            if (addScheduleMenu) {
                addScheduleMenu.hidden = true;
                if (addScheduleBtn) addScheduleBtn.setAttribute('aria-expanded', 'false');
            }
        }
        if (addScheduleBtn && addScheduleMenu) {
            addScheduleBtn.addEventListener('click', function(ev) {
                ev.stopPropagation();
                toggleMenu();
            });
            document.addEventListener('click', function() { closeMenu(); });
            addScheduleMenu.addEventListener('click', function(ev) { ev.stopPropagation(); });
        }
        document.querySelectorAll('.day-btn').forEach(function(btn) {
            btn.addEventListener('click', function() {
                var day = btn.getAttribute('data-day');
                if (day) {
                    showForm(day);
                    closeMenu();
                }
            });
        });
        /* Instructor, Subject, Course/Student always visible (even for VACANT/BREAK) */
        if (formCode && formYearRow) {
            formCode.addEventListener('change', function() {
                if (formCode.value) {
                    formYearRow.style.display = 'block';
                    if (formYear) formYear.value = '';
                    if (formSectionRow) formSectionRow.style.display = 'none';
                    if (formSection) formSection.value = '';
                    if (formStudentCountRow) formStudentCountRow.style.display = 'none';
                    if (formStudentCount) formStudentCount.value = '';
                } else {
                    formYearRow.style.display = 'none';
                    if (formSectionRow) formSectionRow.style.display = 'none';
                    if (formStudentCountRow) formStudentCountRow.style.display = 'none';
                }
            });
        }
        if (formYear && formSectionRow) {
            formYear.addEventListener('change', function() {
                if (formYear.value) {
                    formSectionRow.style.display = 'block';
                    if (formSection) formSection.value = '';
                    if (formStudentCountRow) formStudentCountRow.style.display = 'none';
                    if (formStudentCount) formStudentCount.value = '';
                } else {
                    formSectionRow.style.display = 'none';
                    if (formStudentCountRow) formStudentCountRow.style.display = 'none';
                }
            });
        }
        if (formSection && formStudentCountRow) {
            formSection.addEventListener('change', function() {
                if (formSection.value) {
                    formStudentCountRow.style.display = 'block';
                    if (formStudentCount) formStudentCount.value = '';
                } else {
                    formStudentCountRow.style.display = 'none';
                }
            });
        }
        var formSaveBtn = document.getElementById('formSaveBtn');
        if (formSaveBtn) formSaveBtn.addEventListener('click', onFormSave);
        var formCancelBtn = document.getElementById('formCancelBtn');
        if (formCancelBtn) formCancelBtn.addEventListener('click', hideForm);
    }
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();

/* ==========================================================
   Galgotias University Exam Seat Finder
   ----------------------------------------------------------
   Privacy rules this file follows:
   - The spreadsheet is read with SheetJS inside the browser.
   - No fetch / XMLHttpRequest / beacon calls exist in this file.
   - Nothing is written to localStorage, sessionStorage, cookies or IndexedDB.
   - Student data is never printed to the console.
   - All data lives in the in-memory `appState` object and is
     discarded on "Remove file", page reload or tab close.
   - Spreadsheet text is inserted with textContent (never innerHTML),
     so a malicious cell cannot inject HTML or scripts.
   ========================================================== */

(function () {
  'use strict';

  /* ---------------- Configuration ---------------- */

  const MAX_FILE_SIZE_BYTES = 40 * 1024 * 1024; // 40 MB
  const ALLOWED_EXTENSIONS = ['xls', 'xlsx'];
  const HEADER_SCAN_ROWS = 30;      // look for the header row in the first 30 rows
  const MAX_COLUMNS = 200;          // guard against corrupt sheets claiming 16,000+ columns
  const MAX_ROWS_PER_SHEET = 200000;
  const MAX_NAME_RESULTS = 30;
  const ALL_SHEETS = '__all__';

  /*
   * Fields we try to detect. Patterns run against a normalised header
   * (lower case, punctuation turned into spaces), e.g. "Admission No." -> "admission no".
   * Order matters: earlier fields claim a column first.
   */
  const FIELD_DEFINITIONS = [
    { key: 'admission', label: 'Admission number', required: true,
      patterns: [/\badmission\b/, /\badmn?\b/] },
    { key: 'name', label: 'Student name', required: true,
      patterns: [/^(student|candidate)?\s*name$/, /\b(student|candidate)s?\s*name\b/, /^name of (the )?(student|candidate)$/, /^full name$/],
      exclude: /father|mother|guardian|program|course|subject|paper|school|department|college/ },
    { key: 'seat', label: 'Seat number', required: true,
      patterns: [/\bseat\b/], exclude: /seating|plan/ },
    { key: 'enrollment', label: 'Enrollment / roll number',
      patterns: [/\benrol+ment\b/, /\benrol+\b/, /\broll\b/, /\bregistration\b/, /\breg\b/] },
    { key: 'room', label: 'Room',
      patterns: [/\broom\b/, /\bhall\b/, /\bvenue\b/, /\bclassroom\b/] },
    { key: 'building', label: 'Building', patterns: [/\bbuilding\b/, /\bbldg\b/] },
    { key: 'block', label: 'Block', patterns: [/\bblock\b/, /\bwing\b/] },
    { key: 'floor', label: 'Floor', patterns: [/\bfloor\b/] },
    { key: 'subject', label: 'Subject code',
      patterns: [/\bsub(ject)?\s*code\b/, /\bpaper\s*code\b/, /\bcourse\s*code\b/, /\bsubject\b/] },
    { key: 'program', label: 'Program',
      patterns: [/\bprogram(me)?\b/, /\bcourse\b/, /\bbranch\b/, /\bdepartment\b/] },
    { key: 'semester', label: 'Semester', patterns: [/^sem\b/, /\bsemester\b/] },
    { key: 'section', label: 'Section', patterns: [/^sec\b/, /\bsection\b/] },
    { key: 'date', label: 'Exam date', patterns: [/\bdate\b/] },
    { key: 'shift', label: 'Shift / time', patterns: [/\bshift\b/, /\btimings?\b/, /\bsession\b/, /\btime\b/] }
  ];

  // Order in which optional details appear on a result card.
  const DETAIL_FIELD_ORDER = ['name', 'admission', 'enrollment', 'program', 'semester', 'section',
    'subject', 'building', 'block', 'floor', 'date', 'shift'];

  /* ---------------- Text helpers ---------------- */

  function cleanText(value) {
    if (value === null || value === undefined) return '';
    return String(value).replace(/\s+/g, ' ').trim();
  }

  function normalizeHeader(value) {
    return cleanText(value).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  }

  // Admission numbers: ignore case and all spaces ("24scse 101" === "24SCSE101").
  function normalizeId(value) {
    return cleanText(value).toUpperCase().replace(/\s+/g, '');
  }

  // Names: ignore case, dots, commas and repeated spaces ("RAHUL  KUMAR." === "rahul kumar").
  function normalizeName(value) {
    return cleanText(value).toLowerCase().replace(/[.,]/g, ' ').replace(/\s+/g, ' ').trim();
  }

  function formatNumber(n) {
    return Number(n).toLocaleString('en-IN');
  }

  function pluralize(count, singular, plural) {
    return formatNumber(count) + ' ' + (count === 1 ? singular : (plural || singular + 's'));
  }

  // Show only the last 4 characters of an ID in lists of other students.
  function maskId(id) {
    const text = cleanText(id);
    if (text.length <= 4) return text;
    return '•'.repeat(Math.min(text.length - 4, 8)) + text.slice(-4);
  }

  /* ---------------- Spreadsheet reading ---------------- */

  /*
   * Convert one SheetJS cell into display text.
   * Large whole numbers (e.g. numeric admission numbers) can be formatted by
   * Excel as "2.41E+11"; in that case the exact integer is used instead.
   */
  function cellToText(cell) {
    if (!cell || cell.v === undefined || cell.v === null) return '';
    if (cell.t === 'e') return ''; // Excel error values such as #N/A
    if (cell.t === 'n') {
      if (Number.isInteger(cell.v) && (!cell.w || /e[+-]/i.test(cell.w))) return String(cell.v);
      return cell.w !== undefined ? cell.w : String(cell.v);
    }
    return cell.w !== undefined ? String(cell.w) : String(cell.v);
  }

  // Read a worksheet into a 2-D array of strings, working with both dense and sparse SheetJS modes.
  function readSheetRows(worksheet) {
    if (!worksheet || !worksheet['!ref']) return { rows: [], startRow: 0, startCol: 0 };

    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const lastCol = Math.min(range.e.c, range.s.c + MAX_COLUMNS - 1);
    const lastRow = Math.min(range.e.r, range.s.r + MAX_ROWS_PER_SHEET - 1);
    const denseData = Array.isArray(worksheet['!data']) ? worksheet['!data'] : null;
    const rows = [];

    for (let r = range.s.r; r <= lastRow; r++) {
      const denseRow = denseData ? denseData[r] : null;
      const row = [];
      for (let c = range.s.c; c <= lastCol; c++) {
        const cell = denseData
          ? (denseRow ? denseRow[c] : undefined)
          : worksheet[XLSX.utils.encode_cell({ r: r, c: c })];
        row.push(cellToText(cell));
      }
      rows.push(row);
    }
    return { rows: rows, startRow: range.s.r, startCol: range.s.c };
  }

  function headerMatchesField(header, field) {
    if (!header || header.length > 60) return false; // long cells are titles, not headers
    if (field.exclude && field.exclude.test(header)) return false;
    return field.patterns.some(function (pattern) { return pattern.test(header); });
  }

  // Map each field to the first column whose header matches it. Each column is used once.
  function autoDetectMapping(headerCells) {
    const mapping = {};
    const usedColumns = new Set();
    const normalized = headerCells.map(normalizeHeader);

    FIELD_DEFINITIONS.forEach(function (field) {
      for (let col = 0; col < normalized.length; col++) {
        if (usedColumns.has(col)) continue;
        if (headerMatchesField(normalized[col], field)) {
          mapping[field.key] = col;
          usedColumns.add(col);
          break;
        }
      }
    });
    return mapping;
  }

  /*
   * University sheets often start with title rows ("GALGOTIAS UNIVERSITY", exam name, date...)
   * before the real header. Score the first rows and pick the one that looks most like a header.
   */
  function detectHeaderRow(rows) {
    let bestIndex = 0;
    let bestScore = 0;
    const limit = Math.min(rows.length, HEADER_SCAN_ROWS);

    for (let i = 0; i < limit; i++) {
      const mapping = autoDetectMapping(rows[i]);
      let score = Object.keys(mapping).length;
      if (mapping.admission !== undefined) score += 2;
      if (mapping.name !== undefined) score += 2;
      if (mapping.seat !== undefined) score += 2;
      if (score > bestScore) {
        bestScore = score;
        bestIndex = i;
      }
    }
    return bestScore >= 3 ? bestIndex : 0;
  }

  /*
   * Lines above the header that describe the exam (date, shift, timings).
   * These are copied word-for-word from the sheet; nothing is invented.
   */
  function extractSessionLines(rows, headerRowIndex) {
    const lines = [];
    for (let i = 0; i < headerRowIndex && i < HEADER_SCAN_ROWS; i++) {
      rows[i].forEach(function (cellText) {
        // Cells like "Date of Exam: 23rd OCT      SHIFT 1 - TIMINGS..." use long runs of spaces as separators.
        String(cellText).split(/\s{10,}/).forEach(function (part) {
          const text = cleanText(part);
          if (text && /date|shift|timing|time|session|exam/i.test(text) && lines.indexOf(text) === -1) {
            lines.push(text);
          }
        });
      });
    }
    return lines.slice(0, 4);
  }

  function getColumnLetter(sheet, columnIndex) {
    return XLSX.utils.encode_col(sheet.startCol + columnIndex);
  }

  function getHeaderCells(sheet) {
    return sheet.rows[sheet.headerRowIndex] || [];
  }

  // Build searchable records for one sheet using its current header row and mapping.
  function buildSheetRecords(sheet) {
    const mapping = sheet.mapping;
    const header = getHeaderCells(sheet);
    const stats = { loaded: 0, malformed: 0, missingSeat: 0, missingAdmission: 0, repeatedHeaders: 0 };
    const records = [];
    const hasIdentity = mapping.admission !== undefined || mapping.name !== undefined;

    if (!hasIdentity) {
      sheet.records = records;
      sheet.stats = stats;
      return;
    }

    const getValue = function (row, key) {
      return mapping[key] === undefined ? '' : cleanText(row[mapping[key]]);
    };

    for (let r = sheet.headerRowIndex + 1; r < sheet.rows.length; r++) {
      const row = sheet.rows[r];
      if (!row || row.every(function (cell) { return cell === ''; })) continue; // blank row

      const admission = getValue(row, 'admission');
      const name = getValue(row, 'name');

      // Some printouts repeat the header row on every page.
      const admissionHeader = mapping.admission !== undefined ? normalizeHeader(header[mapping.admission]) : '';
      const nameHeader = mapping.name !== undefined ? normalizeHeader(header[mapping.name]) : '';
      if ((admission && normalizeHeader(admission) === admissionHeader) ||
          (!admission && name && normalizeHeader(name) === nameHeader)) {
        stats.repeatedHeaders++;
        continue;
      }

      // A row with neither name nor admission number cannot be identified (footer text, totals, etc.).
      if (!admission && !name) {
        stats.malformed++;
        continue;
      }

      const values = {};
      FIELD_DEFINITIONS.forEach(function (field) {
        const value = getValue(row, field.key);
        if (value) values[field.key] = value;
      });

      if (!values.seat) stats.missingSeat++;
      if (!admission && !values.enrollment) stats.missingAdmission++;

      records.push({
        sheetName: sheet.name,
        sessionLines: sheet.sessionLines,
        excelRow: sheet.startRow + r + 1,
        values: values,
        admissionKey: normalizeId(admission),
        enrollmentKey: normalizeId(values.enrollment || ''),
        nameKey: normalizeName(name)
      });
      stats.loaded++;
    }

    sheet.records = records;
    sheet.stats = stats;
  }

  function prepareSheet(sheetName, worksheet) {
    const read = readSheetRows(worksheet);
    const headerRowIndex = detectHeaderRow(read.rows);
    const sheet = {
      name: sheetName,
      rows: read.rows,
      startRow: read.startRow,
      startCol: read.startCol,
      headerRowIndex: headerRowIndex,
      mapping: autoDetectMapping(read.rows[headerRowIndex] || []),
      sessionLines: extractSessionLines(read.rows, headerRowIndex),
      records: [],
      stats: null
    };
    buildSheetRecords(sheet);
    return sheet;
  }

  /* ---------------- Searching ---------------- */

  // Index both admission and enrollment numbers for exact, instant lookups.
  function buildIdIndex(records) {
    const index = new Map();
    const add = function (key, record) {
      if (!key) return;
      if (!index.has(key)) index.set(key, []);
      const list = index.get(key);
      if (list.indexOf(record) === -1) list.push(record);
    };
    records.forEach(function (record) {
      add(record.admissionKey, record);
      add(record.enrollmentKey, record);
    });
    return index;
  }

  function studentKeyOf(record) {
    return record.admissionKey || record.enrollmentKey || ('row:' + record.sheetName + ':' + record.excelRow);
  }

  // Group exam entries by student, since one student can appear in several shifts or subjects.
  function groupByStudent(records) {
    const groups = new Map();
    records.forEach(function (record) {
      const key = studentKeyOf(record);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(record);
    });
    return Array.from(groups.values());
  }

  function searchById(index, query) {
    const key = normalizeId(query);
    return key ? (index.get(key) || []).slice() : [];
  }

  /*
   * Exact full-name match first. If nothing matches exactly, fall back to names
   * that contain every word typed (whole words, or a 3+ letter prefix).
   */
  function searchByName(records, query) {
    const queryKey = normalizeName(query);
    if (!queryKey) return { exact: true, groups: [] };

    const exactMatches = records.filter(function (record) { return record.nameKey === queryKey; });
    if (exactMatches.length) return { exact: true, groups: groupByStudent(exactMatches) };

    const queryWords = queryKey.split(' ');
    const partialMatches = records.filter(function (record) {
      if (!record.nameKey) return false;
      const nameWords = record.nameKey.split(' ');
      return queryWords.every(function (word) {
        return nameWords.some(function (nameWord) {
          return nameWord === word || (word.length >= 3 && nameWord.indexOf(word) === 0);
        });
      });
    });
    return { exact: false, groups: groupByStudent(partialMatches) };
  }

  /* ---------------- Node.js test hook (ignored in browsers) ---------------- */

  if (typeof module === 'object' && module.exports) {
    module.exports = {
      studentKeyOf: studentKeyOf, prepareSheet: prepareSheet, buildSheetRecords: buildSheetRecords, autoDetectMapping: autoDetectMapping,
      buildIdIndex: buildIdIndex, searchById: searchById, searchByName: searchByName,
      groupByStudent: groupByStudent, normalizeId: normalizeId, normalizeName: normalizeName
    };
    return;
  }

  /* ==========================================================
     Browser UI
     ========================================================== */

  const appState = {
    fileName: '',
    sheets: [],
    scope: ALL_SHEETS,
    mappingSheetName: '',
    records: [],
    idIndex: new Map(),
    isLoading: false
  };

  const dom = {};

  function cacheDom() {
    [
      'fileInput', 'dropzone', 'uploadArea', 'loader', 'loaderText', 'dataSummary', 'summaryFile',
      'summaryCounts', 'summaryWarnings', 'sheetSelect', 'removeFileButton', 'mappingPanel',
      'mappingSheetField', 'mappingSheetSelect', 'headerRowInput', 'mappingGrid', 'mappingError',
      'applyMappingButton', 'resetMappingButton', 'statusMessage', 'searchForm', 'searchInput',
      'searchLabel', 'searchButton', 'clearButton', 'searchHint', 'searchError', 'results', 'introSection'
    ].forEach(function (id) { dom[id] = document.getElementById(id); });
  }

  // Small helper to build elements safely (text is always set with textContent).
  function createElement(tag, options, children) {
    const element = document.createElement(tag);
    const opts = options || {};
    if (opts.className) element.className = opts.className;
    if (opts.text !== undefined) element.textContent = opts.text;
    if (opts.attrs) {
      Object.keys(opts.attrs).forEach(function (name) { element.setAttribute(name, opts.attrs[name]); });
    }
    (children || []).forEach(function (child) { if (child) element.appendChild(child); });
    return element;
  }

  // Let the browser paint the loader before heavy synchronous parsing starts.
  function waitForPaint() {
    return new Promise(function (resolve) {
      requestAnimationFrame(function () { setTimeout(resolve, 30); });
    });
  }

  function showStatus(type, message) {
    dom.statusMessage.className = 'status is-' + type;
    dom.statusMessage.setAttribute('role', type === 'error' ? 'alert' : 'status');
    dom.statusMessage.textContent = message;
    dom.statusMessage.hidden = false;
  }

  function hideStatus() {
    dom.statusMessage.hidden = true;
    dom.statusMessage.textContent = '';
  }

  function setLoading(isLoading, fileName) {
    appState.isLoading = isLoading;
    dom.uploadArea.classList.toggle('is-loading', isLoading);
    dom.uploadArea.setAttribute('aria-busy', String(isLoading));
    dom.loader.hidden = !isLoading;
    dom.fileInput.disabled = isLoading;
    if (isLoading) dom.loaderText.textContent = 'Reading ' + (fileName || 'spreadsheet') + '…';
  }

  function getFileExtension(fileName) {
    const parts = String(fileName).toLowerCase().split('.');
    return parts.length > 1 ? parts.pop() : '';
  }

  function describeReadError(error) {
    const message = String((error && error.message) || '');
    if (/password|encrypt/i.test(message)) {
      return 'This file is password-protected. Open it in Excel, remove the password, save it, and try again.';
    }
    if (/unsupported|corrupt|zip|cfb|end of data|not a|invalid/i.test(message)) {
      return 'This file could not be read as an Excel workbook. Open it in Excel, use Save As → Excel Workbook (.xlsx), and try again.';
    }
    return 'The file could not be processed. Make sure it is an .xls or .xlsx file that opens normally in Excel.';
  }

  async function handleFile(file) {
    if (!file || appState.isLoading) return;
    hideStatus();

    if (typeof XLSX === 'undefined') {
      showStatus('error', 'The spreadsheet reader did not load. Check your internet connection and reload the page.');
      return;
    }
    if (ALLOWED_EXTENSIONS.indexOf(getFileExtension(file.name)) === -1) {
      showStatus('error', 'Please choose an Excel file ending in .xls or .xlsx.');
      return;
    }
    if (file.size === 0) {
      showStatus('error', 'This file is empty. Download the seating plan again and retry.');
      return;
    }
    if (file.size > MAX_FILE_SIZE_BYTES) {
      showStatus('error', 'This file is larger than 40 MB. Remove unrelated sheets in Excel and try again.');
      return;
    }

    setLoading(true, file.name);
    await waitForPaint();

    let sheets;
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(new Uint8Array(buffer), {
        type: 'array',
        dense: true,          // faster and lighter for large sheets
        cellFormula: false,   // we only need values
        cellHTML: false,
        cellStyles: false
      });
      await waitForPaint();
      sheets = workbook.SheetNames.map(function (name) { return prepareSheet(name, workbook.Sheets[name]); });
    } catch (error) {
      setLoading(false);
      resetData();
      showStatus('error', describeReadError(error));
      return;
    }

    setLoading(false);

    if (!sheets.length || sheets.every(function (sheet) { return sheet.rows.length === 0; })) {
      resetData();
      showStatus('error', 'This workbook has no data in any worksheet.');
      return;
    }

    appState.fileName = file.name;
    appState.sheets = sheets;
    appState.scope = sheets.length > 1 ? ALL_SHEETS : sheets[0].name;
    appState.mappingSheetName = sheets[0].name;

    populateSheetSelects();
    rebuildSearchIndex();
    renderMappingEditor();
    dom.dataSummary.hidden = false;
    dom.introSection.hidden = true;
    enableSearch(true);
    clearResults();
    announceLoadResult();
  }

  function resetData() {
    appState.fileName = '';
    appState.sheets = [];
    appState.records = [];
    appState.idIndex = new Map();
    appState.scope = ALL_SHEETS;
    appState.mappingSheetName = '';
    dom.fileInput.value = '';
    dom.dataSummary.hidden = true;
    dom.introSection.hidden = false;
    dom.mappingPanel.open = false;
    enableSearch(false);
    clearResults();
  }

  function sheetsInScope() {
    if (appState.scope === ALL_SHEETS) return appState.sheets;
    return appState.sheets.filter(function (sheet) { return sheet.name === appState.scope; });
  }

  function sheetHasRequiredColumns(sheet) {
    return (sheet.mapping.admission !== undefined || sheet.mapping.name !== undefined) &&
      sheet.mapping.seat !== undefined;
  }

  function rebuildSearchIndex() {
    appState.records = [];
    sheetsInScope().forEach(function (sheet) {
      // Avoid spread on large arrays (stack limits); push in a loop instead.
      for (let i = 0; i < sheet.records.length; i++) appState.records.push(sheet.records[i]);
    });
    appState.idIndex = buildIdIndex(appState.records);
    renderSummary();
  }

  function renderSummary() {
    const scoped = sheetsInScope();
    const studentKeys = new Set(appState.records.map(studentKeyOf));

    dom.summaryFile.textContent = appState.fileName;
    dom.summaryCounts.textContent =
      'Loaded ' + pluralize(appState.records.length, 'exam entry', 'exam entries') +
      ' for ' + pluralize(studentKeys.size, 'student') +
      ' from ' + pluralize(scoped.length, 'worksheet') + '.';

    dom.summaryWarnings.replaceChildren();
    const addWarning = function (text) {
      dom.summaryWarnings.appendChild(createElement('li', { text: text }));
    };

    let missingSeat = 0, malformed = 0, missingAdmission = 0;
    scoped.forEach(function (sheet) {
      if (!sheetHasRequiredColumns(sheet)) {
        addWarning('"' + sheet.name + '": the name/admission or seat column was not found. Use "Adjust columns" to choose them.');
      }
      missingSeat += sheet.stats.missingSeat;
      malformed += sheet.stats.malformed;
      missingAdmission += sheet.stats.missingAdmission;
    });
    if (missingSeat) addWarning(pluralize(missingSeat, 'entry has', 'entries have') + ' no seat number in the spreadsheet.');
    if (missingAdmission) addWarning(pluralize(missingAdmission, 'entry has', 'entries have') + ' no admission or roll number and can only be found by name.');
    if (malformed) addWarning(pluralize(malformed, 'row was', 'rows were') + ' skipped because they had neither a name nor an admission number.');
  }

  function announceLoadResult() {
    const scoped = sheetsInScope();
    if (!appState.records.length) {
      dom.mappingPanel.open = true;
      showStatus('warning', 'The file opened, but no student rows were recognised. Choose the correct columns under "Adjust columns".');
      return;
    }
    if (scoped.some(function (sheet) { return !sheetHasRequiredColumns(sheet); })) {
      dom.mappingPanel.open = true;
      showStatus('warning', 'File loaded, but some columns need checking. See the notes below the file name.');
      return;
    }
    showStatus('success', 'Seating plan ready. ' + pluralize(appState.records.length, 'exam entry', 'exam entries') + ' loaded.');
  }

  function populateSheetSelects() {
    dom.sheetSelect.replaceChildren();
    dom.mappingSheetSelect.replaceChildren();

    if (appState.sheets.length > 1) {
      const total = appState.sheets.reduce(function (sum, sheet) { return sum + sheet.records.length; }, 0);
      dom.sheetSelect.appendChild(createElement('option', {
        text: 'All worksheets (' + formatNumber(total) + ' entries)', attrs: { value: ALL_SHEETS }
      }));
    }
    appState.sheets.forEach(function (sheet) {
      const label = sheet.name + ' (' + formatNumber(sheet.records.length) + ' entries)';
      dom.sheetSelect.appendChild(createElement('option', { text: label, attrs: { value: sheet.name } }));
      dom.mappingSheetSelect.appendChild(createElement('option', { text: sheet.name, attrs: { value: sheet.name } }));
    });

    dom.sheetSelect.value = appState.scope;
    dom.mappingSheetSelect.value = appState.mappingSheetName;
    dom.mappingSheetField.hidden = appState.sheets.length < 2;
  }

  function refreshSheetOptionCounts() {
    // Keep the entry counts in the worksheet dropdown accurate after a mapping change.
    const currentScope = appState.scope;
    populateSheetSelects();
    dom.sheetSelect.value = currentScope;
  }

  /* ---------------- Column mapping editor ---------------- */

  function getMappingSheet() {
    return appState.sheets.find(function (sheet) { return sheet.name === appState.mappingSheetName; });
  }

  // Draft mapping shown in the editor; applied only when the user clicks "Apply columns".
  let draftMapping = {};
  let draftHeaderRowIndex = 0;

  function renderMappingEditor(useDraft) {
    const sheet = getMappingSheet();
    if (!sheet) return;

    if (!useDraft) {
      draftMapping = Object.assign({}, sheet.mapping);
      draftHeaderRowIndex = sheet.headerRowIndex;
    }

    dom.headerRowInput.value = String(sheet.startRow + draftHeaderRowIndex + 1);
    dom.headerRowInput.min = String(sheet.startRow + 1);
    dom.headerRowInput.max = String(sheet.startRow + Math.max(sheet.rows.length, 1));
    dom.mappingError.hidden = true;
    dom.mappingGrid.replaceChildren();

    const headerCells = sheet.rows[draftHeaderRowIndex] || [];
    const columnCount = sheet.rows.reduce(function (max, row) { return Math.max(max, row.length); }, 0);

    FIELD_DEFINITIONS.forEach(function (field) {
      const selectId = 'map-' + field.key;
      const label = createElement('label', { className: 'field-label', attrs: { for: selectId } });
      label.textContent = field.label + ' ';
      if (field.required) label.appendChild(createElement('span', { className: 'required-mark', text: '*', attrs: { 'aria-hidden': 'true' } }));

      const select = createElement('select', { className: 'select', attrs: { id: selectId, 'data-field': field.key } });
      select.appendChild(createElement('option', { text: 'Not in this sheet', attrs: { value: '' } }));
      for (let col = 0; col < columnCount; col++) {
        const headerText = cleanText(headerCells[col]);
        const optionText = 'Column ' + getColumnLetter(sheet, col) + (headerText ? ': ' + headerText.slice(0, 40) : '');
        select.appendChild(createElement('option', { text: optionText, attrs: { value: String(col) } }));
      }
      select.value = draftMapping[field.key] === undefined ? '' : String(draftMapping[field.key]);
      select.addEventListener('change', function () {
        if (select.value === '') delete draftMapping[field.key];
        else draftMapping[field.key] = Number(select.value);
      });

      dom.mappingGrid.appendChild(createElement('div', { className: 'field' }, [label, select]));
    });
  }

  function handleHeaderRowChange() {
    const sheet = getMappingSheet();
    if (!sheet) return;
    const excelRow = parseInt(dom.headerRowInput.value, 10);
    const index = excelRow - sheet.startRow - 1;
    if (!Number.isInteger(index) || index < 0 || index >= sheet.rows.length) {
      dom.mappingError.textContent = 'Enter a row number between ' + (sheet.startRow + 1) + ' and ' + (sheet.startRow + sheet.rows.length) + '.';
      dom.mappingError.hidden = false;
      return;
    }
    draftHeaderRowIndex = index;
    draftMapping = autoDetectMapping(sheet.rows[index]);
    renderMappingEditor(true);
  }

  function applyMapping() {
    const sheet = getMappingSheet();
    if (!sheet) return;

    if (draftMapping.admission === undefined && draftMapping.name === undefined) {
      dom.mappingError.textContent = 'Choose at least the admission number or the student name column.';
      dom.mappingError.hidden = false;
      return;
    }
    const usedColumns = Object.keys(draftMapping).map(function (key) { return draftMapping[key]; });
    if (new Set(usedColumns).size !== usedColumns.length) {
      dom.mappingError.textContent = 'Each column can be used for only one field.';
      dom.mappingError.hidden = false;
      return;
    }

    sheet.headerRowIndex = draftHeaderRowIndex;
    sheet.mapping = Object.assign({}, draftMapping);
    sheet.sessionLines = extractSessionLines(sheet.rows, sheet.headerRowIndex);
    buildSheetRecords(sheet);

    refreshSheetOptionCounts();
    rebuildSearchIndex();
    clearResults();
    renderMappingEditor();

    if (!sheetHasRequiredColumns(sheet)) {
      showStatus('warning', 'Columns updated for "' + sheet.name + '", but no seat column is selected. Seats will show as not listed.');
    } else {
      showStatus('success', 'Columns updated for "' + sheet.name + '". ' + pluralize(sheet.records.length, 'entry', 'entries') + ' loaded from this worksheet.');
    }
  }

  function resetMappingToAutomatic() {
    const sheet = getMappingSheet();
    if (!sheet) return;
    draftHeaderRowIndex = detectHeaderRow(sheet.rows);
    draftMapping = autoDetectMapping(sheet.rows[draftHeaderRowIndex] || []);
    renderMappingEditor(true);
    applyMapping();
  }

  /* ---------------- Search UI ---------------- */

  function getSearchMode() {
    const checked = document.querySelector('input[name="searchMode"]:checked');
    return checked ? checked.value : 'admission';
  }

  function updateSearchModeUi() {
    const isName = getSearchMode() === 'name';
    dom.searchLabel.textContent = isName ? 'Full name' : 'Admission number';
    dom.searchInput.placeholder = isName ? 'Enter your full name as on your ID card' : 'Enter your admission number';
    dom.searchInput.setAttribute('autocapitalize', isName ? 'words' : 'characters');
    if (!dom.searchInput.disabled) {
      dom.searchHint.textContent = isName
        ? 'Names are matched without regard to capital letters or extra spaces.'
        : 'Admission number gives the most accurate result. Enrollment numbers also work if the sheet has them.';
    }
    clearSearchError();
  }

  function enableSearch(enabled) {
    dom.searchInput.disabled = !enabled;
    dom.searchButton.disabled = !enabled;
    dom.clearButton.disabled = !enabled;
    if (!enabled) {
      dom.searchInput.value = '';
      dom.searchHint.textContent = 'Open the seating plan above to enable search.';
    } else {
      updateSearchModeUi();
    }
  }

  function showSearchError(message) {
    dom.searchError.textContent = message;
    dom.searchError.hidden = false;
    dom.searchInput.setAttribute('aria-invalid', 'true');
  }

  function clearSearchError() {
    dom.searchError.hidden = true;
    dom.searchError.textContent = '';
    dom.searchInput.removeAttribute('aria-invalid');
  }

  function clearResults() {
    dom.results.replaceChildren();
  }

  function resetSearch() {
    dom.searchInput.value = '';
    clearSearchError();
    clearResults();
    dom.searchInput.focus();
  }

  function handleSearch(event) {
    event.preventDefault();
    clearSearchError();

    if (!appState.records.length) {
      showSearchError('Open the seating plan file first.');
      return;
    }
    const query = cleanText(dom.searchInput.value);
    dom.searchInput.value = query;
    if (!query) {
      showSearchError(getSearchMode() === 'name' ? 'Enter your full name to search.' : 'Enter your admission number to search.');
      dom.searchInput.focus();
      return;
    }
    if (getSearchMode() === 'name' && query.length < 2) {
      showSearchError('Enter at least 2 letters of your name.');
      return;
    }

    if (getSearchMode() === 'admission') {
      const records = searchById(appState.idIndex, query);
      if (!records.length) renderNoMatch();
      else renderStudentResults(records);
    } else {
      const result = searchByName(appState.records, query);
      if (!result.groups.length) renderNoMatch();
      else if (result.groups.length === 1) renderStudentResults(result.groups[0], result.exact ? '' : 'No exact match for that full name. Showing the closest name found.');
      else renderDuplicateChooser(query, result);
    }
  }

  function renderNoMatch() {
    clearResults();
    const heading = createElement('h3', { text: 'No matching student found. Please check your details.', attrs: { tabindex: '-1' } });
    const scopeHint = appState.scope !== ALL_SHEETS && appState.sheets.length > 1
      ? ' You are searching only "' + appState.scope + '". Try "All worksheets".'
      : '';
    const tip = createElement('p', {
      text: 'Check the spelling or try your admission number instead.' + scopeHint
    });
    dom.results.appendChild(createElement('div', { className: 'empty-result' }, [heading, tip]));
    heading.focus();
  }

  function renderResultsHeading(title, subtitle) {
    const heading = createElement('h3', { className: 'results-heading', text: title, attrs: { tabindex: '-1' } });
    const wrapper = createElement('div', {}, [heading]);
    if (subtitle) wrapper.appendChild(createElement('p', { className: 'results-sub', text: subtitle }));
    dom.results.appendChild(wrapper);
    return heading;
  }

  // One student, possibly with several exam entries (different shifts or subjects).
  function renderStudentResults(records, note) {
    clearResults();
    const first = records[0];
    const displayName = first.values.name || 'Name not listed';
    const subtitle = first.values.admission
      ? 'Admission number ' + first.values.admission
      : (first.values.enrollment ? 'Enrollment / roll number ' + first.values.enrollment : '');
    const heading = renderResultsHeading(displayName, subtitle);

    if (note) dom.results.appendChild(createElement('p', { className: 'result-note is-warning', text: note }));

    if (records.length > 1) {
      const sheetCounts = {};
      records.forEach(function (record) { sheetCounts[record.sheetName] = (sheetCounts[record.sheetName] || 0) + 1; });
      const repeatedInSameSheet = Object.keys(sheetCounts).some(function (name) { return sheetCounts[name] > 1; });
      let message = 'You have ' + records.length + ' exam entries in this seating plan.';
      if (repeatedInSameSheet) message += ' Some are in the same worksheet, so check the subject code on each.';
      dom.results.appendChild(createElement('p', { className: 'result-note', text: message }));
    }

    records.forEach(function (record) { dom.results.appendChild(buildTicket(record)); });

    const resetRow = createElement('div', { className: 'button-row' }, [
      createElement('button', { className: 'button button-ghost', text: 'Search again', attrs: { type: 'button' } })
    ]);
    resetRow.firstChild.addEventListener('click', resetSearch);
    dom.results.appendChild(resetRow);
    heading.focus();
  }

  function buildTicket(record) {
    const values = record.values;
    const fieldLabel = function (key) {
      return FIELD_DEFINITIONS.find(function (field) { return field.key === key; }).label;
    };

    // Stub: room and seat, shown large.
    const stub = createElement('div', { className: 'ticket-stub' });
    if (values.room) {
      stub.appendChild(createElement('div', { className: 'stub-item' }, [
        createElement('span', { className: 'stub-label', text: 'Room' }),
        createElement('span', { className: 'stub-value', text: values.room })
      ]));
    }
    stub.appendChild(createElement('div', { className: 'stub-item is-seat' }, [
      createElement('span', { className: 'stub-label', text: 'Seat' }),
      values.seat
        ? createElement('span', { className: 'stub-value', text: values.seat })
        : createElement('span', { className: 'stub-value is-missing', text: 'Not listed in the spreadsheet' })
    ]));

    // Body: session header lines + every detail that exists in the sheet.
    const sessionBlock = createElement('div', {}, [
      createElement('h4', { className: 'ticket-session', text: record.sheetName })
    ]);
    if (record.sessionLines.length) {
      const list = createElement('ul', { className: 'session-lines' });
      record.sessionLines.forEach(function (line) { list.appendChild(createElement('li', { text: line })); });
      sessionBlock.appendChild(list);
    }

    const details = createElement('dl', { className: 'detail-list' });
    DETAIL_FIELD_ORDER.forEach(function (key) {
      if (!values[key]) return;
      details.appendChild(createElement('div', {}, [
        createElement('dt', { text: fieldLabel(key) }),
        createElement('dd', { text: values[key] })
      ]));
    });

    const source = createElement('p', {
      className: 'ticket-source',
      text: 'From worksheet "' + record.sheetName + '", row ' + record.excelRow + '. Confirm with the official notice.'
    });

    const body = createElement('div', { className: 'ticket-body' }, [sessionBlock, details, source]);
    return createElement('article', {
      className: 'ticket',
      attrs: { 'aria-label': 'Seat for ' + record.sheetName }
    }, [stub, body]);
  }

  /*
   * Several different students share the name. Show who they are (with masked IDs)
   * and ask for the full admission number to reveal the right seat.
   */
  function renderDuplicateChooser(query, result) {
    clearResults();
    const groups = result.groups;
    const shownGroups = groups.slice(0, MAX_NAME_RESULTS);
    const title = result.exact
      ? pluralize(groups.length, 'student') + ' named "' + query + '"'
      : pluralize(groups.length, 'student') + ' with names containing "' + query + '"';
    const heading = renderResultsHeading(title, 'Enter your admission number to see your seat.');

    if (groups.length > MAX_NAME_RESULTS) {
      dom.results.appendChild(createElement('p', {
        className: 'result-note is-warning',
        text: 'Showing the first ' + MAX_NAME_RESULTS + '. Type your full name or use your admission number to narrow it down.'
      }));
    }

    const list = createElement('ul', { className: 'candidate-list', attrs: { 'aria-label': 'Matching students' } });
    shownGroups.forEach(function (group) {
      const v = group[0].values;
      const meta = [v.program, v.semester, v.section].filter(Boolean).join(', ');
      list.appendChild(createElement('li', {}, [
        createElement('span', { className: 'candidate-name', text: v.name || 'Name not listed' }),
        createElement('span', { className: 'candidate-meta', text: meta || 'No program details in sheet' }),
        createElement('span', {
          className: 'candidate-id',
          text: (v.admission || v.enrollment) ? 'ID ending ' + maskId(v.admission || v.enrollment) : 'No ID in sheet',
          attrs: (v.admission || v.enrollment) ? { 'aria-label': 'ID ending in ' + (v.admission || v.enrollment).slice(-4) } : {}
        })
      ]));
    });
    dom.results.appendChild(list);

    // Confirmation by admission number (or enrollment number) within these matches only.
    const inputId = 'confirmAdmissionInput';
    const errorId = 'confirmAdmissionError';
    const input = createElement('input', {
      className: 'text-input',
      attrs: { id: inputId, type: 'text', autocomplete: 'off', spellcheck: 'false', autocapitalize: 'characters', 'aria-describedby': errorId }
    });
    const error = createElement('p', { className: 'field-error', attrs: { id: errorId }, text: '' });
    error.hidden = true;
    const form = createElement('form', { className: 'confirm-form', attrs: { novalidate: '' } }, [
      createElement('div', { className: 'field' }, [
        createElement('label', { className: 'field-label', text: 'Your admission number', attrs: { for: inputId } }),
        input
      ]),
      createElement('button', { className: 'button button-primary', text: 'Show my seat', attrs: { type: 'submit' } }),
      createElement('button', { className: 'button button-ghost', text: 'Clear search', attrs: { type: 'button' } })
    ]);
    form.querySelector('button[type="button"]').addEventListener('click', resetSearch);

    form.addEventListener('submit', function (event) {
      event.preventDefault();
      const key = normalizeId(input.value);
      if (!key) {
        error.textContent = 'Enter your admission number.';
        error.hidden = false;
        input.setAttribute('aria-invalid', 'true');
        input.focus();
        return;
      }
      const match = groups.find(function (group) {
        return group.some(function (record) { return record.admissionKey === key || record.enrollmentKey === key; });
      });
      if (!match) {
        error.textContent = 'That admission number does not belong to any of these students. Check it and try again.';
        error.hidden = false;
        input.setAttribute('aria-invalid', 'true');
        return;
      }
      // Show every entry for that admission number across the selected worksheets.
      const matchKey = match[0].admissionKey || match[0].enrollmentKey;
      const allEntries = matchKey ? searchById(appState.idIndex, matchKey) : match;
      renderStudentResults(allEntries.filter(function (record) { return studentKeyOf(record) === studentKeyOf(match[0]); }));
    });

    dom.results.appendChild(error);
    dom.results.insertBefore(form, error);
    heading.focus();
  }

  /* ---------------- Event wiring ---------------- */

  function bindEvents() {
    dom.fileInput.addEventListener('change', function () {
      handleFile(dom.fileInput.files && dom.fileInput.files[0]);
    });

    // Drag and drop onto the dropzone.
    ['dragenter', 'dragover'].forEach(function (type) {
      dom.dropzone.addEventListener(type, function (event) {
        event.preventDefault();
        dom.dropzone.classList.add('is-dragging');
      });
    });
    ['dragleave', 'drop'].forEach(function (type) {
      dom.dropzone.addEventListener(type, function (event) {
        event.preventDefault();
        dom.dropzone.classList.remove('is-dragging');
      });
    });
    dom.dropzone.addEventListener('drop', function (event) {
      const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
      handleFile(file);
    });
    // Stop the browser from opening a file dropped anywhere else on the page.
    window.addEventListener('dragover', function (event) { event.preventDefault(); });
    window.addEventListener('drop', function (event) { event.preventDefault(); });

    dom.removeFileButton.addEventListener('click', function () {
      resetData();
      showStatus('success', 'File removed. No student data remains on this page.');
      dom.fileInput.focus();
    });

    dom.sheetSelect.addEventListener('change', function () {
      appState.scope = dom.sheetSelect.value;
      if (appState.scope !== ALL_SHEETS) {
        appState.mappingSheetName = appState.scope;
        dom.mappingSheetSelect.value = appState.scope;
        renderMappingEditor();
      }
      rebuildSearchIndex();
      clearResults();
    });

    dom.mappingSheetSelect.addEventListener('change', function () {
      appState.mappingSheetName = dom.mappingSheetSelect.value;
      renderMappingEditor();
    });
    dom.headerRowInput.addEventListener('change', handleHeaderRowChange);
    dom.applyMappingButton.addEventListener('click', applyMapping);
    dom.resetMappingButton.addEventListener('click', resetMappingToAutomatic);

    document.querySelectorAll('input[name="searchMode"]').forEach(function (radio) {
      radio.addEventListener('change', function () {
        updateSearchModeUi();
        clearResults();
        if (!dom.searchInput.disabled) dom.searchInput.focus();
      });
    });

    dom.searchForm.addEventListener('submit', handleSearch);
    dom.clearButton.addEventListener('click', resetSearch);
    dom.searchInput.addEventListener('input', clearSearchError);
    dom.searchInput.addEventListener('keydown', function (event) {
      if (event.key === 'Escape') resetSearch();
    });
  }

  function init() {
    cacheDom();
    bindEvents();
    updateSearchModeUi();
    enableSearch(false);
    if (typeof XLSX === 'undefined') {
      showStatus('error', 'The spreadsheet reader did not load. Check your internet connection and reload the page.');
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();

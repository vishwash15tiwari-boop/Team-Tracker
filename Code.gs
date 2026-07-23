/**
 * Onboarding Team — Task Tracker
 * Google Apps Script Web App
 *
 * Sheet ID : 1nc7qCA-VsAO3_kW76PAwwdsyWWVf-0dtYdZfxeKYbtU
 * Sheet tab : "Task Tracker"
 *
 * Columns (A–L)
 *   A  #                   Auto-numbered
 *   B  Rank                1 | 2 | 3 | 4
 *   C  Owning Function
 *   D  Task
 *   E  Primary Owner
 *   F  Secondary Owner
 *   G  Start Date
 *   H  Due Date
 *   I  End Date
 *   J  TAT (Days)          Auto-calculated server-side
 *   K  Task Brief / Details
 *   L  Volume
 */

const SS_ID   = '1nc7qCA-VsAO3_kW76PAwwdsyWWVf-0dtYdZfxeKYbtU';
const SH_NAME = 'Task Tracker';
const HEADERS = [
  '#', 'Rank', 'Owning Function', 'Task',
  'Primary Owner', 'Secondary Owner',
  'Start Date', 'Due Date', 'End Date',
  'TAT (Days)', 'Task Brief / Details', 'Volume'
];

// ─── Entry point ─────────────────────────────────────────────────────────────

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Onboarding Team — Task Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── Sheet bootstrap ─────────────────────────────────────────────────────────

function getSheet() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sh = ss.getSheetByName(SH_NAME);
  if (!sh) sh = _buildSheet(ss);
  return sh;
}

function _buildSheet(ss) {
  const sh = ss.insertSheet(SH_NAME, 0);

  // Header row
  const hdrRange = sh.getRange(1, 1, 1, HEADERS.length);
  hdrRange.setValues([HEADERS]);
  hdrRange.setBackground('#16213e');
  hdrRange.setFontColor('#ffffff');
  hdrRange.setFontWeight('bold');
  hdrRange.setFontFamily('Arial');
  hdrRange.setFontSize(10);
  hdrRange.setHorizontalAlignment('center');
  hdrRange.setVerticalAlignment('middle');
  hdrRange.setWrap(true);
  sh.setRowHeight(1, 48);

  // Column widths
  const widths = [45, 70, 165, 235, 130, 130, 108, 108, 108, 90, 290, 85];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  sh.setFrozenRows(1);
  return sh;
}

// ─── Public API ───────────────────────────────────────────────────────────────

/** Return all task rows as JSON-safe objects. */
function getTasks() {
  try {
    const sh = getSheet();
    const lr = sh.getLastRow();
    if (lr < 2) return [];

    const tz   = Session.getScriptTimeZone();
    const vals = sh.getRange(2, 1, lr - 1, HEADERS.length).getValues();

    return vals
      .filter(r => r[3] || r[2])   // skip completely blank rows
      .map((r, i) => ({
        rowIndex : i + 2,
        num      : r[0],
        rank     : String(r[1]),
        owFn     : r[2],
        task     : r[3],
        priOwner : r[4],
        secOwner : r[5],
        startDate: _fmtDate(r[6], tz),
        dueDate  : _fmtDate(r[7], tz),
        endDate  : _fmtDate(r[8], tz),
        tat      : r[9],
        brief    : r[10],
        volume   : r[11]
      }));

  } catch (e) {
    return { error: e.toString() };
  }
}

/** Insert or update a task row. */
function saveTask(t) {
  try {
    const sh  = getSheet();
    const tat = _calcTAT(t.startDate, t.endDate);

    const row = [
      t.num      != null ? t.num : '',
      t.rank       || '',
      t.owFn       || '',
      t.task       || '',
      t.priOwner   || '',
      t.secOwner   || '',
      t.startDate  ? new Date(t.startDate) : '',
      t.dueDate    ? new Date(t.dueDate)   : '',
      t.endDate    ? new Date(t.endDate)   : '',
      tat,
      t.brief      || '',
      t.volume     || ''
    ];

    let targetRow;

    if (t.rowIndex && t.rowIndex >= 2) {
      // ── Update existing ─────────────────────────────────────────
      sh.getRange(t.rowIndex, 1, 1, HEADERS.length).setValues([row]);
      targetRow = t.rowIndex;
    } else {
      // ── Append new row ──────────────────────────────────────────
      const newNum = Math.max(sh.getLastRow(), 1);  // 1-based task #
      row[0] = newNum;
      sh.appendRow(row);
      targetRow = sh.getLastRow();
    }

    // Date number-format
    ['G', 'H', 'I'].forEach(col => {
      const c = sh.getRange(col + targetRow);
      if (c.getValue()) c.setNumberFormat('dd-mmm-yyyy');
    });

    _styleDataRows(sh);
    return { success: true };

  } catch (e) {
    return { error: e.toString() };
  }
}

/** Delete a row by its 1-based sheet row index and renumber. */
function deleteTask(rowIndex) {
  try {
    const sh = getSheet();
    sh.deleteRow(rowIndex);
    _renumber(sh);
    _styleDataRows(sh);
    return { success: true };
  } catch (e) {
    return { error: e.toString() };
  }
}

// ─── Private helpers ──────────────────────────────────────────────────────────

function _fmtDate(val, tz) {
  if (!val || val === '') return '';
  try {
    return Utilities.formatDate(new Date(val), tz, 'yyyy-MM-dd');
  } catch (_) {
    return String(val);
  }
}

/** Days between start and end (or today if end is blank). */
function _calcTAT(startStr, endStr) {
  if (!startStr) return '';
  const s    = new Date(startStr);
  const e    = endStr ? new Date(endStr) : new Date();
  const days = Math.ceil((e - s) / 86400000);
  if (days < 0) return 0;
  return endStr ? days : (days + '+ (running)');
}

function _renumber(sh) {
  const lr = sh.getLastRow();
  for (let r = 2; r <= lr; r++) sh.getRange(r, 1).setValue(r - 1);
}

function _styleDataRows(sh) {
  const lr = sh.getLastRow();
  if (lr < 2) return;

  for (let r = 2; r <= lr; r++) {
    const rng = sh.getRange(r, 1, 1, HEADERS.length);
    rng.setBackground(r % 2 === 0 ? '#ffffff' : '#f7f8fc');
    rng.setFontFamily('Arial');
    rng.setFontSize(10);
    rng.setVerticalAlignment('middle');
    sh.setRowHeight(r, 40);
  }

  sh.getRange(2, 1, lr - 1, HEADERS.length)
    .setBorder(
      true, true, true, true, true, true,
      '#dde1ea',
      SpreadsheetApp.BorderStyle.SOLID
    );
}

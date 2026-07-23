/**
 * Onboarding Team — Task Tracker  (Google Apps Script)
 *
 * Team  : Ajay (Manager) | Harshita | Vamsi | Naveen | Vishwas
 * Sheet : 1nc7qCA-VsAO3_kW76PAwwdsyWWVf-0dtYdZfxeKYbtU → "Task Tracker" tab
 *
 * Columns A–N
 *   A  #                  Auto-numbered
 *   B  Status             Not Started | In Progress | Completed | On Hold
 *   C  Owning Function    Marketplace | Open Marketplace | EPR & Sustainability |
 *                         SOPs / MIS / Tracker | Onboarding | Admin |
 *                         Automation | Marketing | SOP | MIS | 3rd Party
 *   D  Task
 *   E  Primary Owner      Ajay | Harshita | Vamsi | Naveen | Vishwas
 *   F  Secondary Owner
 *   G  Start Date
 *   H  End Date
 *   I  TAT (Days)         Auto-calculated: End–Start or Today–Start (running)
 *   J  Task Brief / Details
 *   K  Volume
 *   L  Blocker
 *   M  Priority           High | Medium | Low
 */

const SS_ID   = '1nc7qCA-VsAO3_kW76PAwwdsyWWVf-0dtYdZfxeKYbtU';
const SH_NAME = 'Task Tracker';
const HEADERS = [
  '#', 'Status', 'Owning Function', 'Task',
  'Primary Owner', 'Secondary Owner',
  'Start Date', 'End Date',
  'TAT (Days)', 'Task Brief / Details', 'Volume',
  'Blocker', 'Priority'
];

// ── Entry point ──────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Onboarding Team — Task Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Sheet bootstrap ──────────────────────────────────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sh = ss.getSheetByName(SH_NAME);
  if (!sh) sh = _buildSheet(ss);
  return sh;
}

function _buildSheet(ss) {
  const sh = ss.insertSheet(SH_NAME, 0);

  const hdr = sh.getRange(1, 1, 1, HEADERS.length);
  hdr.setValues([HEADERS]);
  hdr.setBackground('#16213e');
  hdr.setFontColor('#ffffff');
  hdr.setFontWeight('bold');
  hdr.setFontFamily('Arial');
  hdr.setFontSize(10);
  hdr.setHorizontalAlignment('center');
  hdr.setVerticalAlignment('middle');
  hdr.setWrap(true);
  sh.setRowHeight(1, 48);

  //                  #    Status  OwFn  Task  PriOwn SecOwn Start End  TAT  Brief Vol  Blocker  Priority
  const widths = [45, 120, 170,   240,  130,  130,   108,  108,  90,  280,  80,  200,    100];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);
  return sh;
}

// ── Public API ───────────────────────────────────────────────────────────────

function getTasks() {
  try {
    const sh  = getSheet();
    const lr  = sh.getLastRow();
    if (lr < 2) return [];

    const tz   = Session.getScriptTimeZone();
    const vals = sh.getRange(2, 1, lr - 1, HEADERS.length).getValues();

    return vals
      .filter(r => r[3] || r[2])
      .map((r, i) => ({
        rowIndex : i + 2,
        num      : r[0],
        status   : r[1]  || '',
        owFn     : r[2]  || '',
        task     : r[3]  || '',
        priOwner : r[4]  || '',
        secOwner : r[5]  || '',
        startDate: _fmt(r[6], tz),
        endDate  : _fmt(r[7], tz),
        tat      : r[8]  || '',
        brief    : r[9]  || '',
        volume   : r[10] || '',
        blocker  : r[11] || '',
        priority : r[12] || ''
      }));

  } catch (e) {
    return { error: e.toString() };
  }
}

/** Insert (rowIndex falsy) or update (rowIndex ≥ 2) a task row.
 *  Returns { success, rowIndex } so the client can track new rows. */
function saveTask(t) {
  try {
    const sh  = getSheet();
    const tat = _tat(t.startDate, t.endDate);

    const row = [
      t.num != null ? t.num : '',
      t.status    || '',
      t.owFn      || '',
      t.task      || '',
      t.priOwner  || '',
      t.secOwner  || '',
      t.startDate ? new Date(t.startDate) : '',
      t.endDate   ? new Date(t.endDate)   : '',
      tat,
      t.brief     || '',
      t.volume    || '',
      t.blocker   || '',
      t.priority  || ''
    ];

    let targetRow;

    if (t.rowIndex && t.rowIndex >= 2) {
      sh.getRange(t.rowIndex, 1, 1, HEADERS.length).setValues([row]);
      targetRow = t.rowIndex;
    } else {
      row[0] = Math.max(sh.getLastRow(), 1);
      sh.appendRow(row);
      targetRow = sh.getLastRow();
    }

    ['G', 'H'].forEach(col => {
      const c = sh.getRange(col + targetRow);
      if (c.getValue()) c.setNumberFormat('dd-mmm-yyyy');
    });

    _styleRows(sh);
    return { success: true, rowIndex: targetRow };

  } catch (e) {
    return { error: e.toString() };
  }
}

function deleteTask(rowIndex) {
  try {
    const sh = getSheet();
    sh.deleteRow(rowIndex);
    _renumber(sh);
    _styleRows(sh);
    return { success: true };
  } catch (e) {
    return { error: e.toString() };
  }
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function _fmt(val, tz) {
  if (!val || val === '') return '';
  try { return Utilities.formatDate(new Date(val), tz, 'yyyy-MM-dd'); }
  catch (_) { return String(val); }
}

function _tat(startStr, endStr) {
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

function _styleRows(sh) {
  const lr = sh.getLastRow();
  if (lr < 2) return;
  for (let r = 2; r <= lr; r++) {
    const rng = sh.getRange(r, 1, 1, HEADERS.length);
    rng.setBackground(r % 2 === 0 ? '#ffffff' : '#f7f8fc');
    rng.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
    sh.setRowHeight(r, 40);
  }
  sh.getRange(2, 1, lr - 1, HEADERS.length)
    .setBorder(true, true, true, true, true, true,
               '#dde1ea', SpreadsheetApp.BorderStyle.SOLID);
}

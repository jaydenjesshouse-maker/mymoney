/**
 * Jayden & Jess Money — the bit that talks to the spreadsheet.
 *
 * You do NOT set the passcode in here any more.
 * It lives in a tab called "Settings" in the spreadsheet itself, cell B2.
 * Change it there and it works straight away — no redeploying.
 *
 * The line below is only used the very first time, to create that tab.
 */
const FIRST_RUN_PASSCODE = 'mymoney';

/* ------------------------------------------------------------------ */
/* Nothing below here needs touching.                                  */
/* ------------------------------------------------------------------ */

const DATA_SHEET = '_data';
const SNAP_SHEET = 'Snapshot';
const BACKUP_SHEET = '_backup';
const SETTINGS_SHEET = 'Settings';
const CODE_CELL = 'B2';
const CHUNK = 40000;   // a cell holds 50,000 characters; leave headroom
const MAX_ROWS = 200;

function doGet() {
  settingsSheet_();  // make sure the Settings tab exists before anyone knocks
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('My Money')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** The app asks for everything when it opens. */
function loadState(code) {
  if (!ok_(code)) return { badcode: true };
  const raw = readRaw_();
  return { ok: true, rev: getRev_(), state: raw ? JSON.parse(raw) : null };
}

/** Normal save. Refuses if the other person saved first. */
function saveState(code, json, rev, snapshot) {
  if (!ok_(code)) return { badcode: true };
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const current = getRev_();
    if (Number(rev) !== current) {
      const raw = readRaw_();
      return { conflict: true, rev: current, state: raw ? JSON.parse(raw) : null };
    }
    return commit_(json, snapshot, current);
  } finally {
    lock.releaseLock();
  }
}

/** "Put my changes back" — writes over whatever is there. */
function forceSave(code, json, snapshot) {
  if (!ok_(code)) return { badcode: true };
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    return commit_(json, snapshot, getRev_());
  } finally {
    lock.releaseLock();
  }
}

/* ---------------------------- passcode ---------------------------- */

/**
 * Forgiving on purpose: ignores capitals, spaces at either end, and any
 * quote marks someone pastes in by accident. A door code, not a password.
 */
function tidy_(v) {
  return String(v == null ? '' : v)
    .replace(/[‘’“”'"`]/g, '')
    .replace(/\s+/g, '')
    .toLowerCase();
}

function ok_(code) {
  const want = tidy_(currentPasscode_());
  const got = tidy_(code);
  return want !== '' && got === want;
}

function currentPasscode_() {
  const sh = settingsSheet_();
  const v = sh.getRange(CODE_CELL).getValue();
  if (String(v).trim() !== '') return v;
  sh.getRange(CODE_CELL).setValue(FIRST_RUN_PASSCODE);
  return FIRST_RUN_PASSCODE;
}

function settingsSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SETTINGS_SHEET, 0);
    sh.getRange('A1').setValue('Jayden & Jess Money — settings').setFontWeight('bold').setFontSize(13);
    sh.getRange('A2').setValue('Passcode').setFontWeight('bold');
    sh.getRange(CODE_CELL).setValue(FIRST_RUN_PASSCODE);
    sh.getRange('C2').setValue('Type a new code here and it works immediately. Capitals and spaces are ignored.');
    sh.getRange('A4').setValue('Anyone with the link needs this code to see your numbers.');
    sh.setColumnWidth(1, 110);
    sh.setColumnWidth(2, 190);
    sh.setColumnWidth(3, 430);
    sh.getRange(CODE_CELL).setBackground('#E8ECE5').setFontWeight('bold').setFontSize(13);
  }
  return sh;
}

/* ---------------------------- internals --------------------------- */

function commit_(json, snapshot, currentRev) {
  backupOnce_();
  writeRaw_(String(json));
  const next = currentRev + 1;
  setRev_(next);
  try { writeSnapshot_(snapshot); } catch (e) { /* the snapshot is a nicety, never block a save */ }
  return { ok: true, rev: next };
}

/** Keep one copy of whatever was here before the first write of a new layout. */
function backupOnce_() {
  var ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(BACKUP_SHEET)) return;
  var raw = readRaw_();
  if (!raw) return;
  var sh = ss.insertSheet(BACKUP_SHEET);
  sh.getRange('A1').setValue('Backup taken ' + new Date());
  var rows = [], CH = 40000;
  for (var i = 0; i < raw.length; i += CH) rows.push([raw.substr(i, CH)]);
  if (rows.length) sh.getRange(2, 1, rows.length, 1).setValues(rows);
  sh.hideSheet();
}

function dataSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DATA_SHEET);
    sh.getRange('C1').setValue('Do not edit this tab by hand — the app keeps its data here.');
    sh.hideSheet();
  }
  return sh;
}

function readRaw_() {
  const sh = dataSheet_();
  const vals = sh.getRange(1, 1, MAX_ROWS, 1).getValues();
  let out = '';
  for (let i = 0; i < vals.length; i++) {
    const v = vals[i][0];
    if (v === '' || v === null) break;
    out += String(v);
  }
  return out;
}

function writeRaw_(str) {
  const sh = dataSheet_();
  sh.getRange(1, 1, MAX_ROWS, 1).clearContent();
  const rows = [];
  for (let i = 0; i < str.length; i += CHUNK) rows.push([str.substr(i, CHUNK)]);
  if (rows.length > MAX_ROWS) throw new Error('Too much data to store.');
  if (rows.length) sh.getRange(1, 1, rows.length, 1).setValues(rows);
}

function getRev_() {
  const v = dataSheet_().getRange('B1').getValue();
  return Number(v) || 0;
}

function setRev_(n) {
  dataSheet_().getRange('B1').setValue(n);
}

/** A plain-English tab you can read in the Sheets app — pure backup, the app never reads it. */
function writeSnapshot_(rows) {
  if (!rows || !rows.length) return;
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SNAP_SHEET);
  if (!sh) sh = ss.insertSheet(SNAP_SHEET);
  sh.clear();
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sh.setFrozenRows(1);
}

/**
 * School Student Data - Google Sheet backend (receives data from the offline HTML app)
 *
 * 1) Open your Google Sheet > Extensions > Apps Script. Replace ALL code with this file.
 * 2) Run setup() (creates/updates the "Students" and "Teacher PINs" tabs, class tabs and "Summary").
 *    Running it again is safe: it keeps your data and PINs.
 * 3) Deploy > Manage deployments > Edit (pencil) > New version > Deploy.
 *    (First time: Deploy > New deployment > Web app: Execute as = Me, Who has access = Anyone.)
 */

const CLASSES = ['Nursery', 'Prep', 'Class 1', 'Class 2', 'Class 3', 'Class 4', 'Class 5'];
const OFFICE = 'Office (All Classes)';   // special login that can enter/see every class
const HEAD = ['ID', 'Student Name', "Father's Name", "Father's CNIC", 'Student CNIC / Form B', 'Mobile Number',
              'Class', 'Added On', 'Date of Birth', 'Admission Date', 'Age'];
const TAB_HEAD = ['Student Name', "Father's Name", "Father's CNIC", 'Student CNIC / Form B', 'Mobile Number',
                  'Class', 'Date of Birth', 'Admission Date', 'Age'];
const CNIC_RE = /^\d{5}-?\d{7}-?\d$/;
const MOBILE_RE = /^(\+?92|0)3\d{2}-?\d{7}$/;

// ---------- one-time setup ----------
// Safe to run again: keeps existing data and PINs.
function setup() {
  const ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('SS', ss.getId());

  // Main database: all classes together
  let st = ss.getSheetByName('Students');
  if (!st) { st = ss.getSheets()[0]; st.setName('Students'); }
  if (st.getMaxRows() < 5000) st.insertRowsAfter(st.getMaxRows(), 5000 - st.getMaxRows());   // room for 5000 students
  st.getRange('A:G').setNumberFormat('@');
  st.getRange('I:J').setNumberFormat('dd-mmm-yyyy');
  st.getRange(1, 1, 1, HEAD.length).setValues([HEAD])
    .setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  st.setFrozenRows(1);
  st.setColumnWidths(2, 10, 150);
  st.hideColumns(1);
  st.getRange('G2:G5000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(CLASSES, true).setAllowInvalid(false).build());
  // Age formula for every existing student row
  const last = st.getLastRow();
  if (last >= 2) {
    const f = [];
    for (let r = 2; r <= last; r++) f.push([ageFormula_(r)]);
    st.getRange(2, 11, f.length, 1).setFormulas(f);
  }

  // PINs: one per class (adds PINs only for classes that do not have one yet)
  let tp = ss.getSheetByName('Teacher PINs');
  if (!tp) {
    tp = ss.insertSheet('Teacher PINs');
    tp.getRange('A:B').setNumberFormat('@');
    tp.getRange(1, 1, 1, 2).setValues([['Class', 'PIN (give to that class teacher)']]).setFontWeight('bold');
    tp.setColumnWidths(1, 2, 220);
  }
  const have = tp.getDataRange().getValues().slice(1).map(r => String(r[0]));
  const missing = [OFFICE].concat(CLASSES).filter(c => have.indexOf(c) === -1)
    .map(c => [c, String(Math.floor(100000 + Math.random() * 900000))]);
  if (missing.length) tp.getRange(tp.getLastRow() + 1, 1, missing.length, 2).setValues(missing);

  // One automatic tab per class (read-only view, fills itself from Students)
  CLASSES.forEach(c => {
    let t = ss.getSheetByName(c);
    if (!t) t = ss.insertSheet(c);
    t.getRange(1, 1, 1, TAB_HEAD.length).setValues([TAB_HEAD])
      .setFontWeight('bold').setBackground('#2E75B6').setFontColor('#FFFFFF');
    t.setFrozenRows(1);
    t.setColumnWidths(1, TAB_HEAD.length, 150);
    t.getRange('A2').setFormula(
      '=IFERROR(FILTER({Students!B2:G, Students!I2:K}, Students!G2:G="' + c + '"), "")');
    t.getRange('G:H').setNumberFormat('dd-mmm-yyyy');
  });

  // Summary: number of students per class
  let sm = ss.getSheetByName('Summary');
  if (!sm) {
    sm = ss.insertSheet('Summary');
    sm.getRange(1, 1, 1, 2).setValues([['Class', 'Students']]).setFontWeight('bold');
    sm.setColumnWidths(1, 2, 150);
  }
  const srows = CLASSES.map(c => [c, '=COUNTIF(Students!G2:G, "' + c + '")']);
  sm.getRange(2, 1, srows.length, 2).setValues(srows);
  sm.getRange(srows.length + 2, 1, 1, 2).setValues([['Total', '=SUM(B2:B' + (srows.length + 1) + ')']]).setFontWeight('bold');

  SpreadsheetApp.getUi().alert('Setup done. Now publish the change: Deploy > Manage deployments > Edit > New version > Deploy.');
}

// ---------- web endpoints ----------
function doGet() {
  return ContentService.createTextOutput('Student data service is running.');
}

function doPost(e) {
  let out;
  try {
    out = handleSync_(JSON.parse(e.postData.contents));
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// Receives pending records, saves them, returns the list (one class, or all for the Office login).
function handleSync_(req) {
  const cls = String(req.cls || '');
  auth_(cls, req.pin);
  const isAll = (cls === OFFICE);

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sh = sheet_();
    const tz = ss_().getSpreadsheetTimeZone();
    const v = sh.getDataRange().getValues();
    const rowOf = {};
    for (let i = 1; i < v.length; i++) rowOf[String(v[i][0])] = { row: i + 1, cls: String(v[i][6]) };

    const inserts = [], dels = [], done = [], failed = {};
    (req.records || []).forEach(r => {
      const id = String(r.id || '');
      try {
        if (!id) throw new Error('Missing ID');
        const hit = rowOf[id];
        if (!isAll && hit && hit.cls !== cls) throw new Error('Record belongs to another class');
        if (r.del) { if (hit) dels.push(hit.row); done.push(id); return; }
        const tc = isAll ? clean_(r.cls) : cls;          // class the student studies in
        if (CLASSES.indexOf(tc) === -1) throw new Error('Choose the class');

        const name = clean_(r.name), father = clean_(r.father), fcnic = clean_(r.fcnic),
              scnic = clean_(r.scnic), mobile = clean_(r.mobile);
        if (!name || !father) throw new Error('Student name and father name are required');
        if (!CNIC_RE.test(fcnic)) throw new Error("Father's CNIC must be 13 digits");
        if (scnic && !CNIC_RE.test(scnic)) throw new Error('Student CNIC / Form B must be 13 digits');
        if (!MOBILE_RE.test(mobile)) throw new Error('Mobile number is not valid');
        const dob = parseDate_(r.dob);
        if (!dob) throw new Error('Date of birth is required');
        if (dob > new Date()) throw new Error('Date of birth cannot be in the future');
        const adm = parseDate_(r.adm);

        if (hit) {
          sh.getRange(hit.row, 2, 1, 6).setValues([[name, father, fcnic, scnic, mobile, tc]]);
          sh.getRange(hit.row, 9, 1, 3).setValues([[dob, adm, '']]);
          sh.getRange(hit.row, 11).setFormula(ageFormula_(hit.row));
        } else {
          inserts.push([id, name, father, fcnic, scnic, mobile, tc, new Date(), dob, adm]);
        }
        done.push(id);
      } catch (err) {
        failed[id] = String(err.message || err);
      }
    });

    if (inserts.length) {
      const start = sh.getLastRow() + 1;
      const need = start + inserts.length - 1 - sh.getMaxRows();
      if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
      sh.getRange(start, 1, inserts.length, 10).setValues(inserts);
      sh.getRange(start, 11, inserts.length, 1).setFormulas(inserts.map((x, i) => [ageFormula_(start + i)]));
    }
    dels.sort((a, b) => b - a).forEach(rw => sh.deleteRow(rw));

    const list = sh.getDataRange().getValues().slice(1)
      .filter(r => isAll ? String(r[0]) !== '' : String(r[6]) === cls)
      .map(r => ({ id: String(r[0]), name: String(r[1]), father: String(r[2]),
                   fcnic: String(r[3]), scnic: String(r[4]), mobile: String(r[5]), cls: String(r[6]),
                   dob: fmtDate_(r[8], tz), adm: fmtDate_(r[9], tz) }));
    return { ok: true, done: done, failed: failed, list: list };
  } finally {
    lock.releaseLock();
  }
}

// ---------- helpers ----------
function ss_() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SS')); }
function sheet_() { return ss_().getSheetByName('Students'); }
function clean_(s) { return String(s == null ? '' : s).trim(); }

// Age in years and months, calculated by the sheet from Date of Birth (column I)
function ageFormula_(r) {
  return '=IF(I' + r + '="","",IFERROR(DATEDIF(I' + r + ',TODAY(),"Y")&" years "&DATEDIF(I' + r + ',TODAY(),"YM")&" months",""))';
}

// "2019-03-15" -> Date (noon, to avoid time-zone day shifts); empty -> ''
function parseDate_(s) {
  s = clean_(s);
  if (!s) return '';
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (!m) throw new Error('Invalid date');
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
  if (d.getMonth() !== Number(m[2]) - 1 || d.getFullYear() < 1980) throw new Error('Invalid date');
  return d;
}

function fmtDate_(v, tz) {
  return (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : '';
}

function auth_(cls, pin) {
  const rows = ss_().getSheetByName('Teacher PINs').getDataRange().getValues();
  const ok = rows.some(r => String(r[0]) === cls && String(r[1]) === String(pin).trim());
  if (!ok) throw new Error('Wrong class or PIN.');
}

// ============================================================
// AKTIVITAS HARIAN TIM — Google Apps Script Backend
// Deploy sebagai Web App: Execute as "Me", Access "Anyone"
// ============================================================

const SPREADSHEET_ID = 'GANTI_DENGAN_ID_SPREADSHEET_ANDA';
const SHEET_LAPORAN  = 'Laporan';
const SHEET_STAFF    = 'Staff';

// ─── CORS Helper ────────────────────────────────────────────
function cors(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON);
}

function ok(data) {
  return cors(ContentService.createTextOutput(JSON.stringify({ ok: true, data })));
}

function err(msg) {
  return cors(ContentService.createTextOutput(JSON.stringify({ ok: false, error: msg })));
}

// ─── Init Sheet Headers ─────────────────────────────────────
function initSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let sl = ss.getSheetByName(SHEET_LAPORAN);
  if (!sl) { sl = ss.insertSheet(SHEET_LAPORAN); }
  if (sl.getLastRow() === 0) {
    sl.appendRow([
      'ID','Timestamp','Nama Staff','Tanggal','Status Kehadiran',
      'Leads Baru Masuk','Leads Di-Follow-up','Leads Hasil Follow-up','Leads Terkonversi',
      'Conversion Rate (%)','Metode Follow-up','Tindakan yang Dilakukan',
      'Checklist Aktivitas','Aktivitas Lainnya','Hambatan','Rencana Besok'
    ]);
    sl.getRange(1,1,1,16).setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
    sl.setFrozenRows(1);
  }

  let ss2 = ss.getSheetByName(SHEET_STAFF);
  if (!ss2) {
    ss2 = ss.insertSheet(SHEET_STAFF);
    ss2.appendRow(['Nama Staff','Jabatan','Aktif']);
    ss2.getRange(1,1,1,3).setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
    const defaultStaff = [
      ['Bella Sintia','Marketing Staff','TRUE'],
      ['Irfandi Nyondri','Marketing Staff','TRUE'],
      ['Kasmira','Marketing Staff','TRUE'],
      ['Salma','Marketing Staff','TRUE'],
    ];
    ss2.getRange(2,1,defaultStaff.length,3).setValues(defaultStaff);
  }
}

// ─── doGet — Read Data & Delete ─────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action || 'list';

    if (action === 'list') {
      return getReports(e.parameter.nama || '');
    }
    if (action === 'staff') {
      return getStaff();
    }
    if (action === 'delete') {
      return deleteReport(e.parameter.id);
    }
    if (action === 'rekap') {
      return getRekapStats(e.parameter.nama || '', e.parameter.dari || '', e.parameter.sampai || '');
    }
    return err('Unknown action');
  } catch (ex) {
    return err(ex.message);
  }
}

// ─── doPost — Save Report ────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action || 'save';

    if (action === 'save')       return saveReport(payload);
    if (action === 'save_staff') return saveStaff(payload);
    return err('Unknown action');
  } catch (ex) {
    return err(ex.message);
  }
}

// ─── Save Laporan ────────────────────────────────────────────
function saveReport(d) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_LAPORAN);
  if (!sheet) { initSheets(); return saveReport(d); }

  const fu    = parseInt(d.lFu)   || 0;
  const konv  = parseInt(d.lKonv) || 0;
  const cvr   = fu > 0 ? ((konv / fu) * 100).toFixed(1) : '0.0';

  const row = [
    d.id,
    new Date().toLocaleString('id-ID'),
    d.nama,
    d.tgl,
    d.status,
    parseInt(d.lBaru)   || 0,
    fu,
    parseInt(d.lHasil)  || 0,
    konv,
    parseFloat(cvr),
    (d.metode   || []).join(', '),
    d.tindakan  || '',
    (d.aktivitas|| []).join(', '),
    d.aktLain   || '',
    d.hambatan  || '',
    d.rencana   || ''
  ];

  sheet.appendRow(row);

  // Auto-format conversion rate cell (column J = 10)
  const lastRow = sheet.getLastRow();
  const cvrVal  = parseFloat(cvr);
  const cvrCell = sheet.getRange(lastRow, 10);
  if      (cvrVal >= 70) cvrCell.setBackground('#c8e6c9');
  else if (cvrVal >= 40) cvrCell.setBackground('#fff9c4');
  else                    cvrCell.setBackground('#ffcdd2');

  return ok({ id: d.id, cvr });
}

// ─── Get Reports ─────────────────────────────────────────────
function getReports(filterNama) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_LAPORAN);
  if (!sheet || sheet.getLastRow() <= 1) return ok([]);

  const values  = sheet.getDataRange().getValues();
  const headers = values[0];
  let rows = values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });

  if (filterNama) rows = rows.filter(r => r['Nama Staff'] === filterNama);
  return ok(rows.reverse()); // newest first
}

// ─── Get Rekap Stats ─────────────────────────────────────────
function getRekapStats(filterNama, dari, sampai) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_LAPORAN);
  if (!sheet || sheet.getLastRow() <= 1) return ok({ stats: [], summary: {} });

  const values  = sheet.getDataRange().getValues();
  const headers = values[0];
  let rows = values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });

  if (filterNama) rows = rows.filter(r => r['Nama Staff'] === filterNama);
  if (dari)       rows = rows.filter(r => r['Tanggal'] >= dari);
  if (sampai)     rows = rows.filter(r => r['Tanggal'] <= sampai);

  // Per-staff aggregation
  const byStaff = {};
  rows.forEach(r => {
    const nama = r['Nama Staff'];
    if (!byStaff[nama]) byStaff[nama] = { nama, laporan:0, lBaru:0, lFu:0, lHasil:0, lKonv:0 };
    byStaff[nama].laporan++;
    byStaff[nama].lBaru  += parseInt(r['Leads Baru Masuk'])      || 0;
    byStaff[nama].lFu    += parseInt(r['Leads Di-Follow-up'])    || 0;
    byStaff[nama].lHasil += parseInt(r['Leads Hasil Follow-up']) || 0;
    byStaff[nama].lKonv  += parseInt(r['Leads Terkonversi'])     || 0;
  });

  const stats = Object.values(byStaff).map(s => ({
    ...s,
    cvr: s.lFu > 0 ? ((s.lKonv / s.lFu) * 100).toFixed(1) : '0.0',
    fuRate: s.lBaru > 0 ? ((s.lFu / s.lBaru) * 100).toFixed(1) : '0.0'
  }));

  const totBaru  = rows.reduce((a, r) => a + (parseInt(r['Leads Baru Masuk'])      || 0), 0);
  const totFu    = rows.reduce((a, r) => a + (parseInt(r['Leads Di-Follow-up'])    || 0), 0);
  const totKonv  = rows.reduce((a, r) => a + (parseInt(r['Leads Terkonversi'])     || 0), 0);

  return ok({
    stats,
    summary: {
      totalLaporan: rows.length,
      totBaru, totFu, totKonv,
      cvrTotal: totFu > 0 ? ((totKonv / totFu) * 100).toFixed(1) : '0.0'
    }
  });
}

// ─── Get Staff ───────────────────────────────────────────────
function getStaff() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_STAFF);
  if (!sheet || sheet.getLastRow() <= 1) return ok([]);

  const values  = sheet.getDataRange().getValues();
  const headers = values[0];
  const rows = values.slice(1)
    .map(row => { const o={}; headers.forEach((h,i) => o[h]=row[i]); return o; })
    .filter(r => r['Aktif'] === true || r['Aktif'] === 'TRUE');
  return ok(rows);
}

// ─── Save Staff ──────────────────────────────────────────────
function saveStaff(d) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_STAFF);
  if (!sheet) { initSheets(); return saveStaff(d); }
  sheet.appendRow([d.nama, d.jabatan || 'Staff', 'TRUE']);
  return ok({ nama: d.nama });
}

// ─── Delete Report ───────────────────────────────────────────
function deleteReport(id) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_LAPORAN);
  if (!sheet) return err('Sheet not found');

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return ok({ deleted: id });
    }
  }
  return err('Row not found');
}

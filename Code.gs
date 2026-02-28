// ══════════════════════════════════════════════════════════════════
//  SMK Ledger Nilai — Google Apps Script Web App
//  Cara deploy:
//    1. Buka spreadsheet → Extensions → Apps Script
//    2. Paste seluruh kode ini → Save
//    3. Deploy → New Deployment → Web App
//       • Execute as: Me
//       • Who has access: Anyone
//    4. Copy URL yang muncul → paste di UI Pengaturan (URL GAS)
// ══════════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1y6mLBymxnmzP9khH1ZLU5V3eOWmXc-X8lp9cBNf0Raw';
const SHEET_GID      = '1121452969';
const DATA_START_ROW = 3;  // baris pertama data (bukan header)

// Offset kolom awal tiap semester (0-indexed dari kolom A)
// SESUAIKAN jika struktur spreadsheet berbeda!
const SEM_COL_START = {
  1: 6,   // Sem 1: kolom G
  2: 19,  // Sem 2: kolom T
  3: 32,  // Sem 3: kolom AG
  4: 44,  // Sem 4
  5: 56,  // Sem 5
  6: 66,  // Sem 6
};

const SEM_MAPEL = {
  1: ['Pend. Agama','Pend. Pancasila','Bhs. Indonesia','PJOK','Sejarah','Mulok','Seni Budaya','Matematika','Bhs. Inggris','Informatika','Projek IPAS','DPK','Rata-Rata'],
  2: ['Pend. Agama','Pend. Pancasila','Bhs. Indonesia','PJOK','Sejarah','Mulok','Seni Budaya','Matematika','Bhs. Inggris','Informatika','Projek IPAS','DPK','Rata-Rata'],
  3: ['Pend. Agama','Pend. Pancasila','Bhs. Indonesia','PJOK','Sejarah','Mulok','Matematika','Bhs. Inggris','KK','PKK','Mapel Pilihan','Rata-Rata'],
  4: ['Pend. Agama','Pend. Pancasila','Bhs. Indonesia','PJOK','Sejarah','Mulok','Matematika','Bhs. Inggris','KK','PKK','Mapel Pilihan','Rata-Rata'],
  5: ['Pend. Agama','Pend. Pancasila','Bhs. Indonesia','Mulok','Matematika','Bhs. Inggris','KK','PKK','Mapel Pilihan','Rata-Rata'],
  6: ['Pend. Agama','Pend. Pancasila','Bhs. Indonesia','Mulok','Matematika','Bhs. Inggris','KK','PKK','Mapel Pilihan','Rata-Rata'],
};

// ─── Helper JSON output ───────────────────────────────────────────
function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── GET: ping & ambil data ───────────────────────────────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'ping')     return jsonOut({ status: 'ok', message: 'GAS aktif!' });
  if (action === 'getSiswa') return jsonOut(getSiswaData());
  return jsonOut({ status: 'ok', message: 'SMK Ledger GAS aktif.' });
}

// ─── POST: simpan nilai & data siswa ─────────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    const data   = body.data;

    if (action === 'updateNilai')  return jsonOut(updateNilai(data));
    if (action === 'addSiswa')     return jsonOut(addSiswa(data));
    if (action === 'updateSiswa')  return jsonOut(updateSiswa(data));

    return jsonOut({ status: 'error', message: 'Action tidak dikenal: ' + action });
  } catch (err) {
    return jsonOut({ status: 'error', message: err.toString() });
  }
}

// ══════════════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════════════

function getSheet() {
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  for (const sh of sheets) {
    if (String(sh.getSheetId()) === String(SHEET_GID)) return sh;
  }
  return ss.getActiveSheet();
}

/** Cari nomor baris (1-indexed) berdasarkan NIS di kolom B. Return -1 jika tidak ada. */
function findRowByNIS(sheet, nis) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return -1;
  const vals = sheet.getRange(DATA_START_ROW, 2, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(nis).trim()) return DATA_START_ROW + i;
  }
  return -1;
}

// ══════════════════════════════════════════════════════════════════
//  FUNGSI UTAMA
// ══════════════════════════════════════════════════════════════════

/**
 * Update nilai semester.
 * Payload: { nis, semester, nilai: { 'Mapel': angka, ... } }
 */
function updateNilai(data) {
  if (!data || !data.nis || !data.semester || !data.nilai) {
    return { status: 'error', message: 'Field wajib: nis, semester, nilai' };
  }

  const sheet    = getSheet();
  const rowNum   = findRowByNIS(sheet, data.nis);

  if (rowNum === -1) {
    return { status: 'error', message: 'NIS "' + data.nis + '" tidak ditemukan di spreadsheet.' };
  }

  const sem      = parseInt(data.semester);
  const mapel    = SEM_MAPEL[sem];
  const colStart = SEM_COL_START[sem];

  if (!mapel || colStart === undefined) {
    return { status: 'error', message: 'Semester tidak valid: ' + sem };
  }

  let count = 0;
  mapel.forEach(function(m, idx) {
    const val = data.nilai[m];
    if (val !== undefined && val !== null && String(val).trim() !== '') {
      sheet.getRange(rowNum, colStart + idx + 1).setValue(Number(val));
      count++;
    }
  });

  return {
    status  : 'ok',
    message : 'Berhasil menyimpan ' + count + ' nilai Sem-' + sem + ' untuk NIS ' + data.nis,
    row     : rowNum,
    saved   : count
  };
}

/**
 * Tambah siswa baru.
 * Payload: { no, nis, nisn, nama, prodi, kelas }
 */
function addSiswa(data) {
  if (!data || !data.nis || !data.nama) {
    return { status: 'error', message: 'Field wajib: nis, nama' };
  }
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();
  const newRow  = lastRow + 1;

  sheet.getRange(newRow, 1).setValue(data.no    || '');
  sheet.getRange(newRow, 2).setValue(data.nis   || '');
  sheet.getRange(newRow, 3).setValue(data.nisn  || '');
  sheet.getRange(newRow, 4).setValue(data.nama  || '');
  sheet.getRange(newRow, 5).setValue(data.prodi || '');
  sheet.getRange(newRow, 6).setValue(data.kelas || '');

  return { status: 'ok', message: 'Siswa ditambahkan di baris ' + newRow, row: newRow };
}

/**
 * Update identitas siswa (cari by NIS).
 * Payload: { nis, nisn?, nama?, prodi?, kelas? }
 */
function updateSiswa(data) {
  if (!data || !data.nis) return { status: 'error', message: 'NIS wajib diisi' };

  const sheet  = getSheet();
  const rowNum = findRowByNIS(sheet, data.nis);
  if (rowNum === -1) return { status: 'error', message: 'NIS "' + data.nis + '" tidak ditemukan' };

  if (data.nisn  !== undefined) sheet.getRange(rowNum, 3).setValue(data.nisn);
  if (data.nama  !== undefined) sheet.getRange(rowNum, 4).setValue(data.nama);
  if (data.prodi !== undefined) sheet.getRange(rowNum, 5).setValue(data.prodi);
  if (data.kelas !== undefined) sheet.getRange(rowNum, 6).setValue(data.kelas);

  return { status: 'ok', message: 'Data siswa diperbarui', row: rowNum };
}

/** Ambil semua data siswa */
function getSiswaData() {
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { status: 'ok', data: [] };

  const vals = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 6).getValues();
  const data = vals
    .filter(function(r) { return r[0] !== '' && r[3] !== ''; })
    .map(function(r) {
      return { no: r[0], nis: String(r[1]||''), nisn: String(r[2]||''), nama: String(r[3]||''), prodi: String(r[4]||''), kelas: String(r[5]||'') };
    });

  return { status: 'ok', data: data };
}

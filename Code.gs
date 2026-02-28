// ══════════════════════════════════════════════════════════════════
//  SMK Ledger Nilai — Google Apps Script Web App
//  FIX: Menggunakan doGet dengan query params untuk bypass CORS.
//       Browser tidak bisa POST ke GAS karena CORS preflight blocked.
//       Semua operasi dikirim via GET + ?payload=<JSON encoded>
//
//  Cara deploy:
//    1. Buka spreadsheet → Extensions → Apps Script
//    2. HAPUS semua kode lama, paste seluruh kode ini → Save
//    3. Deploy → New Deployment → Web App
//       • Execute as: Me
//       • Who has access: Anyone
//    4. Copy URL → paste di UI Pengaturan (URL GAS)
//    ⚠️  Setiap update kode: buat New Deployment (bukan redeploy)!
// ══════════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1y6mLBymxnmzP9khH1ZLU5V3eOWmXc-X8lp9cBNf0Raw';
const SHEET_GID      = '1121452969';
const DATA_START_ROW = 3;

// Offset kolom awal tiap semester (0-indexed dari kolom A)
const SEM_COL_START = {
  1: 6,
  2: 19,
  3: 32,
  4: 44,
  5: 56,
  6: 66,
};

const SEM_MAPEL = {
  1: ['Pend. Agama','Pend. Pancasila','Bhs. Indonesia','PJOK','Sejarah','Mulok','Seni Budaya','Matematika','Bhs. Inggris','Informatika','Projek IPAS','DPK','Rata-Rata'],
  2: ['Pend. Agama','Pend. Pancasila','Bhs. Indonesia','PJOK','Sejarah','Mulok','Seni Budaya','Matematika','Bhs. Inggris','Informatika','Projek IPAS','DPK','Rata-Rata'],
  3: ['Pend. Agama','Pend. Pancasila','Bhs. Indonesia','PJOK','Sejarah','Mulok','Matematika','Bhs. Inggris','KK','PKK','Mapel Pilihan','Rata-Rata'],
  4: ['Pend. Agama','Pend. Pancasila','Bhs. Indonesia','PJOK','Sejarah','Mulok','Matematika','Bhs. Inggris','KK','PKK','Mapel Pilihan','Rata-Rata'],
  5: ['Pend. Agama','Pend. Pancasila','Bhs. Indonesia','Mulok','Matematika','Bhs. Inggris','KK','PKK','Mapel Pilihan','Rata-Rata'],
  6: ['Pend. Agama','Pend. Pancasila','Bhs. Indonesia','Mulok','Matematika','Bhs. Inggris','KK','PKK','Mapel Pilihan','Rata-Rata'],
};

// ─── Helper JSON output dengan CORS headers ───────────────────────
function jsonOut(data) {
  // GAS ContentService tidak support custom headers,
  // tapi untuk doGet response akan dilewatkan browser tanpa CORS block
  // karena GAS meng-handle redirect ke googleusercontent domain yang punya wildcard CORS.
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════════════
//  doGet — SEMUA request ditangani sini (GET + payload param)
//  FIX UTAMA: Browser mengirim semua action via GET request dengan
//             ?action=xxx&payload=<URLEncoded JSON>
//             sehingga tidak ada CORS preflight issue.
// ══════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    const params  = e ? (e.parameter || {}) : {};
    const action  = params.action || '';

    // ── PING: test koneksi ────────────────────────────────────────
    if (action === 'ping') {
      return jsonOut({ status: 'ok', message: 'GAS aktif dan siap digunakan!' });
    }

    // ── GET SISWA: ambil data ─────────────────────────────────────
    if (action === 'getSiswa') {
      return jsonOut(getSiswaData());
    }

    // ── UPDATE NILAI: simpan nilai semester ───────────────────────
    if (action === 'updateNilai') {
      const payload = JSON.parse(decodeURIComponent(params.payload || '{}'));
      return jsonOut(updateNilai(payload));
    }

    // ── ADD SISWA: tambah baris baru ──────────────────────────────
    if (action === 'addSiswa') {
      const payload = JSON.parse(decodeURIComponent(params.payload || '{}'));
      return jsonOut(addSiswa(payload));
    }

    // ── UPDATE SISWA: edit identitas ──────────────────────────────
    if (action === 'updateSiswa') {
      const payload = JSON.parse(decodeURIComponent(params.payload || '{}'));
      return jsonOut(updateSiswa(payload));
    }

    // Default
    return jsonOut({ status: 'ok', message: 'SMK Ledger GAS aktif. Gunakan parameter ?action=ping untuk tes.' });

  } catch (err) {
    return jsonOut({ status: 'error', message: 'doGet error: ' + err.toString() });
  }
}

// ══════════════════════════════════════════════════════════════════
//  doPost — tetap ada sebagai fallback (jika dipanggil langsung)
// ══════════════════════════════════════════════════════════════════
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
    return jsonOut({ status: 'error', message: 'doPost error: ' + err.toString() });
  }
}

// ══════════════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════════════

function getSheet() {
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (String(sheets[i].getSheetId()) === String(SHEET_GID)) return sheets[i];
  }
  return ss.getActiveSheet();
}

/** Cari nomor baris (1-indexed) berdasarkan NIS di kolom B. Return -1 jika tidak ada. */
function findRowByNIS(sheet, nis) {
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return -1;
  var vals = sheet.getRange(DATA_START_ROW, 2, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
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
    return { status: 'error', message: 'Field wajib kurang: nis=' + data.nis + ', semester=' + data.semester };
  }

  var sheet    = getSheet();
  var rowNum   = findRowByNIS(sheet, data.nis);

  if (rowNum === -1) {
    return { status: 'error', message: 'NIS "' + data.nis + '" tidak ditemukan di spreadsheet.' };
  }

  var sem      = parseInt(data.semester);
  var mapel    = SEM_MAPEL[sem];
  var colStart = SEM_COL_START[sem];

  if (!mapel || colStart === undefined) {
    return { status: 'error', message: 'Semester tidak valid: ' + sem };
  }

  var count = 0;
  var updates = [];

  for (var idx = 0; idx < mapel.length; idx++) {
    var m   = mapel[idx];
    var val = data.nilai[m];
    if (val !== undefined && val !== null && String(val).trim() !== '') {
      // Batch update lebih efisien
      updates.push({ col: colStart + idx + 1, val: Number(val) });
      count++;
    }
  }

  // Tulis ke sheet
  for (var j = 0; j < updates.length; j++) {
    sheet.getRange(rowNum, updates[j].col).setValue(updates[j].val);
  }

  // Flush untuk memastikan data tersimpan sebelum response
  SpreadsheetApp.flush();

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

  var sheet   = getSheet();

  // Cek duplikat NIS
  var existing = findRowByNIS(sheet, data.nis);
  if (existing !== -1) {
    return { status: 'error', message: 'NIS "' + data.nis + '" sudah ada di baris ' + existing };
  }

  var lastRow = sheet.getLastRow();
  var newRow  = lastRow + 1;

  var rowData = [
    data.no    || '',
    data.nis   || '',
    data.nisn  || '',
    data.nama  || '',
    data.prodi || '',
    data.kelas || ''
  ];

  sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
  SpreadsheetApp.flush();

  return { status: 'ok', message: 'Siswa "' + data.nama + '" ditambahkan di baris ' + newRow, row: newRow };
}

/**
 * Update identitas siswa (cari by NIS).
 * Payload: { nis, nisn?, nama?, prodi?, kelas? }
 */
function updateSiswa(data) {
  if (!data || !data.nis) return { status: 'error', message: 'NIS wajib diisi' };

  var sheet  = getSheet();
  var rowNum = findRowByNIS(sheet, data.nis);
  if (rowNum === -1) return { status: 'error', message: 'NIS "' + data.nis + '" tidak ditemukan' };

  if (data.nisn  !== undefined) sheet.getRange(rowNum, 3).setValue(data.nisn);
  if (data.nama  !== undefined) sheet.getRange(rowNum, 4).setValue(data.nama);
  if (data.prodi !== undefined) sheet.getRange(rowNum, 5).setValue(data.prodi);
  if (data.kelas !== undefined) sheet.getRange(rowNum, 6).setValue(data.kelas);

  SpreadsheetApp.flush();

  return { status: 'ok', message: 'Data siswa diperbarui di baris ' + rowNum, row: rowNum };
}

/** Ambil semua data siswa */
function getSiswaData() {
  var sheet   = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { status: 'ok', data: [] };

  var vals = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 6).getValues();
  var data = vals
    .filter(function(r) { return r[0] !== '' && r[3] !== ''; })
    .map(function(r) {
      return {
        no    : r[0],
        nis   : String(r[1] || ''),
        nisn  : String(r[2] || ''),
        nama  : String(r[3] || ''),
        prodi : String(r[4] || ''),
        kelas : String(r[5] || '')
      };
    });

  return { status: 'ok', data: data };
}

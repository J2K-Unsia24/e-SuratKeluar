// ==========================================
// KONFIGURASI SPREADSHEET
// ==========================================
const SPREADSHEET_ID = '1QeR9tDGIdhjoQf2nqKItmHBJuFGqPBrvCpiegidVzlA'; // Ganti dengan ID Google Sheet Anda

// Nama sheet
const SHEET_USERS = 'Users';
const SHEET_FORMAT = 'FormatNomor';
const SHEET_ARSIP = 'ArsipSurat';

// ==========================================
// FUNGSI UTAMA (Web App)
// ==========================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sistem Informasi Persuratan TI')
    .setFaviconUrl('https://img.icons8.com/color/48/000000/mail.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// INISIALISASI SPREADSHEET
// ==========================================
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  let sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_USERS);
    sheet.getRange('A1:E1').setValues([['ID', 'Nama', 'Username', 'Password', 'Role']]);
    sheet.getRange('A2:E2').setValues([['U1', 'Administrator TI', 'admin', 'admin2026', 'Admin']]);
  }
  
  sheet = ss.getSheetByName(SHEET_FORMAT);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_FORMAT);
    sheet.getRange('A1:C1').setValues([['ID', 'Jenis', 'Kode']]);
    sheet.getRange('A2:C4').setValues([
      ['F1', 'Surat Keputusan', 'SK-TI'],
      ['F2', 'Surat Undangan', 'UND-TI'],
      ['F3', 'Surat Keterangan', 'KET-TI']
    ]);
  }
  
  sheet = ss.getSheetByName(SHEET_ARSIP);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ARSIP);
    sheet.getRange('A1:J1').setValues([['ID', 'Nomor', 'TglSurat', 'TglKeluar', 'Jenis', 'Perihal', 'Tujuan', 'Pembuat', 'FileUrl', 'FileName']]);
  }
  
  return { success: true, message: 'Spreadsheet berhasil diinisialisasi' };
}

// ==========================================
// FUNGSI LOGIN
// ==========================================
function login(username, password) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === username && data[i][3] === password) {
        return {
          success: true,
          data: { id: data[i][0], nama: data[i][1], username: data[i][2], role: data[i][4] }
        };
      }
    }
    return { success: false, message: 'Username atau password salah!' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// FUNGSI DASHBOARD
// ==========================================
function getDashboardData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_ARSIP);
    const data = sheet.getDataRange().getValues();
    
    const total = data.length - 1;
    const now = new Date();
    const bulanIni = now.getMonth() + 1;
    const tahunIni = now.getFullYear();
    
    let countBulanIni = 0;
    for (let i = 1; i < data.length; i++) {
      const tglSurat = new Date(data[i][2]);
      if (tglSurat.getMonth() + 1 === bulanIni && tglSurat.getFullYear() === tahunIni) {
        countBulanIni++;
      }
    }
    
    return { success: true, total: total, bulanIni: countBulanIni };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// FUNGSI FORMAT NOMOR
// ==========================================
function getFormatNomor() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_FORMAT);
    const data = sheet.getDataRange().getValues();
    
    const formats = [];
    for (let i = 1; i < data.length; i++) {
      formats.push({ id: data[i][0], jenis: data[i][1], kode: data[i][2] });
    }
    return { success: true, data: formats };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function tambahFormat(jenis, kode) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_FORMAT);
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === jenis) return { success: false, message: 'Jenis surat sudah ada!' };
      if (data[i][2] === kode) return { success: false, message: 'Kode surat sudah digunakan!' };
    }
    
    const id = 'F' + new Date().getTime();
    sheet.appendRow([id, jenis, kode]);
    return { success: true, message: 'Format berhasil ditambahkan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function hapusFormat(id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_FORMAT);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Format berhasil dihapus' };
      }
    }
    return { success: false, message: 'Format tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// FUNGSI ARSIP SURAT
// ==========================================
function getArsip(role, namaUser) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_ARSIP);
    const data = sheet.getDataRange().getValues();
    
    const arsip = [];
    for (let i = 1; i < data.length; i++) {
      const item = {
        id: data[i][0],
        nomor: data[i][1],
        tgl_surat: formatDate(data[i][2]),
        tgl_keluar: formatDate(data[i][3]),
        jenis: data[i][4],
        perihal: data[i][5],
        tujuan: data[i][6],
        pembuat: data[i][7],
        file: data[i][8],
        fileName: data[i][9]
      };
      
      if (role === 'Admin' || item.pembuat === namaUser) {
        arsip.push(item);
      }
    }
    
    arsip.sort((a, b) => b.tgl_keluar.localeCompare(a.tgl_keluar));
    return { success: true, data: arsip };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function buatSurat(dataSurat) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetArsip = ss.getSheetByName(SHEET_ARSIP);
    const sheetFormat = ss.getSheetByName(SHEET_FORMAT);
    
    // Generate nomor surat
    const formatData = sheetFormat.getDataRange().getValues();
    let kodeFormat = '';
    for (let i = 1; i < formatData.length; i++) {
      if (formatData[i][1] === dataSurat.jenis) {
        kodeFormat = formatData[i][2];
        break;
      }
    }
    
    if (!kodeFormat) return { success: false, message: 'Format nomor tidak ditemukan' };
    
    // Hitung nomor urut
    const arsipData = sheetArsip.getDataRange().getValues();
    const tahun = new Date(dataSurat.tgl_surat).getFullYear();
    let maxUrut = 0;
    
    for (let i = 1; i < arsipData.length; i++) {
      if (arsipData[i][4] === dataSurat.jenis && arsipData[i][2]) {
        const tglSurat = new Date(arsipData[i][2]);
        if (tglSurat.getFullYear() === tahun) {
          const urut = parseInt(arsipData[i][1].split('/')[0]);
          if (!isNaN(urut) && urut > maxUrut) maxUrut = urut;
        }
      }
    }
    
    const nextUrut = maxUrut + 1;
    const nomorUrut = String(nextUrut).padStart(3, '0');
    const finalNomor = `${nomorUrut}/${kodeFormat}/${tahun}`;
    
    // Upload file ke Google Drive jika ada
    let fileUrl = '';
    let fileName = '';
    if (dataSurat.file_base64) {
      try {
        // GANTI 'YOUR_FOLDER_ID' dengan ID folder Google Drive Anda
        const folderId = '1nwFsrRaCJ8piMOOh_e39goWAIol4ppMX'; // <-- GANTI DI SINI
        const folder = DriveApp.getFolderById(folderId);
        const blob = Utilities.newBlob(
          Utilities.base64Decode(dataSurat.file_base64), 
          dataSurat.mime_type, 
          dataSurat.file_name
        );
        const file = folder.createFile(blob);
        fileUrl = file.getUrl();
        fileName = file.getName();
      } catch (uploadError) {
        console.log('Upload file gagal: ' + uploadError.toString());
        // Tetap lanjutkan menyimpan data tanpa file
      }
    }
    
    // Simpan ke sheet
    const id = 'S' + new Date().getTime();
    sheetArsip.appendRow([
      id, finalNomor, new Date(dataSurat.tgl_surat), new Date(), 
      dataSurat.jenis, dataSurat.perihal, dataSurat.tujuan, 
      dataSurat.pembuat, fileUrl, fileName
    ]);
    
    return { success: true, nomor: finalNomor };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function editArsip(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_ARSIP);
    const arsipData = sheet.getDataRange().getValues();
    
    for (let i = 1; i < arsipData.length; i++) {
      if (arsipData[i][0] === data.id) {
        sheet.getRange(i + 1, 3).setValue(new Date(data.tgl_surat));
        sheet.getRange(i + 1, 6).setValue(data.perihal);
        sheet.getRange(i + 1, 7).setValue(data.tujuan);
        return { success: true, message: 'Data berhasil diupdate' };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// FUNGSI PENGELOLAAN USER
// ==========================================
function getPengguna() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    
    const users = [];
    for (let i = 1; i < data.length; i++) {
      users.push({ id: data[i][0], nama: data[i][1], username: data[i][2], password: data[i][3], role: data[i][4] });
    }
    return { success: true, data: users };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function tambahPengguna(nama, username, password, role) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USERS);
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === username) return { success: false, message: 'Username sudah digunakan' };
    }
    
    const id = 'U' + new Date().getTime();
    sheet.appendRow([id, nama, username, password, role]);
    return { success: true, message: 'Pengguna berhasil ditambahkan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function editPengguna(id, nama, username, password, role) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    
    // Cek username sudah ada (kecuali milik sendiri)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== id && data[i][2] === username) {
        return { success: false, message: 'Username sudah digunakan oleh pengguna lain!' };
      }
    }
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.getRange(i + 1, 2).setValue(nama);
        sheet.getRange(i + 1, 3).setValue(username);
        sheet.getRange(i + 1, 4).setValue(password);
        sheet.getRange(i + 1, 5).setValue(role);
        return { success: true, message: 'Data pengguna berhasil diperbarui' };
      }
    }
    
    return { success: false, message: 'Pengguna tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function hapusPengguna(id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        if (data[i][4] === 'Admin') {
          let adminCount = 0;
          for (let j = 1; j < data.length; j++) if (data[j][4] === 'Admin') adminCount++;
          if (adminCount <= 1) return { success: false, message: 'Tidak dapat menghapus satu-satunya Admin' };
        }
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Pengguna berhasil dihapus' };
      }
    }
    return { success: false, message: 'Pengguna tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// FUNGSI HELPER
// ==========================================
function formatDate(date) {
  if (!date || date === '') return '';
  const d = new Date(date);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

/**
 * Aplikasi Penilaian Praktek Khutbah - Backend GAS
 * ------------------------------------------------
 * Skrip ini dipasang pada Google Apps Script yang terhubung dengan Google Sheets.
 * Menyediakan fungsi doGet untuk mengambil rangkuman data dan doPost untuk menyimpan data baru.
 */

// ID Google Spreadsheet Anda (diambil dari URL yang Anda berikan)
var SPREADSHEET_ID = "1Pd_4GsH5LQSH74r_L3D6hpUq--7e-3t3dqpKZyAQHk4";
// Nama Sheet tempat menyimpan data mentah
var SHEET_NAME = "RawData";

/**
 * Mempersiapkan Sheet "RawData" dengan header kolom jika belum ada.
 * Mengaktifkan spreadsheet berdasarkan ID spesifik agar data terjamin masuk ke sheet yang benar.
 */
function setupSheet() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (error) {
    // Fallback jika terjadi kesalahan pencarian ID, coba gunakan Active Spreadsheet
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  if (!ss) {
    throw new Error("Spreadsheet tidak ditemukan! Pastikan ID Spreadsheet benar dan skrip diizinkan mengaksesnya.");
  }
  
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Menulis header kolom
    sheet.appendRow([
      "Timestamp", 
      "Nama_Penilai", 
      "Nama_Khotib", 
      "Daya_Tarik", 
      "Isi", 
      "Gaya_Penyampaian", 
      "Komentar"
    ]);
    // Format header: Bold, latar abu muda, text terpusat
    var headerRange = sheet.getRange("A1:G1");
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#f3f4f6");
    headerRange.setHorizontalAlignment("center");
    
    // Bekukan baris pertama agar header tidak ikut tergulir
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Menangani permintaan HTTP POST dari Form Penilaian (index.html).
 * Menyimpan data penilaian baru ke dalam Google Sheets.
 */
function doPost(e) {
  var responseHeaders = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };
  
  var result = { status: "success", message: "Data penilaian berhasil disimpan!" };
  
  try {
    var sheet = setupSheet();
    var payload;
    
    // Parsing payload data dari JSON
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      payload = e.parameter;
    }
    
    var timestamp = new Date();
    var namaPenilai = payload.Nama_Penilai || payload.namaPenilai || "";
    var namaKhotib = payload.Nama_Khotib || payload.namaKhotib || "";
    var dayaTarik = Number(payload.Daya_Tarik || payload.dayaTarik || 0);
    var isi = Number(payload.Isi || payload.isi || 0);
    var gayaPenyampaian = Number(payload.Gaya_Penyampaian || payload.gayaPenyampaian || 0);
    var komentar = payload.Komentar || payload.komentar || "";
    
    // Validasi data masukan
    if (!namaPenilai.toString().trim()) {
      throw new Error("Nama Penilai wajib diisi.");
    }
    if (!namaKhotib.toString().trim()) {
      throw new Error("Nama Khotib wajib dipilih.");
    }
    if (dayaTarik < 1 || dayaTarik > 5 || isi < 1 || isi > 5 || gayaPenyampaian < 1 || gayaPenyampaian > 5) {
      throw new Error("Nilai rating untuk ketiga kategori wajib diisi (1-5).");
    }
    
    // Masukkan data ke sheet
    sheet.appendRow([
      timestamp,
      namaPenilai.toString().trim(),
      namaKhotib.toString().trim(),
      dayaTarik,
      isi,
      gayaPenyampaian,
      komentar.toString().trim()
    ]);
    
  } catch (error) {
    result.status = "error";
    result.message = error.toString();
  }
  
  // Mengembalikan respons dengan CORS header
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Menangani permintaan HTTP GET dari Dashboard Monitor (dashboard.html).
 * Mengelompokkan data berdasarkan Nama Khotib dan menghitung rata-rata serta komentar.
 */
function doGet(e) {
  var result = { status: "success", data: {} };
  
  try {
    var sheet = setupSheet();
    var data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      // Jika hanya ada baris header saja
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var khotibGroups = {};
    
    // Memproses baris data (lewati baris 0 karena merupakan header)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var namaKhotib = row[2]; // Kolom C: Nama_Khotib
      
      if (!namaKhotib || namaKhotib.toString().trim() === "") continue;
      
      var key = namaKhotib.toString().trim();
      
      if (!khotibGroups[key]) {
        khotibGroups[key] = {
          nama: key,
          totalPenilai: 0,
          sumDayaTarik: 0,
          sumIsi: 0,
          sumGayaPenyampaian: 0,
          komentar: []
        };
      }
      
      var group = khotibGroups[key];
      group.totalPenilai += 1;
      group.sumDayaTarik += Number(row[3]) || 0; // Kolom D: Daya_Tarik
      group.sumIsi += Number(row[4]) || 0;       // Kolom E: Isi
      group.sumGayaPenyampaian += Number(row[5]) || 0; // Kolom F: Gaya_Penyampaian
      
      var kom = row[6]; // Kolom G: Komentar
      if (kom && kom.toString().trim() !== "") {
        group.komentar.push({
          penilai: row[1] || "Anonim", // Kolom B: Nama_Penilai
          isi: kom.toString().trim(),
          tanggal: row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "dd MMM yyyy HH:mm") : ""
        });
      }
    }
    
    // Menghitung rata-rata akhir per Khotib
    var formattedData = {};
    for (var khotibName in khotibGroups) {
      var group = khotibGroups[khotibName];
      var total = group.totalPenilai;
      
      formattedData[khotibName] = {
        nama: group.nama,
        totalPenilai: total,
        avgDayaTarik: Number((group.sumDayaTarik / total).toFixed(2)),
        avgIsi: Number((group.sumIsi / total).toFixed(2)),
        avgGayaPenyampaian: Number((group.sumGayaPenyampaian / total).toFixed(2)),
        komentar: group.komentar.reverse() // Komentar terbaru ditampilkan di atas
      };
    }
    
    result.data = formattedData;
    
  } catch (error) {
    result.status = "error";
    result.message = error.toString();
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

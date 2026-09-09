/**
 * ==========================================================================
 * SISTEM MANAJEMEN & SLIP GAJI PEGAWAI - GOOGLE APPS SCRIPT
 * ==========================================================================
 * 
 * PANDUAN DEPLOYMENT (PENTING):
 * 1. Buka Google Apps Script (https://script.google.com/) atau dari Spreadsheet: Extensions > Apps Script.
 * 2. Hapus seluruh kode lama di `Code.gs` dan gantikan dengan seluruh isi file ini.
 * 3. Pastikan ID Spreadsheet di bawah ini sudah sesuai.
 * 4. Klik tombol "Deploy" (Terapkan) di kanan atas > "New deployment" (Deployment baru).
 * 5. Pilih type "Web app" (Aplikasi Web).
 * 6. Setelan Deployment:
 *    - Description : Sistem Gaji API v2 (Input & View)
 *    - Execute as  : Me (Email Anda)
 *    - Who has access: Anyone (Siapa saja)  <-- WAJIB pilih ini agar web bisa akses!
 * 7. Klik "Deploy", izinkan hak akses (Review Permissions > Akun Google Anda > Advanced > Go to ... (unsafe) > Allow).
 * 8. Salin "Web app URL" (URL Aplikasi Web) dan tempelkan di aplikasi web (index.html).
 */

// Ganti ID Spreadsheet jika menggunakan spreadsheet yang berbeda
var SPREADSHEET_ID = "1UKu8peUA3hFZtz0Y2BBDe3ApvV5qU83l4v-8xVuppeE";
var SHEET_NAME = "Sheet1";

/**
 * Handle GET Request: Mengambil seluruh riwayat data gaji pegawai (VIEW)
 */
function doGet(e) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    // Jika sheet belum ada, buatkan
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }
    
    // Cek apakah ada data
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      // Belum ada data atau hanya ada header
      return createJsonResponse({
        status: "success",
        total: 0,
        data: []
      });
    }
    
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var records = [];
    
    // Map data mulai dari baris ke-2 (index 1)
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      
      // Lewati baris kosong
      if (!row[0] && !row[2]) continue;
      
      // Format Timestamp ramah tampilan
      var rawTimestamp = row[0];
      var formattedDate = "";
      if (rawTimestamp instanceof Date) {
        formattedDate = Utilities.formatDate(rawTimestamp, Session.getScriptTimeZone() || "Asia/Jakarta", "dd/MM/yyyy HH:mm");
      } else {
        formattedDate = String(rawTimestamp || "");
      }
      
      var item = {
        row_id: i + 1,
        timestamp: formattedDate,
        toko: row[1] || "-",
        nama: row[2] || "-",
        posisi: row[3] || "-",
        periode: row[4] || "-",
        pokok: Number(row[5]) || 0,
        bonus: Number(row[6]) || 0,
        overtime: Number(row[7]) || 0,
        lainnya_text: row[8] || "Tambahan",
        lainnya: Number(row[9]) || 0,
        bon: Number(row[10]) || 0,
        potongan: Number(row[11]) || 0,
        total: Number(row[12]) || 0
      };
      
      records.push(item);
    }
    
    // Urutkan dari data terbaru di paling atas (LIFO)
    records.reverse();
    
    return createJsonResponse({
      status: "success",
      total: records.length,
      data: records
    });
    
  } catch (err) {
    return createJsonResponse({
      status: "error",
      message: err.message
    });
  }
}

/**
 * Handle POST Request: Menyimpan data slip gaji baru ke Spreadsheet (INPUT)
 */
function doPost(e) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }
    
    // Header standar
    var defaultHeaders = [
      "Timestamp", "Toko", "Nama", "Posisi", "Periode",
      "Gaji Pokok", "Bonus", "Overtime", "Tambahan (ket)", "Tambahan (nominal)",
      "Bon", "Potongan", "Total"
    ];
    
    // Inisialisasi Header jika spreadsheet masih kosong
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(defaultHeaders);
      // Format styling header otomatis
      var headerRange = sheet.getRange(1, 1, 1, defaultHeaders.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#2c3e50");
      headerRange.setFontColor("#ffffff");
    }
    
    // Parsing data JSON dari body request
    var data = {};
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (parseError) {
        // Fallback jika dikirim dalam bentuk form-urlencoded
        data = e.parameter;
      }
    } else if (e.parameter) {
      data = e.parameter;
    }
    
    // Sanitasi data angka
    var pokok = Number(data.pokok) || 0;
    var bonus = Number(data.bonus) || 0;
    var overtime = Number(data.overtime) || 0;
    var lainnya = Number(data.lainnya) || 0;
    var bon = Number(data.bon) || 0;
    var potongan = Number(data.potongan) || 0;
    
    // Hitung ulang total di backend untuk memastikan akurasi data
    var total = (pokok + bonus + overtime + lainnya) - (bon + potongan);
    if (data.total !== undefined && !isNaN(Number(data.total))) {
      total = Number(data.total);
    }
    
    // Cek apakah aksi adalah UPDATE pada baris tertentu
    var rowId = Number(data.row_id);
    if (data.action === "update" && rowId > 1 && rowId <= sheet.getLastRow()) {
      var existingTimestamp = sheet.getRange(rowId, 1).getValue();
      var updateRow = [
        existingTimestamp || new Date(),
        data.toko || "Toko",
        data.nama || "-",
        data.posisi || "-",
        data.periode || "-",
        pokok,
        bonus,
        overtime,
        data.lainnya_text || "Tambahan",
        lainnya,
        bon,
        potongan,
        total
      ];
      sheet.getRange(rowId, 1, 1, updateRow.length).setValues([updateRow]);
      
      return createJsonResponse({
        status: "success",
        message: "Data gaji berhasil diperbarui!",
        row_id: rowId,
        timestamp: new Date().toISOString()
      });
    }

    // Default: Tambahkan baris data baru (CREATE)
    var newRow = [
      new Date(),
      data.toko || "Toko",
      data.nama || "-",
      data.posisi || "-",
      data.periode || "-",
      pokok,
      bonus,
      overtime,
      data.lainnya_text || "Tambahan",
      lainnya,
      bon,
      potongan,
      total
    ];
    
    sheet.appendRow(newRow);
    
    return createJsonResponse({
      status: "success",
      message: "Data gaji berhasil disimpan ke Google Spreadsheet!",
      timestamp: new Date().toISOString()
    });
    
  } catch (err) {
    return createJsonResponse({
      status: "error",
      message: "Gagal menyimpan data: " + err.message
    });
  }
}

/**
 * Helper: Membuat respons JSON standar dengan MimeType JSON
 */
function createJsonResponse(outputObject) {
  return ContentService.createTextOutput(JSON.stringify(outputObject))
    .setMimeType(ContentService.MimeType.JSON);
}


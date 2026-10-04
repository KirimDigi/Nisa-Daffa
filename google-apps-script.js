/**
 * =====================================================================
 * GOOGLE APPS SCRIPT - BUKU TAMU / WISHES & RSVP PERNIKAHAN
 * Nisa & Daffa
 * =====================================================================
 * 
 * PANDUAN CARA MEMASANG:
 * 1. Buat Google Spreadsheet baru di Google Drive (misal: "RSVP & Wishes Nisa & Daffa").
 * 2. Di Google Sheets, klik menu: Ekstensi (Extensions) > Apps Script.
 * 3. Hapus semua kode default di editor Apps Script, lalu salin (paste) SELURUH KODE di file ini.
 * 4. Klik ikon "Simpan" (Save / Ctrl+S).
 * 5. Klik tombol biru "Terapkan" (Deploy) di kanan atas > pilih "Penerapan baru" (New deployment).
 * 6. Klik ikon gerigi (Select type) > pilih "Aplikasi Web" (Web app).
 * 7. Atur konfigurasi berikut:
 *    - Deskripsi: Wishes RSVP Nisa Daffa
 *    - Jalankan sebagai (Execute as): "Saya" (Me / email Anda)
 *    - Siapa yang memiliki akses (Who has access): "Siapa saja" (Anyone) ---> [PENTING!]
 * 8. Klik "Terapkan" (Deploy). Jika diminta izin (Authorization), klik "Tinjau izin", pilih akun Google Anda, klik "Advanced", lalu "Go to (unsafe) / Lanjutkan", dan klik "Allow".
 * 9. Salin URL Aplikasi Web (Web App URL) yang berakhiran "/exec".
 * 10. Buka index.html, cari variabel `webAppUrl` (sekitar baris 9200) dan ganti dengan URL Apps Script Anda:
 *     var webAppUrl = "URL_APPS_SCRIPT_ANDA_DI_SINI";
 * =====================================================================
 */

// Nama sheet (tab) yang digunakan. Jika dibiarkan kosong, akan otomatis memakai sheet pertama.
var SHEET_NAME = "Wishes";

/**
 * Mengambil sheet target atau membuatnya jika belum ada
 */
function getTargetSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.getSheets()[0];
    try {
      sheet.setName(SHEET_NAME);
    } catch (e) {}
  }
  
  // Jika sheet masih kosong, buat baris header
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Nama", "Ucapan / Doa", "Kehadiran", "Jumlah Tamu"]);
    var headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#85A57A");
    headerRange.setFontColor("#FFFFFF");
  }
  
  return sheet;
}

/**
 * Handle POST request (Menerima ucapan dan konfirmasi kehadiran dari web)
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  // Kunci eksekusi selama maksimal 30 detik untuk mencegah tabrakan data (concurrency)
  lock.tryLock(30000);
  
  try {
    var sheet = getTargetSheet();
    var params = {};
    
    // Parse form-urlencoded atau JSON data
    if (e && e.parameter && Object.keys(e.parameter).length > 0) {
      params = e.parameter;
    } else if (e && e.postData && e.postData.contents) {
      try {
        params = JSON.parse(e.postData.contents);
      } catch (err) {
        // Fallback jika dikirim via URL encoded string di contents
        var raw = e.postData.contents.split("&");
        for (var i = 0; i < raw.length; i++) {
          var pair = raw[i].split("=");
          if (pair.length === 2) {
            params[decodeURIComponent(pair[0])] = decodeURIComponent(pair[1].replace(/\+/g, " "));
          }
        }
      }
    }
    
    // Ambil data dari parameter
    var nama = (params.author || params.nama || "").trim();
    var ucapan = (params.comment || params.ucapan || "").trim();
    var rawAttendance = (params.attendance || params.kehadiran || "").trim();
    var guest = (params.guest || params.tamu || "1").toString().trim();
    
    // Normalisasi kehadiran
    var kehadiran = "Hadir";
    if (rawAttendance === "notpresent" || rawAttendance.toLowerCase() === "tidak hadir") {
      kehadiran = "Tidak Hadir";
      guest = "0";
    } else if (rawAttendance === "present" || rawAttendance.toLowerCase() === "hadir") {
      kehadiran = "Hadir";
    }
    
    if (!nama && !ucapan) {
      return createJsonResponse({ status: "error", message: "Nama dan ucapan tidak boleh kosong." });
    }
    
    // Format timestamp WIB (GMT+7)
    var now = new Date();
    var timestampStr = Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd'T'HH:mm:ssXXX");
    var timestampDisplay = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm:ss");
    
    // Tulis ke spreadsheet
    sheet.appendRow([timestampDisplay, nama, ucapan, kehadiran, guest]);
    
    return createJsonResponse({
      status: "success",
      message: "Ucapan berhasil disimpan",
      data: {
        timestamp: timestampStr,
        nama: nama,
        ucapan: ucapan,
        kehadiran: kehadiran,
        guest: guest
      }
    });
    
  } catch (error) {
    return createJsonResponse({ status: "error", message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Handle GET request (Menampilkan daftar ucapan ke halaman web)
 */
function doGet(e) {
  try {
    var sheet = getTargetSheet();
    var lastRow = sheet.getLastRow();
    var result = [];
    
    // Jika ada data selain header
    if (lastRow > 1) {
      // Ambil seluruh data dari baris ke-2 hingga baris terakhir (5 kolom: Timestamp, Nama, Ucapan, Kehadiran, Tamu)
      var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      
      for (var i = 0; i < values.length; i++) {
        var row = values[i];
        var rawTimestamp = row[0];
        var nama = row[1];
        var ucapan = row[2];
        var kehadiran = row[3];
        var guest = row[4];
        
        // Lewati jika nama atau ucapan kosong
        if (!nama && !ucapan) continue;
        
        var isoTimestamp = "";
        if (rawTimestamp instanceof Date) {
          isoTimestamp = Utilities.formatDate(rawTimestamp, "GMT+7", "yyyy-MM-dd'T'HH:mm:ssXXX");
        } else if (typeof rawTimestamp === "string" && rawTimestamp !== "") {
          isoTimestamp = rawTimestamp;
        }
        
        result.push({
          timestamp: isoTimestamp,
          nama: String(nama),
          ucapan: String(ucapan),
          kehadiran: String(kehadiran || "Hadir"),
          guest: String(guest || "1")
        });
      }
    }
    
    return createJsonResponse(result);
    
  } catch (error) {
    return createJsonResponse({ status: "error", message: error.toString() });
  }
}

/**
 * Helper untuk membuat response JSON dengan MIME type JSON
 */
function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

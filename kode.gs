/**
 * SISTEM SPMB - Google Apps Script
 * Script untuk menangani pendaftaran siswa baru
 */

function doPost(e) {
  try {
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;

    setupSheet();

    // 1. Aksi Submit Form Pendaftaran
    if (action === "submitForm") {
      var result = handleSubmit(requestData.data);
      return createResponse({ success: true, regNo: result });
    }

    // 2. Aksi Ambil Data (Untuk Halaman Admin)
    if (action === "getData") {
      var allData = handleGetData();
      return createResponse({ success: true, data: allData });
    }

    // 3. Aksi Update Status (Untuk Halaman Admin)
    if (action === "updateStatus") {
      var update = handleUpdateStatus(requestData.noReg, requestData.status);
      return createResponse({ success: update });
    }

    // 4. Aksi Cek Pendaftaran (Untuk Halaman Cek Status)
    if (action === "cekPendaftaran") {
      var info = handleCekStatus(requestData.keyword);
      return info ? createResponse({ success: true, data: info }) : createResponse({ success: false });
    }

    return createResponse({ success: false, message: "Aksi tidak dikenal!" });

  } catch (error) {
    return createResponse({ success: false, message: "Server Error: " + error.toString() });
  }
}

function createResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Data_SPMB") || ss.insertSheet("Data_SPMB");
  if (sheet.getLastRow() === 0) {
    var headers = ["Timestamp", "No Registrasi", "Nama Lengkap", "NIK", "NISN", "Tempat Lahir", "Tanggal Lahir", "Jenis Kelamin", "ABK", "Alamat",
                   "Nama Ayah", "Pekerjaan Ayah", "Alamat Ayah", "Gaji Ayah", "HP Ayah",
                   "Nama Ibu", "Pekerjaan Ibu", "Alamat Ibu", "Gaji Ibu", "HP Ibu",
                   "Nama Sekolah", "NPSN", "Alamat Sekolah", "Status"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9ead3");
  }
}

function handleSubmit(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_SPMB");
  var year = new Date().getFullYear();
  var regNo = "SPMB-" + year + "-" + ("0000" + sheet.getLastRow()).slice(-4);

  sheet.appendRow([new Date(), regNo, data.nama, data.nik, data.nisn, data.tempatLahir, data.tanggalLahir, data.jk, data.abk, data.alamat,
                   data.namaAyah, data.pekerjaanAyah, data.alamatAyah, data.gajiAyah, data.hpAyah,
                   data.namaIbu, data.pekerjaanIbu, data.alamatIbu, data.gajiIbu, data.hpIbu,
                   data.namaSekolah, data.npsn, data.alamatSekolah, "MENUNGGU"]);
  return regNo;
}


 function handleGetData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_SPMB");
  var rows = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    result.push({
      noReg: rows[i][1],
      nama: rows[i][2],
      hp: rows[i][13] || rows[i][18], // HP Ayah atau Ibu
      status: rows[i][22]
    });
  }
  return result;
}

function handleUpdateStatus(noReg, status) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_SPMB");
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][1] === noReg) {
      sheet.getRange(i + 1, 24).setValue(status); // Kolom Status sekarang di posisi 24 (index 23)
      return true;
    }
  }
  return false;
}

function handleCekStatus(keyword) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_SPMB");
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    // Cek berdasarkan No Registrasi, NIK, atau NISN
    if (rows[i][1] === keyword || rows[i][3].toString() === keyword || rows[i][4].toString() === keyword) {
      return {
        nama: rows[i][2],
        status: rows[i][23]
      };
    }
  }
  return null;
}
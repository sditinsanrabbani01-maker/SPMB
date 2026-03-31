/**
 * PENGATURAN UTAMA
 * Ganti ID Folder di bawah ini dengan ID folder Google Drive tempat menyimpan berkas
 */
const DRIVE_FOLDER_ID = "1X68-LaYIrVPni2utMXLB5AxGK2VmTGo5"; 

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
    var headers = ["Timestamp", "No Registrasi", "Nama Lengkap", "NISN", "TTL", "Jenis Kelamin", "Alamat", "Data Ayah", "Data Ibu", "WhatsApp", "Asal Sekolah", "Status", "Link Foto", "Link KK", "Link Akta", "Link Bukti Bayar"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9ead3");
  }
}

function handleSubmit(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_SPMB");
  var year = new Date().getFullYear();
  var regNo = "SPMB-" + year + "-" + ("0000" + sheet.getLastRow()).slice(-4);

  // PERBAIKAN: Menambahkan error handling untuk folder Drive
  var folder;
  try {
    folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  } catch (e) {
    // Jika folder tidak ditemukan, gunakan folder root atau buat folder baru
    var rootFolder = DriveApp.getRootFolder();
    folder = rootFolder.createFolder("SPMB_Uploads_" + year) || rootFolder.getFoldersByName("SPMB_Uploads_" + year).next();
  }
  
  var urlFoto = data.foto ? uploadToDrive(data.foto, regNo + "_FOTO", folder) : "";
  var urlKk   = data.kk   ? uploadToDrive(data.kk, regNo + "_KK", folder) : "";
  var urlAkta = data.akta ? uploadToDrive(data.akta, regNo + "_AKTA", folder) : "";
  var urlByr  = data.bayar ? uploadToDrive(data.bayar, regNo + "_BUKTI", folder) : "";

  sheet.appendRow([new Date(), regNo, data.nama, data.nisn, data.ttl, data.jk, data.alamat, data.ayah, data.ibu, data.hp, data.sekolah, "MENUNGGU", urlFoto, urlKk, urlAkta, urlByr]);
  return regNo;
}

function uploadToDrive(base64Data, fileName, folder) {
  try {
    var splitData = base64Data.split(',');
    var contentType = splitData.match(/:(.*?);/); // Ambil tipe konten
    var rawData = Utilities.base64Decode(splitData); // Gunakan data setelah koma
    var blob = Utilities.newBlob(rawData, contentType, fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    return "Error: " + e.toString();
  }
}

 function handleGetData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_SPMB");
  var rows = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    result.push({
      noReg: rows[i][1],
      nama: rows[i][2],
      hp: rows[i][9],
      status: rows[i][11]
    });
  }
  return result;
}

function handleUpdateStatus(noReg, status) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_SPMB");
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][1] === noReg) {
      sheet.getRange(i + 1, 12).setValue(status);
      return true;
    }
  }
  return false;
}

function handleCekStatus(keyword) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_SPMB");
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    // Cek berdasarkan No Registrasi atau NISN
    if (rows[i][1] === keyword || rows[i][3].toString() === keyword) {
      return {
        nama: rows[i][2],
        status: rows[i][11]
      };
    }
  }
  return null;
}
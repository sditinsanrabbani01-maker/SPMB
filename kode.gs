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

    // 5. Aksi Ambil Settings
    if (action === "getSettings") {
      var settings = handleGetSettings();
      return createResponse({ success: true, data: settings });
    }

    // 6. Aksi Update Settings
    if (action === "updateSettings") {
      var update = handleUpdateSettings(requestData.key, requestData.value);
      return createResponse({ success: update });
    }

    // 7. Aksi Get Registrant Details
    if (action === "getRegistrantDetails") {
      var details = handleGetRegistrantDetails(requestData.noReg);
      return createResponse({ success: true, data: details });
    }

    // 8. Aksi Save Payment
    if (action === "savePayment") {
      var save = handleSavePayment(requestData.amount, requestData.name, requestData.status, requestData.image, requestData.filename);
      return createResponse({ success: save });
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

  var settingsSheet = ss.getSheetByName("Settings") || ss.insertSheet("Settings");
  if (settingsSheet.getLastRow() === 0) {
    settingsSheet.appendRow(["Key", "Value"]);
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#d9ead3");
    settingsSheet.appendRow(["registration_open", "true"]);
    settingsSheet.appendRow(["hero_text", "Selamat Datang di SPMB SDIT INSAN RABBANI Online"]);
    settingsSheet.appendRow(["hero_subtext", "Tahun Ajaran 2026/2027"]);
    settingsSheet.appendRow(["hero_quote", "\"Wujudkan Generasi Qurani Sejak Dini Bersama Kami\""]);
  }

  var paymentSheet = ss.getSheetByName("Pembayaran") || ss.insertSheet("Pembayaran");
  if (paymentSheet.getLastRow() === 0) {
    paymentSheet.appendRow(["Timestamp", "Nama", "Jumlah", "Status"]);
    paymentSheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#d9ead3");
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

  // Send WhatsApp notification
  var whatsappNumber = data.hpAyah || data.hpIbu;
  if (whatsappNumber) {
    var message = "Pendaftaran SPMB SDIT Insan Rabbani berhasil!\n\nNo Registrasi: " + regNo + "\nNama: " + data.nama + "\nStatus: MENUNGGU\n\nSilakan cek status pendaftaran secara berkala.";
    sendWhatsAppMessage(whatsappNumber, message);
  }

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
        hp: rows[i][14] || rows[i][19], // HP Ayah atau Ibu
        status: rows[i][23]
      });
    }
    return result;
  }

function handleUpdateStatus(noReg, status) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_SPMB");
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][1] === noReg) {
      var oldStatus = rows[i][23];
      sheet.getRange(i + 1, 24).setValue(status); // Kolom Status sekarang di posisi 24 (index 23)

      // Send WhatsApp if status changed
      if (oldStatus !== status) {
        var whatsappNumber = rows[i][13] || rows[i][18]; // HP Ayah or Ibu
        if (whatsappNumber) {
          var message = "Update Status Pendaftaran SPMB SDIT Insan Rabbani\n\nNo Registrasi: " + noReg + "\nNama: " + rows[i][2] + "\nStatus Baru: " + status + "\n\nTerima kasih.";
          sendWhatsAppMessage(whatsappNumber, message);
        }
      }

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

function handleGetSettings() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var rows = sheet.getDataRange().getValues();
  var settings = {};
  for (var i = 1; i < rows.length; i++) {
    settings[rows[i][0]] = rows[i][1];
  }
  return settings;
}

function handleUpdateSettings(key, value) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return true;
    }
  }
  // If not found, add new
  sheet.appendRow([key, value]);
  return true;
}

function handleGetRegistrantDetails(noReg) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_SPMB");
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][1] === noReg) {
      return {
        timestamp: rows[i][0],
        noReg: rows[i][1],
        nama: rows[i][2],
        nik: rows[i][3],
        nisn: rows[i][4],
        tempatLahir: rows[i][5],
        tanggalLahir: rows[i][6],
        jk: rows[i][7],
        abk: rows[i][8],
        alamat: rows[i][9],
        namaAyah: rows[i][10],
        pekerjaanAyah: rows[i][11],
        alamatAyah: rows[i][12],
        gajiAyah: rows[i][13],
        hpAyah: rows[i][14],
        namaIbu: rows[i][15],
        pekerjaanIbu: rows[i][16],
        alamatIbu: rows[i][17],
        gajiIbu: rows[i][18],
        hpIbu: rows[i][19],
        namaSekolah: rows[i][20],
        npsn: rows[i][21],
        alamatSekolah: rows[i][22],
        status: rows[i][23]
      };
    }
  }
  return null;
}

function handleSavePayment(amount, name, status, image, filename) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pembayaran");
  sheet.appendRow([new Date(), name, amount, status]);

  if (image && filename) {
    // Save image to Drive
    var folder = DriveApp.getFolderById("1X68-LaYIrVPni2utMXLB5AxGK2VmTGo5");
    var fileName = name + " - Konfirmasi." + filename.split('.').pop();
    var blob = Utilities.newBlob(Utilities.base64Decode(image), 'image/jpeg', fileName);
    folder.createFile(blob);
  }

  return true;
}

function normalizeWhatsAppNumber(num) {
  if (!num) return '';
  num = num.toString().replace(/\D/g, '');
  if (num.startsWith('0')) {
    num = '62' + num.substring(1);
  } else if (!num.startsWith('62')) {
    num = '62' + num;
  }
  return num;
}

function sendWhatsAppMessage(number, message) {
  var normalizedNumber = normalizeWhatsAppNumber(number);
  var deviceId = "9b33e3a9-e9ff-4f8b-a62a-90b5eee3f946";
  var url = "https://api.whacenter.com/api/send";
  var payload = {
    device_id: deviceId,
    number: normalizedNumber,
    message: message
  };
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };
  try {
    UrlFetchApp.fetch(url, options);
    return true;
  } catch (e) {
    return false;
  }
}
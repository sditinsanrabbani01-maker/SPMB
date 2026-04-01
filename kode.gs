/**
 * SISTEM SPMB - Google Apps Script
 * Script untuk menangani pendaftaran siswa baru
 */

function doPost(e) {
  Logger.log("=== GAS doPost called ===");
  Logger.log("Raw request: " + e.postData.contents);

  try {
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    Logger.log("Action: " + action);

    setupSheet();

    // 1. Aksi Submit Form Pendaftaran
    if (action === "submitForm") {
      Logger.log("Processing submitForm action");
      var result = handleSubmit(requestData.data);
      Logger.log("submitForm result: " + result);
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
      var save = handleSavePayment(requestData.amount, requestData.name, requestData.status, requestData.image, requestData.filename, requestData.notes);
      return createResponse({ success: save });
    }

    // 9. Aksi Verifikasi Login
    if (action === "verifyLogin") {
      var username = requestData.username;
      var password = requestData.password;
      var settings = handleGetSettings();
      var valid = (settings.admin_user === username && settings.admin_pass === password);
      return createResponse({ success: valid });
    }

    // 10. Aksi Ambil Riwayat Pembayaran
    if (action === "getPaymentHistory") {
      var history = handleGetPaymentHistory();
      return createResponse({ success: true, data: history });
    }

    // 11. Aksi Ambil Kontak CP/Contact Person
    if (action === "getCPContacts") {
      var contacts = handleGetCPContacts();
      return createResponse({ success: true, data: contacts });
    }

    // 12. Aksi Tambah Semua Nomor Ibu ke Group
    if (action === "addAllMothersToGroup") {
      Logger.log("Processing addAllMothersToGroup action");
      var result = addAllMothersToGroup();
      Logger.log("addAllMothersToGroup result: " + JSON.stringify(result));
      return createResponse(result);
    }

    // 13. Aksi Get WhatsApp Groups Info
    if (action === "getWhatsAppGroups") {
      Logger.log("Processing getWhatsAppGroups action");
      var groups = getWhatsAppGroups();
      Logger.log("getWhatsAppGroups result: " + JSON.stringify(groups));
      return createResponse({ success: true, data: groups });
    }

    Logger.log("Unknown action: " + action);
    return createResponse({ success: false, message: "Aksi tidak dikenal!" });
  } catch (error) {
    Logger.log("doPost error: " + error.toString());
    Logger.log("Stack trace: " + error.stack);
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
    // Default admin credentials (can be changed via admin dashboard)
    settingsSheet.appendRow(["admin_user", "admin"]);
    settingsSheet.appendRow(["admin_pass", "admin123"]);
  } else {
    // Ensure admin credentials exist even if sheet already has data
    var rows = settingsSheet.getDataRange().getValues();
    var hasAdminUser = false;
    var hasAdminPass = false;
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === "admin_user") hasAdminUser = true;
      if (rows[i][0] === "admin_pass") hasAdminPass = true;
    }
    if (!hasAdminUser) {
      settingsSheet.appendRow(["admin_user", "admin"]);
    }
    if (!hasAdminPass) {
      settingsSheet.appendRow(["admin_pass", "admin123"]);
    }
  }

  var paymentSheet = ss.getSheetByName("Pembayaran") || ss.insertSheet("Pembayaran");
  if (paymentSheet.getLastRow() === 0) {
    paymentSheet.appendRow(["Tanggal", "Nama", "Nominal", "Konfirmasi", "Catatan"]);
    paymentSheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#d9ead3");
  }

  // CP/Contact Person sheet for WhatsApp notifications
  var cpSheet = ss.getSheetByName("CP/Contact Person") || ss.insertSheet("CP/Contact Person");
  if (cpSheet.getLastRow() === 0) {
    var cpHeaders = ["No Registrasi", "Nama", "Nomor Telepon", "Keterangan"];
    cpSheet.appendRow(cpHeaders);
    cpSheet.getRange(1, 1, 1, cpHeaders.length).setFontWeight("bold").setBackground("#d9ead3");
  }
}

function handleSubmit(data) {
  Logger.log("=== handleSubmit called ===");
  Logger.log("Data received: " + JSON.stringify(data));

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_SPMB");
  var year = new Date().getFullYear();
  var regNo = "SPMB-" + year + "-" + ("0000" + sheet.getLastRow()).slice(-4);
  Logger.log("Generated regNo: " + regNo);

  sheet.appendRow([new Date(), regNo, data.nama, data.nik, data.nisn, data.tempatLahir, data.tanggalLahir, data.jk, data.abk, data.alamat,
                    data.namaAyah, data.pekerjaanAyah, data.alamatAyah, data.gajiAyah, data.hpAyah,
                    data.namaIbu, data.pekerjaanIbu, data.alamatIbu, data.gajiIbu, data.hpIbu,
                    data.namaSekolah, data.npsn, data.alamatSekolah, "MENUNGGU"]);
  Logger.log("Data saved to sheet successfully");

  // Send WhatsApp notification to parent
  var whatsappNumber = data.hpAyah || data.hpIbu;
  if (whatsappNumber) {
    Logger.log("Sending WhatsApp to parent: " + whatsappNumber);
    var message = "Pendaftaran SPMB SDIT Insan Rabbani berhasil!\n\nNo Registrasi: " + regNo + "\nNama: " + data.nama + "\nStatus: MENUNGGU\n\nSilakan cek status pendaftaran secara berkala.";
    var waResult = sendWhatsAppMessage(whatsappNumber, message);
    Logger.log("Parent WhatsApp result: " + waResult);
  } else {
    Logger.log("No WhatsApp number found for parent");
  }

  // Send WhatsApp notification to CP/Contact Person
  var cpNumbers = getCPContactNumbers();
  Logger.log("CP numbers found: " + JSON.stringify(cpNumbers));
  if (cpNumbers.length > 0) {
    var cpMessage = "PENDAFTARAN BARU SPMB SDIT INSAN RABBANI\n\nNo Registrasi: " + regNo + "\nNama: " + data.nama + "\nNama Ayah: " + data.namaAyah + "\nHP Ayah: " + data.hpAyah + "\nNama Ibu: " + data.namaIbu + "\nHP Ibu: " + data.hpIbu + "\nStatus: MENUNGGU\n\nSilakan cek data pendaftaran di dashboard admin.";
    for (var i = 0; i < cpNumbers.length; i++) {
      Logger.log("Sending WhatsApp to CP " + (i+1) + ": " + cpNumbers[i]);
      sendWhatsAppMessage(cpNumbers[i], cpMessage);
    }
  } else {
    Logger.log("No CP numbers found");
  }

  // Add mother's WhatsApp number to group
  if (data.hpIbu) {
    Logger.log("Adding mother to group: " + data.hpIbu);
    addMotherToGroup(data.hpIbu);
  } else {
    Logger.log("No mother's WhatsApp number provided");
  }

  Logger.log("handleSubmit completed successfully");
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

      // Send WhatsApp notification if status changed
      if (oldStatus !== status) {
        Logger.log("Status changed for " + noReg + " from " + oldStatus + " to " + status);
        
        // Send to registrant/parent
        var whatsappNumber = getCPContactNumber(noReg);
        var message = "Status pendaftaran Anda telah berubah.\n\nNo Registrasi: " + noReg + "\nNama: " + rows[i][2] + "\nStatus Sebelumnya: " + oldStatus + "\nStatus Baru: " + status + "\n\nSilakan cek status pendaftaran secara berkala.";
        
        if (whatsappNumber) {
          Logger.log("Sending notification to CP: " + whatsappNumber);
          sendWhatsAppMessage(whatsappNumber, message);
        } else {
          // Fallback to parent's phone number
          var fallbackNumber = rows[i][14] || rows[i][19]; // HP Ayah or Ibu
          if (fallbackNumber) {
            Logger.log("Sending notification to parent: " + fallbackNumber);
            sendWhatsAppMessage(fallbackNumber, message);
          } else {
            Logger.log("No phone number found for registrant: " + noReg);
          }
        }
        
        // Send to CP/Contact Person
        var cpNumbers = getCPContactNumbers();
        if (cpNumbers.length > 0) {
          var cpMessage = "PERUBAHAN STATUS PENDAFTARAN SPMB SDIT INSAN RABBANI\n\nNo Registrasi: " + noReg + "\nNama: " + rows[i][2] + "\nStatus Sebelumnya: " + oldStatus + "\nStatus Baru: " + status + "\n\nSilakan cek data pendaftaran di dashboard admin.";
          for (var j = 0; j < cpNumbers.length; j++) {
            Logger.log("Sending notification to CP list: " + cpNumbers[j]);
            sendWhatsAppMessage(cpNumbers[j], cpMessage);
          }
        } else {
          Logger.log("No CP contacts found in CP/Contact Person sheet");
        }
      }

      return true;
    }
  }
  return false;
}

// Helper function to get WhatsApp number from CP/Contact Person sheet for specific registrant
function getCPContactNumber(noReg) {
  try {
    var cpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CP/Contact Person");
    if (!cpSheet) return null;
    
    var rows = cpSheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === noReg) { // No Registrasi match
        var phone = rows[i][2]; // Nomor Telepon
        if (phone) {
          return normalizeWhatsAppNumber(phone.toString());
        }
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}

// Helper function to get all CP/Contact Person WhatsApp numbers
function getCPContactNumbers() {
  try {
    var cpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CP/Contact Person");
    if (!cpSheet) return [];
    
    var rows = cpSheet.getDataRange().getValues();
    var numbers = [];
    for (var i = 1; i < rows.length; i++) {
      var phone = rows[i][2]; // Nomor Telepon
      if (phone) {
        numbers.push(normalizeWhatsAppNumber(phone.toString()));
      }
    }
    return numbers;
  } catch (e) {
    return [];
  }
}

function handleCekStatus(keyword) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_SPMB");
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    // Cek berdasarkan No Registrasi, NIK, atau NISN
    if (rows[i][1] === keyword || rows[i][3].toString() === keyword || rows[i][4].toString() === keyword) {
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

function handleSavePayment(amount, name, status, image, filename, notes) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pembayaran");
  sheet.appendRow([new Date(), name, amount, status, notes || ""]);

  var driveLink = "";
  if (image && filename) {
    // Save image to Drive in student-specific folder
    var rootFolder = DriveApp.getFolderById("1X68-LaYIrVPni2utMXLB5AxGK2VmTGo5");
    var studentFolderName = name.replace(/[\\\/:*?"<>|]/g, '_'); // Remove invalid folder name characters
    var studentFolder;
    
    // Check if student folder already exists
    var folders = rootFolder.getFoldersByName(studentFolderName);
    if (folders.hasNext()) {
      studentFolder = folders.next();
    } else {
      // Create new student folder
      studentFolder = rootFolder.createFolder(studentFolderName);
    }
    
    // Rename file to "Konfirmasi Pembayaran - Nama Siswa"
    var fileExtension = filename.split('.').pop();
    var fileName = "Konfirmasi Pembayaran - " + name + "." + fileExtension;
    var blob = Utilities.newBlob(Utilities.base64Decode(image), 'image/jpeg', fileName);
    var file = studentFolder.createFile(blob);
    
    // Get the Drive link
    driveLink = file.getUrl();
  }

  return { success: true, driveLink: driveLink };
}

function handleGetCPContacts() {
  try {
    var cpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CP/Contact Person");
    if (!cpSheet) {
      Logger.log("CP/Contact Person sheet not found");
      return [];
    }
    
    var rows = cpSheet.getDataRange().getValues();
    Logger.log("CP sheet rows: " + rows.length);
    var contacts = [];
    for (var i = 1; i < rows.length; i++) {
      Logger.log("Row " + i + ": " + JSON.stringify(rows[i]));
      if (rows[i][2]) { // Nomor Telepon
        contacts.push({
          name: rows[i][1] || "CP " + i,
          phone: rows[i][2].toString(),
          description: rows[i][3] || ""
        });
      }
    }
    Logger.log("Found " + contacts.length + " CP contacts");
    return contacts;
  } catch (e) {
    Logger.log("Error in handleGetCPContacts: " + e.toString());
    return [];
  }
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
    Logger.log("Sending WhatsApp to " + normalizedNumber + ": " + message.substring(0, 50) + "...");
    var response = UrlFetchApp.fetch(url, options);
    Logger.log("WhatsApp sent successfully to " + normalizedNumber);
    return true;
  } catch (e) {
    Logger.log("Error sending WhatsApp to " + normalizedNumber + ": " + e.toString());
    return false;
  }
}

// Function to get group list from WhaCenter
function getWhatsAppGroups() {
  var deviceId = "9b33e3a9-e9ff-4f8b-a62a-90b5eee3f946";
  var apiKey = "9b33e3a9-e9ff-4f8b-a62a-90b5eee3f946"; // Same as device ID
  var url = "https://api.whacenter.com/api/getGroup?device_id=" + deviceId + "&api_key=" + apiKey;
  var options = {
    method: "get",
    muteHttpExceptions: true
  };
  try {
    Logger.log("Getting WhatsApp groups from: " + url.replace(apiKey, "***"));
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    Logger.log("Response code: " + responseCode);
    Logger.log("Response text: " + responseText);

    if (responseCode === 200) {
      var result = JSON.parse(responseText);
      Logger.log("Groups parsed successfully: " + JSON.stringify(result));

      // Normalize the response to match expected format
      // API returns: {"success":true,"data":{"groups":[...]},"message":"Success"}
      // We need: {"Status":true,"Data":{"groups":[...]},"Message":"Success"}
      if (result.success) {
        return {
          Status: true,
          Data: result.data,
          Message: result.message || "Success"
        };
      } else {
        return {
          Status: false,
          Message: result.message || "API returned success=false"
        };
      }
    } else {
      Logger.log("API returned error code: " + responseCode + ", response: " + responseText);
      return { Status: false, Message: "API Error: " + responseCode + " - " + responseText };
    }
  } catch (e) {
    Logger.log("Error getting groups: " + e.toString());
    return { Status: false, Message: "Network Error: " + e.toString() };
  }
}

// Function to add numbers to a WhatsApp group
function addNumbersToGroup(rawGroupId, numbers, maxRetries = 2) {
  Logger.log("=== addNumbersToGroup called ===");
  Logger.log("Raw group ID: " + rawGroupId);
  Logger.log("Numbers to add: " + JSON.stringify(numbers));

  // Clean group ID: remove "@g.us" suffix and keep as string
  var cleanGroupId = rawGroupId.replace('@g.us', '');
  Logger.log("Cleaned group ID: " + cleanGroupId);

  // Normalize phone numbers
  var normalizedNumbers = numbers.map(function(num) {
    return normalizeWhatsAppNumber(num.toString());
  });
  Logger.log("Normalized numbers: " + JSON.stringify(normalizedNumbers));

  // TODO: Replace with your actual API key from WhaCenter dashboard
  var apiKey = "9b33e3a9-e9ff-4f8b-a62a-90b5eee3f946"; // This is DEVICE ID, not API key!

  var url = "https://api.whacenter.com/api/addNumberToGroup";

  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    Logger.log("Attempt " + attempt + " of " + maxRetries);

    var payload = {
      api_key: apiKey,
      group_id: cleanGroupId, // String format without @g.us
      data: normalizedNumbers
    };

    Logger.log("Payload: " + JSON.stringify(payload));

    var options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      var response = UrlFetchApp.fetch(url, options);
      var responseCode = response.getResponseCode();
      var responseText = response.getContentText();

      Logger.log("Response code: " + responseCode);
      Logger.log("Response text: " + responseText);

      if (responseCode === 200) {
        var result = JSON.parse(responseText);
        Logger.log("Parsed result: " + JSON.stringify(result));

        if (result.success === true) {
          Logger.log("✅ SUCCESS: Numbers added to group");
          return {
            Status: true,
            Message: result.message || "Numbers added successfully"
          };
        } else {
          Logger.log("❌ API returned success=false: " + result.message);

          // Check for common error types
          if (result.message && result.message.toLowerCase().includes("not admin")) {
            Logger.log("🔒 Permission Error: Bot is not admin of the group");
            return {
              Status: false,
              Message: "Bot is not admin of the group. Make sure the WhaCenter device is added as admin."
            };
          }

          if (result.message && result.message.toLowerCase().includes("invalid api key")) {
            Logger.log("🔑 API Key Error: Invalid API key");
            return {
              Status: false,
              Message: "Invalid API key. Check your WhaCenter API key in dashboard."
            };
          }

          if (result.message && result.message.toLowerCase().includes("invalid group")) {
            Logger.log("👥 Group Error: Invalid or inaccessible group");
            return {
              Status: false,
              Message: "Invalid group ID or group not accessible to the device."
            };
          }

          return {
            Status: false,
            Message: result.message || "Unknown API error"
          };
        }
      } else if (responseCode === 401) {
        Logger.log("🔑 Authentication Error (401): Invalid API key");
        return {
          Status: false,
          Message: "Invalid API key. Please check your WhaCenter API key."
        };
      } else if (responseCode === 403) {
        Logger.log("🚫 Permission Error (403): Not authorized");
        return {
          Status: false,
          Message: "Not authorized. Check device permissions in WhaCenter."
        };
      } else if (responseCode === 404) {
        Logger.log("🔍 Not Found (404): Group or endpoint not found");
        return {
          Status: false,
          Message: "Group not found or API endpoint incorrect."
        };
      } else {
        Logger.log("⚠️ HTTP Error: " + responseCode);
        if (attempt < maxRetries) {
          Logger.log("Retrying in 2 seconds...");
          Utilities.sleep(2000); // Wait 2 seconds before retry
          continue;
        }
        return {
          Status: false,
          Message: "HTTP " + responseCode + ": " + responseText
        };
      }
    } catch (e) {
      Logger.log("Network error on attempt " + attempt + ": " + e.toString());
      if (attempt < maxRetries) {
        Logger.log("Retrying in 2 seconds...");
        Utilities.sleep(2000);
        continue;
      }
      return {
        Status: false,
        Message: "Network Error: " + e.toString()
      };
    }
  }

  Logger.log("❌ All retry attempts failed");
  return {
    Status: false,
    Message: "Failed after " + maxRetries + " attempts"
  };
}

// Function to add mother's WhatsApp number to group when new registrant is submitted
function addMotherToGroup(hpIbu) {
  if (!hpIbu) return;

  // Get group ID for "SPMB 2026/2027"
  var groups = getWhatsAppGroups();
  if (groups && groups.Status && groups.Data && groups.Data.groups) {
    var spmbGroup = groups.Data.groups.find(function(g) {
      return g.name === "SPMB 2026/2027";
    });

    if (spmbGroup) {
      var normalizedNumber = normalizeWhatsAppNumber(hpIbu);
      var numbers = [normalizedNumber];
      var result = addNumbersToGroup(spmbGroup.id, numbers);
      Logger.log("Added new registrant to group: " + normalizedNumber + ", result: " + JSON.stringify(result));
    } else {
      Logger.log("Group 'SPMB 2026/2027' not found - cannot add new registrant");
    }
  }
}

// Function to add all existing mother's WhatsApp numbers to group
function addAllMothersToGroup() {
  Logger.log("=== addAllMothersToGroup called ===");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_SPMB");
  var rows = sheet.getDataRange().getValues();
  var numbers = [];

  Logger.log("Processing " + (rows.length - 1) + " rows from Data_SPMB sheet");

  for (var i = 1; i < rows.length; i++) {
    var hpIbu = rows[i][19]; // HP Ibu column
    if (hpIbu) {
      var normalizedNumber = normalizeWhatsAppNumber(hpIbu.toString());
      Logger.log("Row " + i + " - Original: " + hpIbu + ", Normalized: " + normalizedNumber);
      numbers.push(normalizedNumber);
    } else {
      Logger.log("Row " + i + " - No phone number found");
    }
  }

  Logger.log("Found " + numbers.length + " phone numbers to add to group: " + JSON.stringify(numbers));

  if (numbers.length > 0) {
    // Get group ID for "SPMB 2026/2027"
    Logger.log("Fetching WhatsApp groups...");
    var groups = getWhatsAppGroups();
    Logger.log("Groups response: " + JSON.stringify(groups));

    if (groups && groups.Status && groups.Data && groups.Data.groups) {
      Logger.log("Available groups:");
      for (var i = 0; i < groups.Data.groups.length; i++) {
        Logger.log("Group " + (i+1) + ": " + groups.Data.groups[i].name + " (ID: " + groups.Data.groups[i].id + ", Participants: " + groups.Data.groups[i].participants + ")");
      }

      var spmbGroup = groups.Data.groups.find(function(g) {
        return g.name === "SPMB 2026/2027";
      });

      if (spmbGroup) {
        Logger.log("Found SPMB group with ID: " + spmbGroup.id + ", trying to add " + numbers.length + " numbers");
        var result = addNumbersToGroup(spmbGroup.id, numbers);

        if (result && result.Status) {
          Logger.log("✅ Successfully added " + numbers.length + " numbers to group");
          return { success: true, message: "Berhasil menambahkan " + numbers.length + " nomor ke group WhatsApp" };
        } else {
          Logger.log("❌ Failed to add numbers to group: " + (result ? result.Message : "Unknown error"));
          return { success: false, message: result ? result.Message : "Gagal menambahkan nomor ke group WhatsApp" };
        }
      } else {
        Logger.log("Group 'SPMB 2026/2027' not found in available groups");
        return { success: false, message: "Group 'SPMB 2026/2027' tidak ditemukan. Pastikan group sudah dibuat di WhatsApp dan device terhubung." };
      }
    } else {
      Logger.log("Failed to get groups or no groups found. Groups object: " + JSON.stringify(groups));
      return { success: false, message: "Gagal mendapatkan daftar group WhatsApp. Response: " + JSON.stringify(groups) };
    }
  } else {
    Logger.log("No phone numbers found to add");
    return { success: false, message: "Tidak ada nomor telepon yang ditemukan di database" };
  }
}

function sendEmailNotification(to, subject, body) {
  try {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: body
    });
    return true;
  } catch (e) {
    return false;
  }
}

function handleGetPaymentHistory() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pembayaran");
  var rows = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    result.push({
      timestamp: rows[i][0],
      name: rows[i][1],
      amount: rows[i][2],
      status: rows[i][3],
      notes: rows[i][4] || ""
    });
  }
  return result;
}

// === TEST FUNCTIONS - Run these directly in GAS Editor ===

// Test function to check WhatsApp groups
function testGetWhatsAppGroups() {
  Logger.log("=== Testing getWhatsAppGroups ===");
  var result = getWhatsAppGroups();
  Logger.log("Result: " + JSON.stringify(result));
  return result;
}

// Test function to add all mothers to group
function testAddAllMothersToGroup() {
  Logger.log("=== Testing addAllMothersToGroup ===");
  var result = addAllMothersToGroup();
  Logger.log("Result: " + JSON.stringify(result));
  return result;
}

// Test function to check if SPMB group exists
function testCheckSPMBGroup() {
  Logger.log("=== Testing SPMB Group Check ===");
  var groups = getWhatsAppGroups();

  Logger.log("Groups response: " + JSON.stringify(groups));

  if (groups && groups.Status && groups.Data && groups.Data.groups) {
    Logger.log("Available groups:");
    var spmbFound = false;
    for (var i = 0; i < groups.Data.groups.length; i++) {
      var group = groups.Data.groups[i];
      Logger.log("Group " + (i+1) + ": " + group.name + " (ID: " + group.id + ", Participants: " + (group.participants || "N/A") + ")");

      if (group.name === "SPMB 2026/2027") {
        spmbFound = true;
        Logger.log("✓ SPMB group found with ID: " + group.id);
      }
    }

    if (!spmbFound) {
      Logger.log("✗ SPMB group 'SPMB 2026/2027' not found!");
      Logger.log("Please create the group in WhatsApp first.");
    }
  } else {
    Logger.log("✗ Failed to get groups or invalid response format: " + JSON.stringify(groups));
  }

  return groups;
}

// Test function to check current data in sheets
function testCheckDataSheets() {
  Logger.log("=== Testing Data Sheets ===");

  // Check Data_SPMB sheet
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_SPMB");
  if (dataSheet) {
    var dataRows = dataSheet.getDataRange().getValues();
    Logger.log("Data_SPMB sheet has " + (dataRows.length - 1) + " data rows");

    var phoneCount = 0;
    for (var i = 1; i < dataRows.length; i++) {
      var hpIbu = dataRows[i][19]; // HP Ibu column
      if (hpIbu) {
        var normalized = normalizeWhatsAppNumber(hpIbu.toString());
        Logger.log("Row " + i + " - Mother phone: " + hpIbu + " → " + normalized);
        phoneCount++;
      }
    }
    Logger.log("Found " + phoneCount + " mother phone numbers");
  } else {
    Logger.log("✗ Data_SPMB sheet not found");
  }

  // Check Pembayaran sheet
  var paySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pembayaran");
  if (paySheet) {
    var payRows = paySheet.getDataRange().getValues();
    Logger.log("Pembayaran sheet has " + (payRows.length - 1) + " payment rows");
  } else {
    Logger.log("✗ Pembayaran sheet not found");
  }

  // Check Settings sheet
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  if (settingsSheet) {
    var settings = handleGetSettings();
    Logger.log("Settings: " + JSON.stringify(settings));
  } else {
    Logger.log("✗ Settings sheet not found");
  }
}

// Test function to simulate a form submission
function testSimulateFormSubmission() {
  Logger.log("=== Testing Form Submission Simulation ===");

  var testData = {
    nama: "Test Siswa",
    nik: "1234567890123456",
    nisn: "1234567890",
    tempatLahir: "Jakarta",
    tanggalLahir: "2015-01-01",
    jk: "L",
    abk: "Tidak",
    alamat: "Jl. Test No. 123",
    namaAyah: "Test Ayah",
    pekerjaanAyah: "Pegawai",
    alamatAyah: "Jl. Test No. 123",
    gajiAyah: "5000000",
    hpAyah: "081234567890",
    namaIbu: "Test Ibu",
    pekerjaanIbu: "Ibu Rumah Tangga",
    alamatIbu: "Jl. Test No. 123",
    gajiIbu: "0",
    hpIbu: "081234567891",
    namaSekolah: "SD Test",
    npsn: "12345678",
    alamatSekolah: "Jl. Sekolah Test"
  };

  Logger.log("Submitting test data...");
  var result = handleSubmit(testData);
  Logger.log("Test submission result: " + result);

  return result;
}

// Test function to try different API keys
function testAPIKeys() {
  Logger.log("=== Testing Different API Keys ===");

  var testKeys = [
    "9b33e3a9-e9ff-4f8b-a62a-90b5eee3f946", // Current device ID
    "your_actual_api_key_here", // Replace with real API key from WhaCenter
    // Add more keys to test here
  ];

  var groupId = "120363407082974441@g.us";
  var testNumbers = ["6282196910108"];

  for (var i = 0; i < testKeys.length; i++) {
    Logger.log("Testing API key " + (i+1) + ": " + testKeys[i]);

    try {
      var url = "https://api.whacenter.com/api/addNumberToGroup";
      var payload = {
        api_key: testKeys[i],
        group_id: groupId,
        data: testNumbers
      };
      var options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      var response = UrlFetchApp.fetch(url, options);
      var responseCode = response.getResponseCode();
      var responseText = response.getContentText();

      Logger.log("API Key " + testKeys[i] + " - Response code: " + responseCode);
      Logger.log("Response: " + responseText);

      if (responseCode === 200) {
        var result = JSON.parse(responseText);
        if (result.success) {
          Logger.log("✅ SUCCESS! API key " + testKeys[i] + " works!");
          return testKeys[i];
        }
      }
    } catch (e) {
      Logger.log("API Key " + testKeys[i] + " - Error: " + e.toString());
    }
  }

  Logger.log("❌ No working API key found. Please check your WhaCenter account for the correct API key.");
  return null;
}
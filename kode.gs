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
      var result = addAllMothersToGroup();
      return createResponse(result);
    }

    // 13. Aksi Get WhatsApp Groups Info
    if (action === "getWhatsAppGroups") {
      var groups = getWhatsAppGroups();
      return createResponse({ success: true, data: groups });
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_SPMB");
  var year = new Date().getFullYear();
  var regNo = "SPMB-" + year + "-" + ("0000" + sheet.getLastRow()).slice(-4);

  sheet.appendRow([new Date(), regNo, data.nama, data.nik, data.nisn, data.tempatLahir, data.tanggalLahir, data.jk, data.abk, data.alamat,
                    data.namaAyah, data.pekerjaanAyah, data.alamatAyah, data.gajiAyah, data.hpAyah,
                    data.namaIbu, data.pekerjaanIbu, data.alamatIbu, data.gajiIbu, data.hpIbu,
                    data.namaSekolah, data.npsn, data.alamatSekolah, "MENUNGGU"]);

  // Send WhatsApp notification to parent
  var whatsappNumber = data.hpAyah || data.hpIbu;
  if (whatsappNumber) {
    var message = "Pendaftaran SPMB SDIT Insan Rabbani berhasil!\n\nNo Registrasi: " + regNo + "\nNama: " + data.nama + "\nStatus: MENUNGGU\n\nSilakan cek status pendaftaran secara berkala.";
    sendWhatsAppMessage(whatsappNumber, message);
  }
  
  // Send WhatsApp notification to CP/Contact Person
  var cpNumbers = getCPContactNumbers();
  if (cpNumbers.length > 0) {
    var cpMessage = "PENDAFTARAN BARU SPMB SDIT INSAN RABBANI\n\nNo Registrasi: " + regNo + "\nNama: " + data.nama + "\nNama Ayah: " + data.namaAyah + "\nHP Ayah: " + data.hpAyah + "\nNama Ibu: " + data.namaIbu + "\nHP Ibu: " + data.hpIbu + "\nStatus: MENUNGGU\n\nSilakan cek data pendaftaran di dashboard admin.";
    for (var i = 0; i < cpNumbers.length; i++) {
      sendWhatsAppMessage(cpNumbers[i], cpMessage);
    }
  }
  
  // Add mother's WhatsApp number to group
  if (data.hpIbu) {
    addMotherToGroup(data.hpIbu);
  }
  
  // Send Email notification
  var emailTo = data.hpAyah || data.hpIbu; // Using WhatsApp number as placeholder - in real scenario, we'd have email field
  // For now, we'll skip email if no email field exists in the form
  // In a production system, you would add email fields to the form and use them here
  // var emailTo = data.emailAyah || data.emailIbu;
  // if (emailTo) {
  //   var emailSubject = "Pendaftaran SPMB SDIT Insan Rabbani Berhasil";
  //   var emailBody = "Yth. Bapak/Ibu " + data.namaAyah + ",\n\n" +
  //                   "Pendaftaran anak Anda atas nama " + data.nama + " telah berhasil diterima.\n\n" +
  //                   "No Registrasi: " + regNo + "\n" +
  //                   "Nama: " + data.nama + "\n" +
  //                   "Status: MENUNGGU\n\n" +
  //                   "Silakan cek status pendaftaran secara berkala melalui halaman cek pendaftaran.\n\n" +
  //                   "Terima kasih.\n" +
  //                   "Tim SPMB SDIT Insan Rabbani";
  //   sendEmailNotification(emailTo, emailSubject, emailBody);
  // }

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
      return result;
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
function addNumbersToGroup(groupId, numbers) {
  var apiKey = "9b33e3a9-e9ff-4f8b-a62a-90b5eee3f946"; // Device ID as API key
  var url = "https://api.whacenter.com/api/addNumberToGroup";
  var payload = {
    api_key: apiKey,
    group_id: groupId,
    data: numbers
  };
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    Logger.log("Adding numbers to group: " + groupId + ", numbers: " + JSON.stringify(numbers));
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    Logger.log("Add to group response code: " + responseCode);
    Logger.log("Add to group response text: " + responseText);

    if (responseCode === 200) {
      var result = JSON.parse(responseText);
      Logger.log("Add to group result parsed: " + JSON.stringify(result));
      return result;
    } else {
      Logger.log("Add to group API error: " + responseCode + " - " + responseText);
      return { Status: false, Message: "API Error: " + responseCode + " - " + responseText };
    }
  } catch (e) {
    Logger.log("Error adding numbers to group: " + e.toString());
    return { Status: false, Message: "Network Error: " + e.toString() };
  }
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_SPMB");
  var rows = sheet.getDataRange().getValues();
  var numbers = [];

  for (var i = 1; i < rows.length; i++) {
    var hpIbu = rows[i][19]; // HP Ibu column
    if (hpIbu) {
      var normalizedNumber = normalizeWhatsAppNumber(hpIbu.toString());
      numbers.push(normalizedNumber);
    }
  }

  Logger.log("Found " + numbers.length + " phone numbers to add to group");

  if (numbers.length > 0) {
    // Get group ID for "SPMB 2026/2027"
    var groups = getWhatsAppGroups();
    Logger.log("Groups response: " + JSON.stringify(groups));

    if (groups && groups.Status && groups.Data && groups.Data.groups) {
      Logger.log("Available groups:");
      for (var i = 0; i < groups.Data.groups.length; i++) {
        Logger.log("Group: " + groups.Data.groups[i].name + " (ID: " + groups.Data.groups[i].id + ")");
      }

      var spmbGroup = groups.Data.groups.find(function(g) {
        return g.name === "SPMB 2026/2027";
      });

      if (spmbGroup) {
        Logger.log("Found SPMB group with ID: " + spmbGroup.id);
        var result = addNumbersToGroup(spmbGroup.id, numbers);
        Logger.log("Add to group result: " + JSON.stringify(result));
        if (result && result.Status) {
          Logger.log("Successfully added " + numbers.length + " numbers to group");
          return { success: true, message: "Berhasil menambahkan " + numbers.length + " nomor ke group WhatsApp" };
        } else {
          Logger.log("Failed to add numbers to group");
          return { success: false, message: "Gagal menambahkan nomor ke group WhatsApp" };
        }
      } else {
        Logger.log("Group 'SPMB 2026/2027' not found");
        return { success: false, message: "Group 'SPMB 2026/2027' tidak ditemukan. Pastikan group sudah dibuat di WhatsApp." };
      }
    } else {
      Logger.log("Failed to get groups or no groups found");
      return { success: false, message: "Gagal mendapatkan daftar group WhatsApp" };
    }
  } else {
    Logger.log("No phone numbers found to add");
    return { success: false, message: "Tidak ada nomor telepon yang ditemukan" };
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
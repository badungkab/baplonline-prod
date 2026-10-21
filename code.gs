/* ============================================================
   KONFIGURASI - GANTI DENGAN ID ANDA
   ============================================================ */
const SHEET_ID = "GANTI_DENGAN_ID_SHEET_ANDA";
const FOLDER_ID = "GANTI_DENGAN_ID_FOLDER_DRIVE_ANDA";
const SHEET_PENGEcer = "DataPengecer";
const SHEET_RESTORAN = "DataRestoran";
const SHEET_HASIL = "HasilBAPL";

// Fungsi untuk trigger izin akses Drive - jalankan 1x saja
function setupDrivePermissions() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  Logger.log('✅ Akses Drive berhasil! Folder: ' + folder.getName());
  
  const testBlob = Utilities.newBlob('test', 'text/plain', 'test.txt');
  const testFile = folder.createFile(testBlob);
  Logger.log('✅ Test file created: ' + testFile.getName());
  
  try {
    testFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log('✅ Sharing set berhasil');
  } catch (e) {
    Logger.log('⚠️ Warning: setSharing gagal - ' + e.message);
  }
  
  testFile.setTrashed(true);
  Logger.log('✅ Setup selesai!');
}

function doGet(e) {
  const action = e.parameter.action || "read";
  
  if (action === "read") {
    const pengecer = getAllData(SHEET_PENGEcer);
    const restoran = getAllData(SHEET_RESTORAN);
    const allData = [...pengecer, ...restoran].sort((a, b) => {
      return new Date(b.tanggal) - new Date(a.tanggal);
    });
    return jsonResponse({ success: true, data: allData });
  }
  
  if (action === "readHasil") {
    const allHasil = getAllHasil();
    return jsonResponse({ success: true, data: allHasil });
  }
  
  if (action === "getFile") {
    const fileId = e.parameter.fileId;
    if (!fileId) return jsonResponse({ success: false, message: "File ID kosong" });
    try {
      const file = DriveApp.getFileById(fileId);
      return jsonResponse({ success: true, url: file.getUrl(), name: file.getName() });
    } catch (err) {
      return jsonResponse({ success: false, message: "File tidak ditemukan" });
    }
  }
  
  return jsonResponse({ success: false, message: "Action tidak dikenal" });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const { action, data } = body;
    let result = {};

    switch(action) {
      case "createPengecer":
        result = { success: true, message: "✅ Data Pengecer berhasil disimpan", id: createData(SHEET_PENGEcer, data) };
        break;
      case "createRestoran":
        result = { success: true, message: "✅ Data Restoran/Bar/Hotel berhasil disimpan", id: createData(SHEET_RESTORAN, data) };
        break;
      case "createHasil":
        result = { success: true, message: "✅ Dokumen hasil berhasil dipublikasikan", id: createHasil(data) };
        break;
      case "uploadFile":
        result = uploadFile(data);
        break;
      case "uploadFileHasil":
        result = uploadFileHasil(data);
        break;
      case "deleteFile":
        result = deleteFile(data.fileId);
        break;
      default:
        throw new Error("Action tidak valid");
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

/* ====== UPLOAD FILE (untuk dokumen persyaratan) ====== */
function uploadFile(data) {
  try {
    const { base64, fileName, folderName, noPermohonan } = data;
    
    if (!base64 || !fileName) {
      throw new Error("File tidak valid");
    }
    
    const contentType = base64.split(';')[0].split(':')[1];
    const base64Data = base64.split(',')[1];
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, fileName);
    
    let folder;
    const folders = DriveApp.getFolderById(FOLDER_ID).getFoldersByName(folderName);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.getFolderById(FOLDER_ID).createFolder(folderName);
    }
    
    let subFolder;
    const subFolders = folder.getFoldersByName(noPermohonan);
    if (subFolders.hasNext()) {
      subFolder = subFolders.next();
    } else {
      subFolder = folder.createFolder(noPermohonan);
    }
    
    const file = subFolder.createFile(blob);
    
    // Coba setSharing, jika gagal tetap lanjutkan
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log('Warning: setSharing gagal untuk file ' + fileName + ': ' + shareErr.message);
    }
    
    return {
      success: true,
      fileId: file.getId(),
      fileName: fileName,
      fileUrl: file.getUrl(),
      downloadUrl: "https://drive.google.com/uc?id=" + file.getId() + "&export=download",
      message: "✅ File berhasil diupload"
    };
  } catch (err) {
    return { success: false, message: "Gagal upload: " + err.message };
  }
}

/* ====== UPLOAD FILE HASIL (untuk dokumen BAPL selesai) ====== */
function uploadFileHasil(data) {
  try {
    const { base64, fileName, folderName, noPermohonan } = data;
    
    if (!base64 || !fileName) {
      throw new Error("File tidak valid");
    }
    
    const contentType = base64.split(';')[0].split(':')[1];
    const base64Data = base64.split(',')[1];
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, fileName);
    
    let folder;
    const folders = DriveApp.getFolderById(FOLDER_ID).getFoldersByName(folderName);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.getFolderById(FOLDER_ID).createFolder(folderName);
    }
    
    const file = folder.createFile(blob);
    
    // Coba setSharing, jika gagal tetap lanjutkan
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log('Warning: setSharing gagal untuk file ' + fileName + ': ' + shareErr.message);
    }
    
    return {
      success: true,
      fileId: file.getId(),
      fileName: fileName,
      fileUrl: file.getUrl(),
      downloadUrl: "https://drive.google.com/uc?id=" + file.getId() + "&export=download",
      message: "✅ File hasil berhasil diupload"
    };
  } catch (err) {
    return { success: false, message: "Gagal upload: " + err.message };
  }
}

function deleteFile(fileId) {
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    return { success: true, message: "✅ File berhasil dihapus" };
  } catch (err) {
    return { success: false, message: "Gagal menghapus file" };
  }
}

/* ====== CRUD DATA PERMOHONAN ====== */
const DATA_HEADERS = [
  "ID", "Tanggal", "Jenis", "NamaPerusahaan", "NoPermohonan",
  "Status", "BentukUsaha", "Golongan", "Alamat", "Telp", "Email",
  "NPWP", "BPJS", "Kegiatan", "NamaUsaha", "LokasiUsaha",
  "PenanggungJawab", "MinumanA", "MinumanB", "MinumanC",
  "JarakIbadah", "JarakSekolah", "JarakRS", "Latitude", "Longitude",
  "FileNIB", "FileNPWP", "FileKTP", "FileNPPBKC", "FilePakta",
  "FileSuratPenunjukan", "FileSertifikat", "FileKoordinat"
];

function ensureHeaders(sh, headers) {
  const lastColumn = Math.max(sh.getLastColumn(), 1);
  let current = sh.getRange(1, 1, 1, lastColumn).getValues()[0];
  if (!current[0]) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return headers;
  }
  const missing = headers.filter(h => current.indexOf(h) === -1);
  if (missing.length) {
    sh.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
    current = current.concat(missing);
  }
  return current;
}

function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1, 1, 1, DATA_HEADERS.length).setValues([DATA_HEADERS]);
  } else {
    ensureHeaders(sh, DATA_HEADERS);
  }
  return sh;
}

function getHeaderMap(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i; });
  return map;
}

function getAllData(sheetName) {
  const sh = getSheet(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const map = getHeaderMap(sh);
  const get = (row, name) => map[name] !== undefined ? (row[map[name]] ?? "") : "";

  return values.slice(1).map(row => ({
    id: get(row, "ID"),
    tanggal: get(row, "Tanggal"),
    jenis: get(row, "Jenis"),
    namaPerusahaan: get(row, "NamaPerusahaan"),
    noPermohonan: get(row, "NoPermohonan"),
    status: get(row, "Status"),
    bentukUsaha: get(row, "BentukUsaha"),
    golongan: get(row, "Golongan"),
    alamat: get(row, "Alamat"),
    telp: get(row, "Telp"),
    email: get(row, "Email"),
    npwp: get(row, "NPWP"),
    bpjs: get(row, "BPJS"),
    kegiatan: get(row, "Kegiatan"),
    namaUsaha: get(row, "NamaUsaha"),
    lokasiUsaha: get(row, "LokasiUsaha"),
    penanggungJawab: get(row, "PenanggungJawab"),
    minumanA: get(row, "MinumanA"),
    minumanB: get(row, "MinumanB"),
    minumanC: get(row, "MinumanC"),
    jarakIbadah: get(row, "JarakIbadah"),
    jarakSekolah: get(row, "JarakSekolah"),
    jarakRS: get(row, "JarakRS"),
    latitude: get(row, "Latitude"),
    longitude: get(row, "Longitude"),
    fileNIB: get(row, "FileNIB"),
    fileNPWP: get(row, "FileNPWP"),
    fileKTP: get(row, "FileKTP"),
    fileNPPBKC: get(row, "FileNPPBKC"),
    filePakta: get(row, "FilePakta"),
    fileSuratPenunjukan: get(row, "FileSuratPenunjukan"),
    fileSertifikat: get(row, "FileSertifikat"),
    fileKoordinat: get(row, "FileKoordinat")
  }));
}

function createData(sheetName, d) {
  const sh = getSheet(sheetName);
  const id = "BAPL-" + new Date().getTime();
  const values = DATA_HEADERS.map(h => {
    if (h === "ID") return id;
    const keyMap = {
      Tanggal: "tanggal", Jenis: "jenis", NamaPerusahaan: "namaPerusahaan", NoPermohonan: "noPermohonan",
      Status: "status", BentukUsaha: "bentukUsaha", Golongan: "golongan", Alamat: "alamat", Telp: "telp",
      Email: "email", NPWP: "npwp", BPJS: "bpjs", Kegiatan: "kegiatan", NamaUsaha: "namaUsaha",
      LokasiUsaha: "lokasiUsaha", PenanggungJawab: "penanggungJawab", MinumanA: "minumanA", MinumanB: "minumanB",
      MinumanC: "minumanC", JarakIbadah: "jarakIbadah", JarakSekolah: "jarakSekolah", JarakRS: "jarakRS",
      Latitude: "latitude", Longitude: "longitude", FileNIB: "fileNIB", FileNPWP: "fileNPWP", FileKTP: "fileKTP",
      FileNPPBKC: "fileNPPBKC", FilePakta: "filePakta", FileSuratPenunjukan: "fileSuratPenunjukan",
      FileSertifikat: "fileSertifikat", FileKoordinat: "fileKoordinat"
    };
    const key = keyMap[h];
    return key ? (d[key] || "") : "";
  });
  const map = getHeaderMap(sh);
  const row = new Array(sh.getLastColumn()).fill("");
  DATA_HEADERS.forEach((h, i) => { if (map[h] !== undefined) row[map[h]] = values[i]; });
  sh.appendRow(row);
  return id;
}

/* ====== CRUD HASIL BAPL ====== */
function getSheetHasil() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_HASIL);
  if (!sh) {
    sh = ss.insertSheet(SHEET_HASIL);
    sh.appendRow([
      "ID", "NoPermohonan", "NamaUsaha", "Jenis", "Golongan",
      "TanggalTerbit", "Keterangan", "FileHasilId", "FileHasilName",
      "FileHasilUrl", "FileHasilDownloadUrl", "TanggalUpload", "UploadBy"
    ]);
  }
  return sh;
}

function getAllHasil() {
  const sh = getSheetHasil();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  
  return values.slice(1).map(row => ({
    id: row[0],
    noPermohonan: row[1],
    namaUsaha: row[2],
    jenis: row[3],
    golongan: row[4],
    tanggalTerbit: row[5] ? Utilities.formatDate(new Date(row[5]), "Asia/Makassar", "yyyy-MM-dd") : "",
    keterangan: row[6] || "",
    fileHasilId: row[7] || "",
    fileHasilName: row[8] || "",
    fileHasilUrl: row[9] || "",
    fileHasilDownloadUrl: row[10] || "",
    tanggalUpload: row[11] ? Utilities.formatDate(new Date(row[11]), "Asia/Makassar", "dd/MM/yyyy HH:mm") : ""
  })).sort((a, b) => new Date(b.tanggalTerbit) - new Date(a.tanggalTerbit));
}

function createHasil(d) {
  const sh = getSheetHasil();
  const id = "HASIL-" + new Date().getTime();
  sh.appendRow([
    id, d.noPermohonan, d.namaUsaha, d.jenis, d.golongan,
    d.tanggalTerbit, d.keterangan, d.fileHasilId, d.fileHasilName,
    d.fileHasilUrl, d.fileHasilDownloadUrl, d.tanggalUpload, "Admin"
  ]);
  return id;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
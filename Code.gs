// =========================================================================
// KONFIGURASI GLOBAL AKSES DATABASE GOOGLE SPREADSHEET & DRIVE PTW LRTJ
// =========================================================================
const SPREADSHEET_ID = "1dsgwTqbrGW6d1q5ikguEwQToX5nLP5dozoyFcFDoNXo";  
const FOLDER_KARYAWAN = "1iSMjrpd3idvFBRGuyYMrD81xXevlCyfc";  
const FOLDER_SATPAM = "1QpjMbz5BNjVxc1xospUsGB-Fg5MhxxCJ";  
const FOLDER_PENGAWAS = "1HbCZW9jBLccSUdrTgzP06ZH7z6Y8HMWl";  
  
// Handle GET requests (CORS enabled for GitHub Pages deployment)
function doGet(e) {  
  // Deteksi jika ini adalah panggilan API eksternal (misal dari GitHub Pages)
  if (e && e.parameter && e.parameter.action === "getDatabaseData") {
    try {
      const data = getDatabaseData();
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: data }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  // Serve the index.html template internally inside Google environment
  return HtmlService.createTemplateFromFile('index')  
      .evaluate()  
      .setTitle('E-PTW: Permit To Work System - LRT Jakarta')  
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')  
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);  
}  

// Handle POST requests from GitHub Pages or other external platforms
function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    let resultData;

    if (action === "getDatabaseData") {
      resultData = getDatabaseData();
    } else if (action === "doLogin") {
      resultData = doLogin(requestData.email, requestData.password, requestData.role);
    } else if (action === "doRegisterVendor") {
      resultData = doRegisterVendor(requestData.namaUser, requestData.namaPerusahaan, requestData.email, requestData.password);
    } else if (action === "submitPTW") {
      resultData = submitPTW(requestData.formData, requestData.imageBase64, requestData.imageName);
    } else if (action === "approvePTW") {
      resultData = approvePTW(requestData.idPtw, requestData.approverName, requestData.isApproved, requestData.catatan);
    } else if (action === "verifyPTWArrival") {
      resultData = verifyPTWArrival(requestData.idPtw, requestData.verifierName);
    } else if (action === "completePTW") {
      resultData = completePTW(requestData.ptwId, requestData.station, requestData.operatorName, requestData.isLastStation);
    } else {
      throw new Error("Aksi backend tidak dikenal.");
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true, data: resultData }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
  
function getDatabaseData() {  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);  
    initDatabaseIfNeeded(ss);
    
    return {  
      satpam: getSheetValues(ss.getSheetByName("Data Satpam")),  
      karyawan: getSheetValues(ss.getSheetByName("Data Karyawan")),  
      vendorData: getSheetValues(ss.getSheetByName("Data Vendor")),
      ptw: getSheetValues(ss.getSheetByName("Histori Data PTW")),  
      login: getSheetValues(ss.getSheetByName("Histori Login")),  
      vendor: getSheetValues(ss.getSheetByName("Histori Pekerjaan Vendor"))  
    };  
  } catch (err) {
    throw new Error("Gagal memuatkan database Spreadsheet: " + err.message);
  }
}  
  
function getSheetValues(sheet) {  
  if (!sheet) return [];  
  const data = sheet.getDataRange().getValues();  
  if (data.length <= 1) return [];  
    
  const headers = data[0];  
  const rows = [];  
    
  for (let i = 1; i < data.length; i++) {  
    let rowObj = {};  
    for (let j = 0; j < headers.length; j++) {  
      let val = data[i][j];  
      if (val instanceof Date) {  
        if (val.getFullYear() === 1899) {
          const now = new Date();
          val.setFullYear(now.getFullYear());
          val.setMonth(now.getMonth());
          val.setDate(now.getDate());
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
        } else {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");  
        }
      }  
      rowObj[headers[j]] = val;  
    }  
    rows.push(rowObj);  
  }  
  return rows;  
}  

function doLogin(email, password, role) {  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);  
  initDatabaseIfNeeded(ss);

  let sheetName = "";
  if (role === "Satpam") {
    sheetName = "Data Satpam";
  } else if (role === "Karyawan") {
    sheetName = "Data Karyawan";
  } else if (role === "Vendor") {
    sheetName = "Data Vendor";
  } else {
    throw new Error("Peran pengguna tidak valid.");
  }

  const sheet = ss.getSheetByName(sheetName);  
  if (!sheet) throw new Error("Helaian " + sheetName + " tidak dijumpai.");

  const data = getSheetValues(sheet);  
    
  const cleanEmail = String(email || "").toLowerCase().trim();  
  const cleanPassword = String(password || "").trim();  
  
  const user = data.find(u => {
    const dbEmail = String(u.Email || "").toLowerCase().trim();
    const dbPassword = String(u.Password || "").trim();
    return dbEmail === cleanEmail && dbPassword === cleanPassword;
  });
  
  if (!user) {  
    throw new Error("Kredensial salah atau pengguna tidak terdaftar.");  
  }  

  if (role === "Vendor") {
    user.Nama = user["Nama User"] || "Pelaksana Vendor";
    user.Perusahaan = user["Nama Perusahaan"] || "Perusahaan Konstruksi";
    user["Lokasi Tugas"] = "Luar Lintasan / Lapangan";
  }
    
  const loginSheet = ss.getSheetByName("Histori Login");  
  if (loginSheet) {  
    const timestamp = new Date();
    loginSheet.appendRow([  
      timestamp,  
      user.Email,  
      user.Nama,  
      role,  
      user["Lokasi Tugas"] || "HQ Office LRTJ",  
      user["Lat GPS"] || "",  
      user["Lng Kordinat"] || user["Lng GPS"] || ""  
    ]);  
  }  
    
  return user;  
}  

function doRegisterVendor(namaUser, namaPerusahaan, email, password) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  initDatabaseIfNeeded(ss);

  const sheet = ss.getSheetByName("Data Vendor");
  if (!sheet) throw new Error("Helaian 'Data Vendor' tidak ditemukan.");

  const data = getSheetValues(sheet);
  const cleanEmail = String(email || "").toLowerCase().trim();

  const exists = data.some(u => String(u.Email || "").toLowerCase().trim() === cleanEmail);
  if (exists) {
    throw new Error("Email ini telah terdaftar sebagai akun Vendor.");
  }

  sheet.appendRow([
    cleanEmail,
    String(password).trim(),
    String(namaUser).trim(),
    String(namaPerusahaan).trim()
  ]);

  return { success: true, message: "Pendaftaran vendor berhasil." };
}
  
function submitPTW(formData, imageBase64, imageName) {  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);  
  initDatabaseIfNeeded(ss);
  
  const ptwSheet = ss.getSheetByName("Histori Data PTW");  
  if (!ptwSheet) throw new Error("Helaian Histori Data PTW tidak dijumpai.");

  let imageUrl = "";  
  if (imageBase64 && imageBase64.indexOf("base64,") !== -1) {  
    imageUrl = uploadFileToDrive(imageBase64, imageName, FOLDER_PENGAWAS);  
  }  
    
  const idPtw = "PTW-" + Math.floor(100000 + Math.random() * 900000);  
  const tanggalPengajuan = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");  
    
  const pengawasNama = formData.penanggungJawab || formData.namaPengawas || "Pengawas Lapangan";

  const rowData = [  
    idPtw,  
    tanggalPengajuan,  
    formData.tanggalExpired,  
    formData.jamMulai,  
    formData.jamSelesai,  
    formData.namaPerusahaan,  
    pengawasNama,  
    formData.jumlahPekerja,  
    formData.kelengkapanSafety,  
    formData.jenisPekerjaan,  
    formData.lokasiBekerja,  
    formData.latGps || "-6.1528",  
    formData.lngGps || "106.9114",  
    imageUrl || "https://images.unsplash.com/photo-1560250097-0b93528c311a?auto=format&fit=crop&w=150",  
    "Pending"
  ];  
    
  ptwSheet.appendRow(rowData);  
    
  const vendorSheet = ss.getSheetByName("Histori Pekerjaan Vendor");  
  if (vendorSheet) {  
    vendorSheet.appendRow([  
      idPtw,  
      formData.namaPerusahaan,  
      formData.jenisPekerjaan,  
      formData.lokasiBekerja,  
      "Diajukan oleh Vendor Mandiri",  
      "-",  
      "Menunggu Tinjauan Dokumen oleh Karyawan HSE QSHE",  
      new Date()  
    ]);  
  }  
    
  return { success: true, idPtw: idPtw };  
}  
  
function approvePTW(idPtw, approverName, isApproved, catatan) {  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);  
  initDatabaseIfNeeded(ss);
    
  const ptwSheet = ss.getSheetByName("Histori Data PTW");  
  if (!ptwSheet) throw new Error("Helaian Histori Data PTW tidak dijumpai.");

  const ptwData = ptwSheet.getDataRange().getValues();  
  let foundRow = -1;  
  let ptwDetails = {};  
    
  for (let i = 1; i < ptwData.length; i++) {  
    if (ptwData[i][0] === idPtw) {  
      foundRow = i + 1;  
      ptwDetails = {  
        perusahaan: ptwData[i][5],  
        jenis: ptwData[i][9],  
        lokasi: ptwData[i][10]  
      };  
      break;  
    }  
  }  
    
  if (foundRow === -1) throw new Error("Nomor dokumen PTW ID " + idPtw + " tidak dijumpai.");  
    
  const newStatus = isApproved ? "Awaiting Satpam Verification" : "Expired";  
  ptwSheet.getRange(foundRow, 15).setValue(newStatus); 
  
  const vendorSheet = ss.getSheetByName("Histori Pekerjaan Vendor");  
  if (vendorSheet) {  
    const statusTerakhirText = isApproved  
      ? "Disahkan Karyawan HSE (Menunggu Verifikasi Kehadiran Fisik oleh Satpam Sektor)"  
      : "Ditolak Karyawan HSE. Alasan: " + (catatan || "Persyaratan K3 tidak lengkap");

    vendorSheet.appendRow([  
      idPtw,  
      ptwDetails.perusahaan || "-",  
      ptwDetails.jenis || "-",  
      ptwDetails.lokasi || "-",  
      "-",  
      approverName + " (Karyawan)",  
      statusTerakhirText,  
      new Date()  
    ]);  
  }  
    
  return { success: true };  
}

function verifyPTWArrival(idPtw, verifierName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);  
  initDatabaseIfNeeded(ss);
    
  const ptwSheet = ss.getSheetByName("Histori Data PTW");  
  if (!ptwSheet) throw new Error("Helaian Histori Data PTW tidak dijumpai.");

  const ptwData = ptwSheet.getDataRange().getValues();  
  let foundRow = -1;  
  let ptwDetails = {};  
    
  for (let i = 1; i < ptwData.length; i++) {  
    if (ptwData[i][0] === idPtw) {  
      foundRow = i + 1;  
      ptwDetails = {  
        perusahaan: ptwData[i][5],  
        jenis: ptwData[i][9],  
        lokasi: ptwData[i][10]  
      };  
      break;  
    }  
  }  
    
  if (foundRow === -1) throw new Error("Nomor dokumen PTW ID " + idPtw + " tidak dijumpai.");  
    
  ptwSheet.getRange(foundRow, 15).setValue("Work In Progress");  
  
  const vendorSheet = ss.getSheetByName("Histori Pekerjaan Vendor");  
  if (vendorSheet) {
    vendorSheet.appendRow([  
      idPtw,  
      ptwDetails.perusahaan || "-",  
      ptwDetails.jenis || "-",  
      ptwDetails.lokasi || "-",  
      verifierName + " (Satpam)",  
      "-",  
      "Pekerjaan Dimulai - Fisik Vendor Diverifikasi di Lapangan",  
      new Date()  
    ]);
  }  
    
  return { success: true };  
}

function completePTW(ptwId, station, operatorName, isLastStation) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);  
    initDatabaseIfNeeded(ss);
      
    const ptwSheet = ss.getSheetByName("Histori Data PTW");  
    if (!ptwSheet) throw new Error("Helaian Histori Data PTW tidak dijumpai.");

    const ptwData = ptwSheet.getDataRange().getValues();  
    let foundRow = -1;  
    let ptwDetails = {};  
      
    for (let i = 1; i < ptwData.length; i++) {  
      if (ptwData[i][0] === ptwId) {  
        foundRow = i + 1;  
        ptwDetails = {  
          perusahaan: ptwData[i][5],  
          jenis: ptwData[i][9],  
          lokasi: ptwData[i][10]  
        };  
        break;  
      }  
    }  
      
    if (foundRow === -1) throw new Error("Nomor dokumen PTW ID " + ptwId + " tidak dijumpai.");  
      
    if (isLastStation) {
      ptwSheet.getRange(foundRow, 15).setValue("Expired");  
    }
    
    const vendorSheet = ss.getSheetByName("Histori Pekerjaan Vendor");  
    if (vendorSheet) {
      vendorSheet.appendRow([  
        ptwId,  
        ptwDetails.perusahaan || "-",  
        ptwDetails.jenis || "-",  
        ptwDetails.lokasi || "-",  
        operatorName + " (Satpam)",  
        "-",  
        "Pekerjaan Diakhiri di Sektor: " + station + (isLastStation ? " (Seluruh Sektor Selesai)" : ""),  
        new Date()  
      ]);  
    }  
      
    return { success: true };  
  } catch (err) {
    throw new Error("Gagal memproses akhir kerja di Spreadsheet: " + err.message);
  }
}
  
function uploadFileToDrive(base64Data, fileName, folderId) {  
  try {  
    const folder = DriveApp.getFolderById(folderId);  
    const contentType = base64Data.substring(5, base64Data.indexOf(';base64'));  
    const base64Clean = base64Data.substring(base64Data.indexOf(';base64,') + 8);  
    const data = Utilities.base64Decode(base64Clean);  
    const blob = Utilities.newBlob(data, contentType, fileName);  
    const file = folder.createFile(blob);  
    
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);  
    return file.getUrl();  
  } catch (e) {  
    return "Error Upload: " + e.toString();  
  }  
}

function initDatabaseIfNeeded(ss) {
  let sheetSatpam = ss.getSheetByName("Data Satpam");
  if (!sheetSatpam) {
    sheetSatpam = ss.insertSheet("Data Satpam");
  }
  if (sheetSatpam.getLastRow() <= 1) {
    sheetSatpam.clear();
    sheetSatpam.appendRow(["Email", "Password", "Nama", "NIK KTP", "Jabatan", "Lokasi Tugas", "Foto Satpam", "Lat GPS", "Lng Kordinat"]);
    sheetSatpam.appendRow(["satpam.pos1@company.com", "satpam123", "Bambang Hariyanto", "3216012304890001", "Komandan Regu A", "Depo MCC", "https://images.unsplash.com/photo-1600486913747-55e5470d6f40?auto=format&fit=crop&w=150&h=150&q=80", "-6.15500254", "106.91541503"]);
    sheetSatpam.appendRow(["satpam.pos2@company.com", "satpam123", "Slamet Prasetyo", "3216011209910003", "Anggota Regu B", "Stasiun Pegangsaan Dua", "https://images.unsplash.com/photo-1500648767791-00dcc994a43e?auto=format&fit=crop&w=150&h=150&q=80", "-6.15694188", "106.91420061"]);
  }

  let sheetKaryawan = ss.getSheetByName("Data Karyawan");
  if (!sheetKaryawan) {
    sheetKaryawan = ss.insertSheet("Data Karyawan");
  }
  if (sheetKaryawan.getLastRow() <= 1) {
    sheetKaryawan.clear();
    sheetKaryawan.appendRow(["Email", "Password", "Nama", "NIK KTP", "Jabatan", "Lokasi Tugas", "Foto Karyawan", "Lat GPS", "Lng Kordinat"]);
    sheetKaryawan.appendRow(["hse.karyawan@company.com", "karyawan123", "Dewi Lestari, S.T.", "3216010505850005", "HSE Specialist Coordinator", "Depo MCC - Ruang HSE", "https://images.unsplash.com/photo-1573496359142-b8d87734a5a2?auto=format&fit=crop&w=150&h=150&q=80", "-6.15500254", "106.91541503"]);
  }

  // Database Vendor Sheet Setup
  let sheetVendor = ss.getSheetByName("Data Vendor");
  if (!sheetVendor) {
    sheetVendor = ss.insertSheet("Data Vendor");
  }
  if (sheetVendor.getLastRow() === 0) {
    sheetVendor.appendRow(["Email", "Password", "Nama User", "Nama Perusahaan"]);
    sheetVendor.appendRow(["ahmad.vendor@company.com", "vendor123", "Ahmad Subarjo", "PT Sinar Jaya Abadi"]);
  }

  let sheetPTW = ss.getSheetByName("Histori Data PTW");
  if (!sheetPTW) {
    sheetPTW = ss.insertSheet("Histori Data PTW");
  }
  if (sheetPTW.getLastRow() === 0) {
    sheetPTW.appendRow(["ID PTW", "Tanggal Pengajuan", "Tanggal Expired", "Jam Mulai Pekerjaan", "Jam Selesai Pekerjaan", "Nama Perusahaan", "Nama Pengawas", "Jumlah Pekerja", "Kelengkapan Safety", "Jenis Pekerjaan", "Lokasi Bekerja", "Lat Gps", "Lng Gps", "Foto Pengawas", "Status"]);
    sheetPTW.appendRow(["PTW-20260711-001", "2026-07-11", "2026-07-11", "18:00", "23:00", "PT. Global Konstruksi Utama", "Budi Santoso", 8, "Helm, Sepatu Safety, Rompi", "Perawatan Jalur Rel & Wesel", "Stasiun Pegangsaan Dua", "-6.1528", "106.9114", "", "Work In Progress"]);
  }

  let sheetLogin = ss.getSheetByName("Histori Login");
  if (!sheetLogin) {
    sheetLogin = ss.insertSheet("Histori Login");
  }
  if (sheetLogin.getLastRow() === 0) {
    sheetLogin.appendRow(["Timestamp", "Email", "Nama", "Role", "Lokasi Tugas", "Lat GPS", "Lng GPS"]);
  }

  let sheetVendorHist = ss.getSheetByName("Histori Pekerjaan Vendor");
  if (!sheetVendorHist) {
    sheetVendorHist = ss.insertSheet("Histori Pekerjaan Vendor");
  }
}

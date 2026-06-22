/**
 * CV JAMPE AGENG TRADE - HR SYSTEM PROFESSIONAL (ULTIMATE STABLE VERSION)
 * Role: Senior Google Apps Script Developer & Database Architect
 */

const FOLDER_NAME = "HRIS_PRO_UPLOADS";

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('CV JAMPE AGENG TRADE (HR SYSTEM)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

const SCHEMA = {
  "Users": ["ID", "Email", "Password", "Role", "EmployeeID", "SessionToken"],
  "Employees": ["ID", "NIK", "NamaLengkap", "JenisKelamin", "TanggalLahir", "Alamat", "NoHP", "Email", "Divisi", "Jabatan", "Status", "TanggalMasuk", "GajiPokok", "Tunjangan", "BPJS", "NPWP", "Rekening", "Bank", "Foto"],
  "Shifts": ["ID", "NamaShift", "JamMasuk", "JamPulang", "Toleransi"],
  "Locations": ["ID", "NamaLokasi", "Latitude", "Longitude", "Radius"],
  "Attendance": ["ID", "Tanggal", "JamMasuk", "Nama", "NIK", "LatIn", "LngIn", "JarakIn", "SelfieIn", "StatusMasuk", "JamPulang", "LatOut", "LngOut", "SelfieOut", "StatusPulang", "TerlambatMenit"],
  "Overtime": ["ID", "Tanggal", "NIK", "Nama", "JamPulangAktual", "JamPulangNormal", "JamLembur", "TarifPerJam", "TotalUpah"],
  "Leave": ["ID", "Tanggal", "NIK", "Nama", "Jenis", "Keterangan", "Lampiran", "StatusManager", "StatusHRD"],
  "AttendanceRequests": ["ID", "Tanggal", "NIK", "Nama", "Jenis", "Jam", "Alasan", "StatusManager", "StatusHRD"],
  "ShiftRequests": ["ID", "TanggalMulai", "TanggalSelesai", "NIK", "Nama", "ShiftAwal", "ShiftTujuan", "Alasan", "StatusManager", "StatusHRD"],
  "LeaveBalances": ["ID", "NIK", "Tahun", "KuotaCuti", "SisaCuti"],
  "Payroll": ["ID", "BulanTahun", "NIK", "Nama", "GajiPokok", "Tunjangan", "UangMakan", "UangTransport", "Lembur", "Bonus", "Insentif", "BPJS_Kes", "BPJS_TK", "Pajak_PPh21", "Alpha", "Terlambat", "Pinjaman", "GajiBersih"],
  "Rates": ["ID", "Jabatan", "TarifLembur"],
  "AuditLog": ["ID", "User", "Aktivitas", "Tanggal", "Jam", "IP", "Device"],
  "Settings": ["Key", "Value"]
};

function setupDatabase() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    for (let sheetName in SCHEMA) {
      let sheet = ss.getSheetByName(sheetName);
      
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.appendRow(SCHEMA[sheetName]);
        sheet.getRange(1, 1, 1, SCHEMA[sheetName].length).setFontWeight("bold").setBackground("#e2e8f0");
        sheet.setFrozenRows(1);
      } 
      else {
        let lastCol = sheet.getLastColumn();
        if (lastCol === 0) {
           sheet.appendRow(SCHEMA[sheetName]);
           sheet.getRange(1, 1, 1, SCHEMA[sheetName].length).setFontWeight("bold").setBackground("#e2e8f0");
        } else {
           let existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
           let newHeaders = [...existingHeaders];
           let isMissingColumn = false;
           
           SCHEMA[sheetName].forEach(col => {
             if (!existingHeaders.includes(col)) {
               newHeaders.push(col); 
               isMissingColumn = true;
             }
           });
           
           if (isMissingColumn) {
             sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
             sheet.getRange(1, 1, 1, newHeaders.length).setFontWeight("bold").setBackground("#e2e8f0");
           }
        }
      }
    }
    
    let folders = DriveApp.getFoldersByName(FOLDER_NAME);
    if (!folders.hasNext()) DriveApp.createFolder(FOLDER_NAME);

    let usersSheet = ss.getSheetByName("Users");
    if(usersSheet && usersSheet.getDataRange().getValues().length <= 1) {
      usersSheet.appendRow([generateUUID(), "admin@jampeageng.com", "admin123", "Super Admin", "", ""]);
    }
    return {success: true, message: "Database telah dipindai dan diperbarui otomatis (Auto-Heal)!"};
  } catch(e) {
    throw new Error("Gagal Setup Database: " + e.message);
  }
}

function getDb_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if(!sheet) throw new Error(`Tabel ${sheetName} tidak ditemukan. Silakan jalankan setupDatabase().`);
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i].join("").trim() === "") continue; 
    let obj = {};
    for (let j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
    obj._rowIndex = i + 1; 
    rows.push(obj);
  }
  return { sheet, headers, rows };
}

function generateUUID() {
  return 'xxxx-4xxx-yxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

function getRecords(sheetName) { 
  const rows = getDb_(sheetName).rows; 
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  
  return rows.map(r => {
    let cleanRow = {};
    for(let k in r) {
      if (r[k] instanceof Date) {
        if (r[k].getFullYear() === 1899 || r[k].getFullYear() === 1970) {
          cleanRow[k] = Utilities.formatDate(r[k], tz, "HH:mm");
        } else {
          cleanRow[k] = Utilities.formatDate(r[k], tz, "yyyy-MM-dd");
        }
      } else if (r[k] === null || r[k] === undefined) {
        cleanRow[k] = "";
      } else {
        cleanRow[k] = r[k];
      }
    }
    return cleanRow;
  });
}

function insertRecord(sheetName, payload) {
  const db = getDb_(sheetName);
  payload.ID = generateUUID();
  const rowArray = db.headers.map(h => payload[h] !== undefined ? payload[h] : "");
  db.sheet.appendRow(rowArray);
  SpreadsheetApp.flush();
  return { success: true, message: "Data berhasil disimpan.", data: payload };
}

function updateRecord(sheetName, id, payload) {
  const db = getDb_(sheetName);
  const target = db.rows.find(r => r.ID == id);
  if(!target) throw new Error("Data tidak ditemukan.");
  const rowArray = db.headers.map(h => payload[h] !== undefined ? payload[h] : target[h]);
  db.sheet.getRange(target._rowIndex, 1, 1, rowArray.length).setValues([rowArray]);
  SpreadsheetApp.flush();
  return { success: true, message: "Data berhasil diupdate." };
}

function deleteRecord(sheetName, id) {
  const db = getDb_(sheetName);
  const target = db.rows.find(r => r.ID == id);
  if(target) {
    db.sheet.deleteRow(target._rowIndex);
    SpreadsheetApp.flush();
    return { success: true, message: "Data berhasil dihapus." };
  }
  throw new Error("Data tidak ditemukan.");
}

function login(email, password) {
  const users = getRecords("Users");
  const user = users.find(u => u.Email === email && u.Password === password);
  if(user) {
    const token = generateUUID();
    let updates = { SessionToken: token };
    let role = user.Role;
    if(role.toLowerCase() === 'user' || role.toLowerCase() === 'karyawan') { role = 'Karyawan'; updates.Role = role; }
    
    let employees = getRecords("Employees");
    let employeeData = null;
    
    if(user.EmployeeID) employeeData = employees.find(e => e.NIK === user.EmployeeID || e.ID === user.EmployeeID);
    if(!employeeData) {
        employeeData = employees.find(e => e.Email && String(e.Email).toLowerCase() === String(email).toLowerCase());
        if(employeeData) updates.EmployeeID = employeeData.NIK;
    }
    
    updateRecord("Users", user.ID, updates);
    auditLog(user.Email, "Login System");
    
    return { success: true, token: token, role: role, employee: employeeData, name: employeeData ? employeeData.NamaLengkap : role };
  }
  return { success: false, message: "Akses Ditolak: Email atau Password salah." };
}

function auditLog(user, activity) {
  try {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    insertRecord("AuditLog", { User: user, Aktivitas: activity, Tanggal: Utilities.formatDate(new Date(), tz, "yyyy-MM-dd"), Jam: Utilities.formatDate(new Date(), tz, "HH:mm:ss"), IP: "System Client", Device: "Browser" });
  } catch(e) {}
}

function uploadFileToDrive(base64Data, filename) {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
  const contentType = base64Data.substring(5, base64Data.indexOf(';'));
  const bytes = Utilities.base64Decode(base64Data.substr(base64Data.indexOf('base64,') + 7));
  const file = folder.createFile(Utilities.newBlob(bytes, contentType, filename));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function calculateDistance(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const dLat = (lat2-lat1) * (Math.PI/180);
  const dLon = (lon2-lon1) * (Math.PI/180);
  const a = Math.sin(dLat/2) * Math.sin(dLat/2) + Math.cos(lat1*(Math.PI/180)) * Math.cos(lat2*(Math.PI/180)) * Math.sin(dLon/2) * Math.sin(dLon/2);
  return R * (2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)));
}

function clockIn(payload) {
  const locations = getRecords("Locations");
  if(locations.length === 0) throw new Error("Master lokasi kantor belum disetting oleh HRD.");
  
  let allowed = false, distance = 999999;
  for(let i=0; i<locations.length; i++) {
    let d = calculateDistance(payload.LatIn, payload.LngIn, locations[i].Latitude, locations[i].Longitude);
    if(d <= (parseFloat(locations[i].Radius) || 100)) { allowed = true; distance = Math.round(d); break; }
  }
  if(!allowed) throw new Error(`Absensi ditolak Geofencing. Anda berada di luar radius kantor (${Math.round(calculateDistance(payload.LatIn, payload.LngIn, locations[0].Latitude, locations[0].Longitude))}m).`);
  
  const shifts = getRecords("Shifts");
  let status = "Hadir", lateMins = 0;
  if(shifts.length > 0 && shifts[0].JamMasuk) {
    let s = shifts[0]; 
    let shiftInMin = parseInt(s.JamMasuk.split(":")[0])*60 + parseInt(s.JamMasuk.split(":")[1]);
    let actualInMin = parseInt(payload.JamMasuk.split(":")[0])*60 + parseInt(payload.JamMasuk.split(":")[1]);
    if(actualInMin > (shiftInMin + parseInt(s.Toleransi||0))) {
      status = "Terlambat"; lateMins = actualInMin - shiftInMin;
    }
  }

  let photoUrl = uploadFileToDrive(payload.SelfieData, `IN_${payload.NIK}_${payload.Tanggal}.jpg`);
  insertRecord("Attendance", { Tanggal: payload.Tanggal, JamMasuk: payload.JamMasuk, Nama: payload.Nama, NIK: payload.NIK, LatIn: payload.LatIn, LngIn: payload.LngIn, JarakIn: distance, SelfieIn: photoUrl, StatusMasuk: status, TerlambatMenit: lateMins });
  return { success: true, message: `Clock In sukses. Status Kehadiran: ${status}` };
}

function clockOut(payload) {
  const db = getDb_("Attendance");
  const todayRecord = db.rows.find(r => r.Tanggal === payload.Tanggal && r.NIK === payload.NIK);
  if(!todayRecord) throw new Error("Anda belum melakukan Clock In hari ini.");
  if(todayRecord.JamPulang) throw new Error("Anda sudah melakukan Clock Out hari ini.");

  let photoUrl = uploadFileToDrive(payload.SelfieData, `OUT_${payload.NIK}_${payload.Tanggal}.jpg`);
  updateRecord("Attendance", todayRecord.ID, { JamPulang: payload.JamPulang, LatOut: payload.LatOut, LngOut: payload.LngOut, SelfieOut: photoUrl, StatusPulang: "Selesai" });
  
  const shifts = getRecords("Shifts");
  if(shifts.length > 0 && shifts[0].JamPulang) {
    let s = shifts[0];
    let shiftOutMin = parseInt(s.JamPulang.split(":")[0])*60 + parseInt(s.JamPulang.split(":")[1]);
    let actualOutMin = parseInt(payload.JamPulang.split(":")[0])*60 + parseInt(payload.JamPulang.split(":")[1]);
    
    if(actualOutMin >= (shiftOutMin + 60)) {
       let overTimeHours = parseFloat(((actualOutMin - (shiftOutMin + 60)) / 60).toFixed(2));
       if(overTimeHours > 0) {
          let emp = getRecords("Employees").find(e => e.NIK === payload.NIK);
          let rateData = getRecords("Rates").find(r => r.Jabatan === emp.Jabatan);
          let rate = rateData ? parseFloat(rateData.TarifLembur) : 25000;
          insertRecord("Overtime", { Tanggal: payload.Tanggal, NIK: payload.NIK, Nama: emp.NamaLengkap, JamPulangAktual: payload.JamPulang, JamPulangNormal: s.JamPulang, JamLembur: overTimeHours, TarifPerJam: rate, TotalUpah: overTimeHours * rate });
       }
    }
  }
  return { success: true, message: "Clock Out sukses terpusat ke server." };
}

function getMyAttendanceLog(nik, month) {
  const atts = getRecords("Attendance").filter(a => String(a.NIK) === String(nik) && String(a.Tanggal).startsWith(month));
  const reqs = getRecords("AttendanceRequests").filter(a => String(a.NIK) === String(nik) && String(a.Tanggal).startsWith(month));
  
  let [y, m] = month.split('-');
  let daysInMonth = new Date(y, m, 0).getDate();
  let logData = [];
  
  for(let i=1; i<=daysInMonth; i++) {
    let dateStr = `${y}-${m}-${String(i).padStart(2,'0')}`;
    if(dateStr > Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd")) break; 
    
    let att = atts.find(a => a.Tanggal === dateStr);
    let reqIn = reqs.find(r => r.Tanggal === dateStr && r.Jenis === 'Masuk');
    let reqOut = reqs.find(r => r.Tanggal === dateStr && r.Jenis === 'Pulang');
    
    logData.push({
      Tanggal: dateStr,
      JamMasuk: att ? (att.JamMasuk || '-') : '-',
      StatusMasuk: att ? (att.StatusMasuk || 'Hadir') : 'Belum Absen',
      JamPulang: att ? (att.JamPulang || '-') : '-',
      StatusPulang: att ? (att.StatusPulang || 'Selesai') : 'Belum Clock Out',
      ReqMasuk: reqIn ? `Koreksi: ${reqIn.Jam} (${reqIn.StatusHRD})` : null,
      ReqPulang: reqOut ? `Koreksi: ${reqOut.Jam} (${reqOut.StatusHRD})` : null
    });
  }
  return logData.reverse(); 
}

function submitAttendanceRequest(payload) {
  insertRecord("AttendanceRequests", { Tanggal: payload.Tanggal, NIK: payload.NIK, Nama: payload.Nama, Jenis: payload.Jenis, Jam: payload.Jam, Alasan: payload.Alasan, StatusManager: "Pending", StatusHRD: "Pending" });
  return { success: true, message: "Pengajuan koreksi absen berhasil dikirim ke HRD." };
}

function processAttendanceRequest(id, role, status) {
  let req = getRecords("AttendanceRequests").find(r => r.ID === id);
  if(!req) throw new Error("Data pengajuan tidak ditemukan.");
  
  let upd = {};
  if(role === "Manager") upd.StatusManager = status;
  if(role === "HRD" || role === "Super Admin") {
      upd.StatusHRD = status;
      if(status === 'Approved') {
          let attDb = getDb_("Attendance");
          let att = attDb.rows.find(a => a.Tanggal === req.Tanggal && a.NIK === req.NIK);
          
          if(att) {
              if(req.Jenis === 'Masuk') updateRecord("Attendance", att.ID, { JamMasuk: req.Jam, StatusMasuk: 'Hadir (Koreksi)' });
              else updateRecord("Attendance", att.ID, { JamPulang: req.Jam, StatusPulang: 'Selesai (Koreksi)' });
          } else {
              if(req.Jenis === 'Masuk') insertRecord("Attendance", { Tanggal: req.Tanggal, NIK: req.NIK, Nama: req.Nama, JamMasuk: req.Jam, StatusMasuk: 'Hadir (Koreksi)', JamPulang: '', StatusPulang: ''});
          }
      }
  }
  updateRecord("AttendanceRequests", id, upd);
  return { success: true, message: `Koreksi Absensi di-${status}.` };
}

function submitShiftRequest(payload) {
  insertRecord("ShiftRequests", { TanggalMulai: payload.TanggalMulai, TanggalSelesai: payload.TanggalSelesai, NIK: payload.NIK, Nama: payload.Nama, ShiftAwal: payload.ShiftAwal, ShiftTujuan: payload.ShiftTujuan, Alasan: payload.Alasan, StatusManager: "Pending", StatusHRD: "Pending" });
  return { success: true, message: "Pengajuan ubah shift berhasil dikirim." };
}

function processShiftRequest(id, role, status) {
  let upd = {};
  if(role === "Manager") upd.StatusManager = status;
  if(role === "HRD" || role === "Super Admin") upd.StatusHRD = status;
  updateRecord("ShiftRequests", id, upd);
  return { success: true, message: `Permintaan Shift di-${status}.` };
}

function submitLeave(payload) {
  // PENGAMAN BACKEND CUTI (TIDAK BISA DITEMBUS)
  if (payload.Jenis === "Cuti") {
      let year = String(payload.Tanggal).substring(0,4);
      let bal = getRecords("LeaveBalances").find(b => String(b.NIK) === String(payload.NIK) && String(b.Tahun) === year);
      let sisa = bal ? parseInt(bal.SisaCuti) : 12;
      
      if (sisa <= 0) {
          throw new Error("Pengajuan Ditolak: Hak Cuti Tahunan Anda sudah habis (0 Hari).");
      }
  }

  let fileUrl = payload.LampiranBase64 ? uploadFileToDrive(payload.LampiranBase64, payload.LampiranName) : "";
  insertRecord("Leave", { Tanggal: payload.Tanggal, NIK: payload.NIK, Nama: payload.Nama, Jenis: payload.Jenis, Keterangan: payload.Keterangan, Lampiran: fileUrl, StatusManager: "Pending", StatusHRD: "Pending" });
  return { success: true, message: "Formulir Izin/Cuti berhasil diajukan." };
}

function processLeave(id, role, status) {
  let req = getRecords("Leave").find(r => r.ID === id);
  if(!req) throw new Error("Data pengajuan tidak ditemukan.");
  
  let upd = {};
  if(role === "Manager") upd.StatusManager = status;
  if(role === "HRD" || role === "Super Admin") {
      upd.StatusHRD = status;
      
      // LOGIKA PEMOTONGAN CUTI OTOMATIS
      if(status === 'Approved' && req.Jenis === 'Cuti') {
          let year = String(req.Tanggal).substring(0,4);
          let balDb = getDb_("LeaveBalances");
          let bal = balDb.rows.find(b => b.NIK === req.NIK && String(b.Tahun) === year);
          
          if(bal) {
              let sisa = parseInt(bal.SisaCuti) - 1;
              updateRecord("LeaveBalances", bal.ID, { SisaCuti: sisa });
          } else {
              // Jika ini cuti pertama di tahun ini, buat kuota (12) dan langsung potong 1 jadi (11)
              insertRecord("LeaveBalances", { NIK: req.NIK, Tahun: year, KuotaCuti: 12, SisaCuti: 11 });
          }
      }
  }
  updateRecord("Leave", id, upd);
  return { success: true, message: `Dokumen Izin/Cuti di-${status}.` };
}

function generatePayrollData(month) {
  try {
      const emps = getRecords("Employees");
      if(emps.length === 0) throw new Error("Belum ada data Karyawan.");
      
      const atts = getRecords("Attendance").filter(a => String(a.Tanggal).startsWith(month));
      const overs = getRecords("Overtime").filter(o => String(o.Tanggal).startsWith(month));
      const leaves = getRecords("Leave").filter(l => String(l.Tanggal).startsWith(month) && l.StatusHRD === "Approved");
      const existing = getRecords("Payroll").filter(p => String(p.BulanTahun) === month);
      
      emps.forEach(emp => {
        let e_atts = atts.filter(a => String(a.NIK) === String(emp.NIK));
        let e_overs = overs.filter(o => String(o.NIK) === String(emp.NIK));
        let e_leaves = leaves.filter(l => String(l.NIK) === String(emp.NIK));
        
        let gp = parseFloat(emp.GajiPokok) || 0;
        let tunj = parseFloat(emp.Tunjangan) || 0;
        
        let uMakan = e_atts.length * 50000;
        let uTrans = e_atts.length * 25000; 
        
        let lembur = e_overs.reduce((acc, o) => acc + (parseFloat(o.TotalUpah) || 0), 0);
        
        let bpjs_kes = gp > 0 ? (gp * 0.01) : 0; 
        let bpjs_tk = gp > 0 ? (gp * 0.02) : 0;  
        
        let bruto = gp + tunj + uMakan + uTrans + lembur;
        let pph21 = 0;
        if(bruto * 12 > 54000000) { pph21 = ((bruto * 12 - 54000000) * 0.05) / 12; } 
        
        let telat = e_atts.filter(a => String(a.StatusMasuk).includes("Terlambat")).length * 25000; 
        
        let alphaHari = 22 - e_atts.length - e_leaves.length; 
        if(alphaHari < 0) alphaHari = 0;
        let alpha = alphaHari > 0 ? alphaHari * (gp/22) : 0;
        
        let bersih = bruto - bpjs_kes - bpjs_tk - pph21 - telat - alpha;
        if(bersih < 0) bersih = 0;
        
        let pyLoad = { 
            BulanTahun: month, NIK: emp.NIK, Nama: emp.NamaLengkap, 
            GajiPokok: gp, Tunjangan: tunj, UangMakan: uMakan, UangTransport: uTrans, Lembur: lembur, Bonus: 0, Insentif: 0, 
            BPJS_Kes: bpjs_kes, BPJS_TK: bpjs_tk, Pajak_PPh21: pph21, 
            Alpha: Math.round(alpha), Terlambat: telat, Pinjaman: 0, 
            GajiBersih: Math.round(bersih) 
        };
        
        let check = existing.find(p => String(p.NIK) === String(emp.NIK));
        if(check) updateRecord("Payroll", check.ID, pyLoad);
        else insertRecord("Payroll", pyLoad);
      });
      return { success: true, message: `Kalkulasi Payroll untuk ${emps.length} Karyawan periode ${month} selesai.` };
  } catch(e) { throw new Error(e.message); }
}

function getReportData(type, start, end) {
  let data = [];
  if(type.includes("Absensi") || type === "Keterlambatan") data = getRecords("Attendance");
  else if(type.includes("Izin") || type === "Cuti" || type === "Sakit") data = getRecords("Leave");
  else if(type === "Lembur") data = getRecords("Overtime");
  else if(type === "Payroll") data = getRecords("Payroll");
  
  if(start && end && type !== "Payroll") data = data.filter(r => r.Tanggal >= start && r.Tanggal <= end);
  
  if(type === "Keterlambatan") data = data.filter(r => String(r.StatusMasuk).includes("Terlambat"));
  if(type === "Izin") data = data.filter(r => r.Jenis === "Izin");
  if(type === "Cuti") data = data.filter(r => r.Jenis === "Cuti");
  if(type === "Sakit") data = data.filter(r => r.Jenis === "Sakit");
  return data;
}

function getDashboardData(role, nik) {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const month = today.substring(0,7);
  
  let emps = getRecords("Employees");
  let atts = getRecords("Attendance");
  let leaves = getRecords("Leave"); 
  let overtimes = getRecords("Overtime");
  let payrolls = getRecords("Payroll").filter(p => String(p.BulanTahun) === month);
  
  let todayAtts = atts.filter(a => a.Tanggal === today);
  
  if (role === "Karyawan") {
    let myAtts = atts.filter(a => String(a.NIK) === String(nik) && String(a.Tanggal).startsWith(month));
    
    // Tarik Data Sisa Cuti dari Database
    let year = today.substring(0,4);
    let bal = getRecords("LeaveBalances").find(b => String(b.NIK) === String(nik) && String(b.Tahun) === year);
    let sisaCuti = bal ? parseInt(bal.SisaCuti) : 12; // Default 12 jika belum pernah cuti tahun ini

    return {
      Hadir: myAtts.length, Terlambat: myAtts.filter(a => String(a.StatusMasuk || '').includes("Terlambat")).length,
      Izin: leaves.filter(l => String(l.NIK) === String(nik) && String(l.Tanggal).startsWith(month)).length,
      Lembur: overtimes.filter(o => String(o.NIK) === String(nik) && String(o.Tanggal).startsWith(month)).reduce((acc, o) => acc + parseFloat(o.JamLembur||0), 0),
      SisaCuti: sisaCuti // Kirim ke antarmuka HTML
    };
  } else {
    let attReqs = getRecords("AttendanceRequests");
    let shiftReqs = getRecords("ShiftRequests");
    
    let pLeave = leaves.filter(r => r.StatusHRD === "Pending" || r.StatusHRD === "").length;
    let pAbsen = attReqs.filter(r => r.StatusHRD === "Pending" || r.StatusHRD === "").length;
    let pShift = shiftReqs.filter(r => r.StatusHRD === "Pending" || r.StatusHRD === "").length;

    let chartLabels = [], chartHadir = [], chartTerlambat = [];
    for(let i=6; i>=0; i--) {
        let d = new Date(); d.setDate(d.getDate() - i);
        let dateStr = Utilities.formatDate(d, tz, "yyyy-MM-dd");
        let dayAtts = atts.filter(a => a.Tanggal === dateStr);
        chartLabels.push(dateStr.substring(8,10) + '/' + dateStr.substring(5,7));
        chartHadir.push(dayAtts.filter(a => String(a.StatusMasuk || '').includes("Hadir")).length);
        chartTerlambat.push(dayAtts.filter(a => String(a.StatusMasuk || '').includes("Terlambat")).length);
    }
    
    return {
      TotalKaryawan: emps.length, HadirHariIni: todayAtts.length, 
      TerlambatHariIni: todayAtts.filter(a => String(a.StatusMasuk || '').includes("Terlambat")).length, 
      SedangLembur: overtimes.filter(o => o.Tanggal === today).length,
      TotalGaji: payrolls.reduce((s,p) => s + parseFloat(p.GajiBersih||0), 0),
      TotalPotongan: payrolls.reduce((s,p) => s + (parseFloat(p.BPJS_Kes||0)+parseFloat(p.BPJS_TK||0)+parseFloat(p.Pajak_PPh21||0)+parseFloat(p.Terlambat||0)+parseFloat(p.Alpha||0)), 0),
      PendingLeave: pLeave, PendingAbsen: pAbsen, PendingShift: pShift,
      ChartLabels: chartLabels, ChartHadir: chartHadir, ChartTerlambat: chartTerlambat
    };
  }
}
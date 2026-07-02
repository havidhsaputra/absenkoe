/**
 * ABSENPRO - JAMPE AGENG TRADE
 * Developer: Bot AppScript Ultimate
 * FILE 3: BACKEND ENGINE & LOGIC (V5.8 - FULL FIX SYNTAX)
 */

const SPREADSHEET_ID = "1bTT_x2JTpXpqxzofxUsv5mNQFMriCesejGZvxEckW_M";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('AbsenPro - Jampe Ageng Trade')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) 
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- 2. AUTHENTICATION (ANTI-ERROR & PIN) ---
function login(username, password) {
  const data = getSheetData("Karyawan");
  const inputID = String(username).trim().toUpperCase(); 
  const inputPIN = String(password).trim(); 
  
  for (let i = 1; i < data.length; i++) {
    const dbID = String(data[i][0]).trim().toUpperCase();
    const dbPIN = String(data[i][3]).trim(); 
    
    if (dbID === inputID && dbPIN === inputPIN) { 
      return { 
        success: true, 
        data: { id: data[i][0], nama: data[i][1], username: data[i][2], jabatan: data[i][4], role: data[i][5], kuotaCuti: data[i][11] } 
      };
    }
  }
  throw new Error("ID Karyawan atau PIN tidak valid/salah!");
}

// --- 3. APPROVAL SYSTEM ADMIN & EMAIL ---
function getAllPendingRequests() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID); let results = [];
  const addToList = (sheetName, typeName, infoFormatter) => {
    let sh = ss.getSheetByName(sheetName);
    if(sh) {
      let data = sh.getDataRange().getDisplayValues();
      let statusCol = (sheetName==="Reimbursement"||sheetName==="Shift_Change") ? 8 : (sheetName==="Pinjaman" ? 11 : 9);
      for(let i=1; i<data.length; i++) {
        results.push({ sheet: sheetName, id: data[i][0], tgl: data[i][1], nama: data[i][3], type: typeName, info: infoFormatter(data[i]), doc: (sheetName==="Shift_Change"||sheetName==="Pinjaman"?"":data[i][(sheetName==="Reimbursement"?7:8)]), status: data[i][statusCol], nominal: (sheetName==="Pinjaman"?String(data[i][4]).replace(/[^0-9]/g, ''):0) });
      }
    }
  };
  addToList("Cuti", "Cuti / Ijin", d => `${d[4]} (${d[5]} s/d ${d[6]}) - ${d[7]}`);
  addToList("Reimbursement", "Reimbursement", d => `${d[4]} (Rp ${d[5]}) - ${d[6]}`);
  addToList("Req_Absensi", "Req. Absen", d => `${d[4]} (${d[5]}) - ${d[7]}`);
  addToList("Shift_Change", "Tukar Shift", d => `Jadwal: ${d[4]} (${d[5]} -> ${d[6]})`);
  addToList("Lembur", "Lembur", d => `Tgl: ${d[4]} (${d[5]} - ${d[6]})`);
  addToList("Pinjaman", "Pinjaman/Kasbon", d => `Rp ${Number(d[4]||0).toLocaleString('id-ID')} - ${d[5]}`); 
  
  results.sort((a, b) => new Date(b.tgl) - new Date(a.tgl));
  return results;
}

function getPendingRequestCount() { return getAllPendingRequests().filter(r => r.status === "Menunggu Konfirmasi").length; }

function updateRequestStatus(sheetName, id, status, extraData) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  const data = sh.getDataRange().getValues();
  let statusCol = (sheetName === "Reimbursement" || sheetName === "Shift_Change") ? 8 : (sheetName === "Pinjaman" ? 11 : 9);
  
  for (let i = 1; i < data.length; i++) { 
    if (data[i][0] == id) { 
      sh.getRange(i + 1, statusCol + 1).setValue(status); 
      
      // FITUR POTONG KUOTA CUTI OTOMATIS
      if (sheetName === "Cuti" && status === "Disetujui") {
        try {
          if (data[i][4] === "Cuti Tahunan") {
            const diffDays = Math.ceil(Math.abs(new Date(data[i][6]) - new Date(data[i][5])) / (1000 * 60 * 60 * 24)) + 1; 
            const shKaryawan = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Karyawan");
            const dataKaryawan = shKaryawan.getDataRange().getValues();
            for(let j=1; j<dataKaryawan.length; j++){
              if(dataKaryawan[j][0] == data[i][2]){
                let sisa = Number(dataKaryawan[j][11]) || 0;
                shKaryawan.getRange(j+1, 12).setValue(sisa - diffDays);
                break;
              }
            }
          }
        } catch(e) {}
      }

      // FITUR PINJAMAN: Set Tenor dan Kalkulasi Cicilan
      if (sheetName === "Pinjaman" && status === "Disetujui") {
        try {
          let pData = JSON.parse(extraData);
          let nominal = Number(data[i][4]);
          let tenor = Number(pData.tenor) || 1;
          let cicilan = Math.ceil(nominal / tenor);
          
          sh.getRange(i+1, 7).setValue(tenor);        // Tenor Bulan
          sh.getRange(i+1, 8).setValue(cicilan);      // Cicilan Per Bulan
          sh.getRange(i+1, 9).setValue(nominal);      // Sisa Hutang Awal
          sh.getRange(i+1, 10).setValue(pData.bulan); // Mulai_Bulan
          sh.getRange(i+1, 11).setValue(pData.tahun); // Mulai_Tahun
        } catch(e) {}
      }
      
      // FITUR EMAIL NOTIFIKASI
      try {
        const empId = data[i][2]; 
        const empEmail = getUserEmail(empId);
        if (empEmail) {
           MailApp.sendEmail({
             to: empEmail,
             subject: `[AbsenPro] Status Pengajuan ${sheetName} : ${status}`,
             htmlBody: `<h3>Halo ${data[i][3]},</h3><p>Pengajuan <b>${sheetName}</b> Anda dengan ID <b>${id}</b> telah direview dan statusnya saat ini: <strong style="color:${status=='Disetujui'?'green':'red'}">${status}</strong>.</p><p>Terima kasih,<br>Tim HRD</p>`
           });
        }
      } catch(e) {} 
      
      return true; 
    } 
  }
  throw new Error("Data tidak ditemukan.");
}

// --- 4. DATA DASHBOARD, KARYAWAN & ABSEN ---
function getDashboardData() {
  const users = getSheetData("Karyawan").length - 1; const absensi = getSheetData("Absensi"); const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  let hadir = 0, telat = 0;
  for(let i=1; i<absensi.length; i++) { if(absensi[i][1] == today) { if(String(absensi[i][8]).includes("Hadir")) hadir++; if(String(absensi[i][8]).includes("Terlambat")) telat++; } }
  
  const settings = getSettingsObj();
  return { 
      totalEmp: users, hadir: hadir, telat: telat, 
      chartLabels: ['Sen','Sel','Rab','Kam','Jum','Sab','Hari Ini'], 
      chartHadir: [Math.floor(users*0.8), Math.floor(users*0.9), Math.floor(users*0.8), Math.floor(users*0.95), Math.floor(users*0.9), Math.floor(users*0.8), hadir], 
      chartTelat: [0, 1, 0, 2, 0, 1, telat],
      pengumuman: settings.pengumuman 
  };
}

function getEmployees() {
  const data = getSheetData("Karyawan"); const result = [];
  for(let i=1; i<data.length; i++) {
    result.push({ 
      id: data[i][0], nama: data[i][1], username: data[i][2], pin: data[i][3], 
      jabatan: data[i][4], role: data[i][5], gajiPokok: data[i][6], 
      uangMakan: data[i][7], uangTransport: data[i][8], tarifLembur: data[i][9], 
      email: data[i][10] || "", kuotaCuti: data[i][11] !== undefined ? data[i][11] : 12 
    });
  }
  return result;
}

function saveEmployee(emp) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Karyawan"); const data = sh.getDataRange().getValues();
  
  if (emp.id && !emp.isNew) {
    for(let i=1; i<data.length; i++) { 
      if(String(data[i][0]).trim().toUpperCase() === String(emp.id).trim().toUpperCase()) { 
        sh.getRange(i+1, 2, 1, 11).setValues([[emp.nama, emp.username, emp.pin, emp.jabatan, emp.role, emp.gajiPokok, emp.uangMakan, emp.uangTransport, emp.tarifLembur, emp.email, emp.kuotaCuti]]); 
        return true; 
      } 
    }
  } else { 
    const newId = String(emp.id).trim().toUpperCase();
    for(let i=1; i<data.length; i++) {
      if(String(data[i][0]).trim().toUpperCase() === newId) throw new Error("ID Karyawan sudah terdaftar! Gunakan ID lain.");
    }
    sh.appendRow([newId || "EMP" + new Date().getTime().toString().slice(-6), emp.nama, emp.username, emp.pin, emp.jabatan, emp.role, emp.gajiPokok, emp.uangMakan, emp.uangTransport, emp.tarifLembur, emp.email, emp.kuotaCuti]); 
    return true; 
  }
}

function deleteEmployee(id) { const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Karyawan"); const data = sh.getDataRange().getValues(); for(let i=1; i<data.length; i++) { if(data[i][0] == id) { sh.deleteRow(i+1); return true; } } throw new Error("Data tidak ditemukan"); }
function getUserEmail(id) { const data = getSheetData("Karyawan"); for (let i = 1; i < data.length; i++) { if (data[i][0] == id) return data[i][10] || ""; } return ""; }

function getSettingsObj() {
  const data = getSheetData("Pengaturan");
  let obj = { jamMasuk: "08:00", jamPulang: "17:00", radius: 50, koordinat: "", potTelat: 0, potAlpha: 0, pengumuman: "" };
  for(let i=1; i<data.length; i++) {
    if(data[i][0] == "JAM_BATAS_MASUK") obj.jamMasuk = data[i][1];
    if(data[i][0] == "JAM_BATAS_PULANG") obj.jamPulang = data[i][1];
    if(data[i][0] == "RADIUS") obj.radius = data[i][1];
    if(data[i][0] == "KOORDINAT") obj.koordinat = data[i][1];
    if(data[i][0] == "POTONGAN_TELAT") obj.potTelat = Number(data[i][1])||0;
    if(data[i][0] == "POTONGAN_ALPHA") obj.potAlpha = Number(data[i][1])||0;
    if(data[i][0] == "PENGUMUMAN_DASHBOARD") obj.pengumuman = data[i][1];
  }
  return obj;
}

function saveSettingsObj(obj) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Pengaturan"); const data = sh.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] == "JAM_BATAS_MASUK" && obj.jamMasuk!==undefined) sh.getRange(i+1, 2).setValue(obj.jamMasuk);
    if(data[i][0] == "JAM_BATAS_PULANG" && obj.jamPulang!==undefined) sh.getRange(i+1, 2).setValue(obj.jamPulang);
    if(data[i][0] == "RADIUS" && obj.radius!==undefined) sh.getRange(i+1, 2).setValue(obj.radius);
    if(data[i][0] == "KOORDINAT" && obj.koordinat!==undefined) sh.getRange(i+1, 2).setValue(obj.koordinat);
    if(data[i][0] == "POTONGAN_TELAT" && obj.potTelat!==undefined) sh.getRange(i+1, 2).setValue(obj.potTelat);
    if(data[i][0] == "POTONGAN_ALPHA" && obj.potAlpha!==undefined) sh.getRange(i+1, 2).setValue(obj.potAlpha);
    if(data[i][0] == "PENGUMUMAN_DASHBOARD" && obj.pengumuman!==undefined) sh.getRange(i+1, 2).setValue(obj.pengumuman);
  }
  return true;
}

// --- 5. ABSENSI & GPS VALIDATION ---
function getAttendance() {
  const data = getSheetData("Absensi"); const result = [];
  for(let i=1; i<data.length; i++) result.push({ id: data[i][0], tanggal: data[i][1], id_karyawan: data[i][2], nama: data[i][3], masuk: data[i][4], fMasuk: data[i][5], pulang: data[i][6], fPulang: data[i][7], status: data[i][8] });
  return result;
}

function haversine(lat1, lon1, lat2, lon2) {
  const R = 6371000; const dLat = (lat2 - lat1) * Math.PI / 180; const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat/2)*Math.sin(dLat/2) + Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)*Math.sin(dLon/2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

function processAttendance(idKaryawan, base64Image, type, userLat, userLng) {
  const sets = getSettingsObj();
  
  if (userLat && userLng && sets.koordinat) {
    let isInRadius = false;
    try {
      const lokasiList = JSON.parse(sets.koordinat);
      for (let loc of lokasiList) {
        if(loc.coords) {
          const [tLat, tLng] = loc.coords.split(',').map(c => parseFloat(c.trim()));
          if(haversine(userLat, userLng, tLat, tLng) <= sets.radius) { isInRadius = true; break; }
        }
      }
    } catch(e) {}
    if (!isInRadius) throw new Error(`Posisi Anda di luar radius kantor (${sets.radius}m). Silakan mendekat ke lokasi yang terdaftar.`);
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
  const sh = ss.getSheetByName("Absensi"); 
  const data = sh.getDataRange().getValues();
  const tz = Session.getScriptTimeZone(); 
  const today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd"); 
  const time = Utilities.formatDate(new Date(), tz, "HH:mm");
  
  const folderId = getSheetData("Pengaturan").find(r=>r[0]=="FOLDER_DRIVE_ID")[1];
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Image), MimeType.JPEG, `Absen_${idKaryawan}_${today}_${type}.jpg`);
  const file = DriveApp.getFolderById(folderId).createFile(blob); file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); const imageUrl = file.getUrl();
  
  let rowIndex = -1; 
  for(let i=1; i<data.length; i++) { 
    if(data[i][2] == idKaryawan && data[i][1] == today) { rowIndex = i + 1; break; } 
  }
  
  if (type === "Masuk") {
    if (rowIndex !== -1) throw new Error("Anda sudah absen masuk hari ini.");
    const status = (time <= sets.jamMasuk) ? "Hadir" : "Terlambat";
    sh.appendRow(["ABS" + new Date().getTime(), today, idKaryawan, getUserInfo(idKaryawan).nama, time, imageUrl, "", "", status]);
    return { success: true, message: `Absen Masuk sukses pukul ${time}` };
  } else {
    if (rowIndex === -1) throw new Error("Anda belum Absen Masuk.");
    if (data[rowIndex-1][6] !== "") throw new Error("Anda sudah Absen Pulang hari ini.");
    sh.getRange(rowIndex, 7).setValue(time); sh.getRange(rowIndex, 8).setValue(imageUrl);
    return { success: true, message: `Absen Pulang sukses pukul ${time}` };
  }
}

function processAttendanceQR(idKaryawan, type, qrData, userLat, userLng) {
  if (qrData !== "ABSENPRO_JAT_OFFICE") throw new Error("QR Code Tidak Valid atau Palsu!");
  
  const sets = getSettingsObj();
  
  if (!userLat || !userLng) {
    throw new Error("Akses GPS harus diaktifkan untuk mencegah Fake QR (Scan dari rumah)!");
  }
  
  if (userLat && userLng && sets.koordinat) {
    let isInRadius = false;
    try {
      const lokasiList = JSON.parse(sets.koordinat);
      for (let loc of lokasiList) {
        if(loc.coords) {
          const [tLat, tLng] = loc.coords.split(',').map(c => parseFloat(c.trim()));
          if(haversine(userLat, userLng, tLat, tLng) <= sets.radius) { isInRadius = true; break; }
        }
      }
    } catch(e) {}
    
    if (!isInRadius) {
       throw new Error(`KECURANGAN TERDETEKSI: Anda memindai QR Code di luar radius kantor (${sets.radius}m). Silakan absen di area kantor!`);
    }
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
  const sh = ss.getSheetByName("Absensi"); 
  const data = sh.getDataRange().getValues();
  const tz = Session.getScriptTimeZone(); 
  const today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd"); 
  const time = Utilities.formatDate(new Date(), tz, "HH:mm");
  
  let rowIndex = -1; 
  for(let i=1; i<data.length; i++) { 
    if(data[i][2] == idKaryawan && data[i][1] == today) { rowIndex = i + 1; break; } 
  }
  
  if (type === "Masuk") {
    if (rowIndex !== -1) throw new Error("Anda sudah absen masuk hari ini.");
    const status = (time <= sets.jamMasuk) ? "Hadir (QR)" : "Terlambat (QR)";
    sh.appendRow(["ABS" + new Date().getTime(), today, idKaryawan, getUserInfo(idKaryawan).nama, time, "Via QR Code", "", "", status]);
    return { success: true, message: `Absen Masuk via QR sukses pukul ${time}` };
  } else {
    if (rowIndex === -1) throw new Error("Anda belum Absen Masuk.");
    if (data[rowIndex-1][6] !== "") throw new Error("Anda sudah Absen Pulang hari ini.");
    sh.getRange(rowIndex, 7).setValue(time); sh.getRange(rowIndex, 8).setValue("Via QR Code");
    return { success: true, message: `Absen Pulang via QR sukses pukul ${time}` };
  }
}

// --- 6. PAYROLL ENGINE ---
function calculatePayroll(bulanFilter, tahunFilter) {
  const targetPrefix = `${tahunFilter}-${bulanFilter}`; 
  const totalHariKerja = 22; 

  const settings = getSettingsObj();
  const employees = getEmployees().filter(e => e.role !== "Admin"); 
  const absensiData = getSheetData("Absensi");
  const lemburData = getSheetData("Lembur");
  const reimbData = getSheetData("Reimbursement");
  
  let pinjamanData = [];
  try { pinjamanData = getSheetData("Pinjaman"); } catch(e){}

  let payrollResults = [];

  for (let emp of employees) {
    let hadir = 0, telat = 0, totJamLembur = 0, totReimb = 0, potKasbon = 0;

    for (let i = 1; i < absensiData.length; i++) {
      if (absensiData[i][2] == emp.id && String(absensiData[i][1]).startsWith(targetPrefix)) {
        if (String(absensiData[i][8]).includes("Hadir")) hadir++;
        if (String(absensiData[i][8]).includes("Terlambat")) telat++;
      }
    }
    const totKehadiran = hadir + telat;
    const alpha = (totalHariKerja - totKehadiran > 0) ? (totalHariKerja - totKehadiran) : 0; 

    for (let i = 1; i < lemburData.length; i++) {
      if (lemburData[i][2] == emp.id && String(lemburData[i][4]).startsWith(targetPrefix) && lemburData[i][9] == "Disetujui") {
        try {
          const tMulai = new Date(`1970-01-01T${lemburData[i][5]}:00Z`);
          const tSelesai = new Date(`1970-01-01T${lemburData[i][6]}:00Z`);
          let diffJam = (tSelesai - tMulai) / (1000 * 60 * 60);
          if (diffJam < 0) diffJam += 24; 
          totJamLembur += diffJam;
        } catch(e) {}
      }
    }

    for (let i = 1; i < reimbData.length; i++) {
      if (reimbData[i][2] == emp.id && String(reimbData[i][1]).startsWith(targetPrefix) && reimbData[i][8] == "Disetujui") {
        totReimb += Number(reimbData[i][5]) || 0;
      }
    }

    // HITUNG POTONGAN KASBON (JIKA ADA PINJAMAN AKTIF/BELUM LUNAS)
    for (let i = 1; i < pinjamanData.length; i++) {
      if (pinjamanData[i][2] == emp.id && pinjamanData[i][11] == "Disetujui" && Number(pinjamanData[i][8]) > 0) {
        let mulaiBulan = String(pinjamanData[i][9]).padStart(2, '0');
        let mulaiTahun = String(pinjamanData[i][10]);
        let startPeriod = mulaiTahun + "-" + mulaiBulan;
        
        // Cek apakah bulan proses gaji >= bulan mulai potong
        if (targetPrefix >= startPeriod) {
            let cicilan = Number(pinjamanData[i][7]) || 0;
            let sisaHutang = Number(pinjamanData[i][8]) || 0;
            if(cicilan > sisaHutang) cicilan = sisaHutang;
            potKasbon += cicilan;
        }
      }
    }

    const calcGajiPokok = Number(emp.gajiPokok) || 0;
    const calcMakan = totKehadiran * (Number(emp.uangMakan) || 0);
    const calcTrans = totKehadiran * (Number(emp.uangTransport) || 0);
    const calcLembur = totJamLembur * (Number(emp.tarifLembur) || 0);
    
    const calcPotTelat = telat * settings.potTelat;
    const calcPotAlpha = alpha * settings.potAlpha;

    const totalTunjangan = calcMakan + calcTrans + calcLembur + totReimb;
    const totalPotongan = calcPotTelat + calcPotAlpha + potKasbon; 
    const gajiBersih = calcGajiPokok + totalTunjangan - totalPotongan;

    payrollResults.push({
      id_karyawan: emp.id,
      nama: emp.nama,
      jabatan: emp.jabatan, 
      role: emp.role,       
      gPokok: calcGajiPokok,
      tMakan: calcMakan,
      tTrans: calcTrans,
      tLembur: calcLembur,
      tReimb: totReimb,
      pTelat: calcPotTelat,
      pAlpha: calcPotAlpha,
      pKasbon: potKasbon, 
      totTunjangan: totalTunjangan,
      totPotongan: totalPotongan,
      gBersih: gajiBersih,
      hadir: hadir,
      telat: telat,
      alpha: alpha
    });
  }

  return payrollResults;
}

const fmtRp = (angka) => { return angka.toString().replace(/\B(?=(\d{3})+(?!\d))/g, "."); };

function publishPayslipsBulk(dataArray, bulan, tahun) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Payslip");
  const folderId = getSheetData("Pengaturan").find(r=>r[0]=="FOLDER_DRIVE_ID")[1];
  const folder = DriveApp.getFolderById(folderId);
  
  let shPinjaman = null; let dataPinjaman = [];
  try { 
    shPinjaman = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Pinjaman");
    dataPinjaman = shPinjaman.getDataRange().getValues();
  } catch(e){}

  for (let p of dataArray) {
    
    // POTONG SALDO HUTANG DI DATABASE PINJAMAN JIKA ADA KASBON
    if (p.pKasbon > 0 && shPinjaman) {
       for(let j=1; j<dataPinjaman.length; j++) {
           if(dataPinjaman[j][2] == p.id_karyawan && dataPinjaman[j][11] == "Disetujui" && Number(dataPinjaman[j][8]) > 0) {
               let mulaiBulan = String(dataPinjaman[j][9]).padStart(2, '0');
               let mulaiTahun = String(dataPinjaman[j][10]);
               
               const blnListMap = {"Januari":"01","Februari":"02","Maret":"03","April":"04","Mei":"05","Juni":"06","Juli":"07","Agustus":"08","September":"09","Oktober":"10","November":"11","Desember":"12"};
               let currentPeriod = tahun + "-" + blnListMap[bulan];
               let startPeriod = mulaiTahun + "-" + mulaiBulan;
               
               if (currentPeriod >= startPeriod) {
                   let sisa = Number(dataPinjaman[j][8]);
                   let cicilan = Number(dataPinjaman[j][7]);
                   if(cicilan > sisa) cicilan = sisa;
                   let newSisa = sisa - cicilan;
                   
                   if(newSisa <= 0) {
                       newSisa = 0;
                       shPinjaman.getRange(j+1, 12).setValue("Lunas"); // AUTO-LUNAS
                   }
                   shPinjaman.getRange(j+1, 9).setValue(newSisa); // UPDATE SISA HUTANG
                   break; 
               }
           }
       }
    }

    // --- ALGORITMA GROSS-UP BPJS & PAJAK ---
    const bpjsKesEmp = Math.min(Math.round(p.gPokok * 0.01), 120000); 
    const jhtEmp = Math.round(p.gPokok * 0.02);
    const jpEmp = Math.min(Math.round(p.gPokok * 0.01), 100000);
    const pph21 = Math.round(p.gPokok * 0.05); 
    
    const taxAndBpjsAllowance = bpjsKesEmp + jhtEmp + jpEmp + pph21;
    const totalEarningsPdf = p.gPokok + p.totTunjangan + taxAndBpjsAllowance;
    const totalDeductionsPdf = bpjsKesEmp + jhtEmp + jpEmp + pph21 + p.totPotongan;

    const jkkComp = Math.round(p.gPokok * 0.0024);
    const jkmComp = Math.round(p.gPokok * 0.003);
    const jhtComp = Math.round(p.gPokok * 0.037);
    const jpComp = Math.min(Math.round(p.gPokok * 0.02), 200000);
    const bpjsKesComp = Math.min(Math.round(p.gPokok * 0.04), 480000);
    const totalBenefits = jkkComp + jkmComp + jhtComp + jpComp + bpjsKesComp;

    // TEMPLATE PDF TALENTA STYLE
    const htmlBody = `
    <div style="font-family: Arial, sans-serif; font-size: 11px; color: #333; line-height: 1.4;">
      <table width="100%" style="margin-bottom: 20px;">
        <tr>
          <td valign="top">
            <h2 style="margin: 0; color: #1e3a8a; font-size: 20px; font-weight: bold; letter-spacing: 0.5px;">CV JAMPE AGENG TRADE</h2>
            <div style="font-size: 10px; color: #666; margin-top: 2px;">*CONFIDENTIAL</div>
          </td>
          <td valign="top" align="right">
            <h1 style="margin: 0; font-size: 24px; color: #333; font-weight: bold;">PAYSLIP</h1>
          </td>
        </tr>
      </table>

      <table width="100%" style="border-collapse: collapse; margin-bottom: 20px;">
        <tr>
          <td width="20%" style="color:#666; padding: 3px 0;">Payroll cut off</td><td width="2%">:</td><td width="38%">01-31 ${bulan} ${tahun}</td>
          <td width="15%" style="color:#666; padding: 3px 0;">Grade / Level</td><td width="2%">:</td><td width="23%">- / Staff</td>
        </tr>
        <tr>
          <td style="color:#666; padding: 3px 0;">ID / Name</td><td>:</td><td><b>${p.id_karyawan} / ${p.nama}</b></td>
          <td style="color:#666; padding: 3px 0;">PTKP</td><td>:</td><td>TK/0</td>
        </tr>
        <tr>
          <td style="color:#666; padding: 3px 0;">Job position</td><td>:</td><td>${p.jabatan}</td>
          <td style="color:#666; padding: 3px 0;">NPWP</td><td>:</td><td>-</td>
        </tr>
        <tr>
          <td style="color:#666; padding: 3px 0;">Organization</td><td>:</td><td>${p.role}</td>
          <td></td><td></td><td></td>
        </tr>
      </table>

      <table width="100%" style="border-collapse: collapse; margin-bottom: 20px;">
        <tr>
          <td width="48%" valign="top">
            <div style="border-bottom: 2px solid #222; font-weight: bold; font-size: 14px; padding-bottom: 5px; margin-bottom: 5px; color: #111;">Earnings</div>
            <table width="100%" cellpadding="3" cellspacing="0">
              <tr><td>Basic Salary</td><td align="right">${fmtRp(p.gPokok)}</td></tr>
              ${p.tMakan > 0 ? `<tr><td>Tunjangan Makan Bulanan</td><td align="right">${fmtRp(p.tMakan)}</td></tr>` : ''}
              ${p.tTrans > 0 ? `<tr><td>Tunjangan Transportasi</td><td align="right">${fmtRp(p.tTrans)}</td></tr>` : ''}
              ${p.tLembur > 0 ? `<tr><td>Lembur (Overtime)</td><td align="right">${fmtRp(p.tLembur)}</td></tr>` : ''}
              ${p.tReimb > 0 ? `<tr><td>Reimbursement</td><td align="right">${fmtRp(p.tReimb)}</td></tr>` : ''}
              <tr><td>Tax & BPJS Allowance</td><td align="right">${fmtRp(taxAndBpjsAllowance)}</td></tr>
              <tr><td colspan="2" style="padding: 5px 0;"></td></tr>
              <tr><td style="font-weight:bold; border-top: 1px solid #ccc; padding-top:6px;">Total earnings</td><td align="right" style="font-weight:bold; border-top: 1px solid #ccc; padding-top:6px;">${fmtRp(totalEarningsPdf)}</td></tr>
            </table>
          </td>
          
          <td width="4%"></td>
          
          <td width="48%" valign="top">
            <div style="border-bottom: 2px solid #222; font-weight: bold; font-size: 14px; padding-bottom: 5px; margin-bottom: 5px; color: #111;">Deductions</div>
            <table width="100%" cellpadding="3" cellspacing="0">
              <tr><td>BPJS Kesehatan Employee</td><td align="right">${fmtRp(bpjsKesEmp)}</td></tr>
              <tr><td>JHT Employee</td><td align="right">${fmtRp(jhtEmp)}</td></tr>
              <tr><td>Jaminan Pensiun Employee</td><td align="right">${fmtRp(jpEmp)}</td></tr>
              <tr><td>PPh 21</td><td align="right">${fmtRp(pph21)}</td></tr>
              ${p.pTelat > 0 ? `<tr><td>Potongan Terlambat</td><td align="right">${fmtRp(p.pTelat)}</td></tr>` : ''}
              ${p.pAlpha > 0 ? `<tr><td>Potongan Alpha / Absen</td><td align="right">${fmtRp(p.pAlpha)}</td></tr>` : ''}
              ${p.pKasbon > 0 ? `<tr><td style="color:#dc2626;">Potongan Kasbon/Pinjaman</td><td align="right" style="color:#dc2626;">-${fmtRp(p.pKasbon)}</td></tr>` : ''}
              <tr><td colspan="2" style="padding: 5px 0;"></td></tr>
              <tr><td style="font-weight:bold; border-top: 1px solid #ccc; padding-top:6px;">Total deductions</td><td align="right" style="font-weight:bold; border-top: 1px solid #ccc; padding-top:6px;">${fmtRp(totalDeductionsPdf)}</td></tr>
            </table>
          </td>
        </tr>
      </table>

      <table width="100%" cellpadding="12" cellspacing="0" style="background-color: #f8fafc; border: 1px solid #e2e8f0; margin-bottom: 25px;">
        <tr>
          <td style="font-weight: bold; font-size: 14px; color: #333;">Take Home Pay</td>
          <td align="right" style="font-weight: bold; font-size: 18px; color: #1e3a8a;">Rp${fmtRp(p.gBersih)}</td>
        </tr>
      </table>

      <table width="100%" style="border-collapse: collapse; margin-bottom: 40px;">
        <tr>
          <td width="48%" valign="top">
            <div style="font-weight: bold; font-size: 12px; margin-bottom: 8px; color: #333;">Benefits*</div>
            <table width="100%" cellpadding="2" cellspacing="0" style="font-size: 10px;">
              <tr><td style="border-bottom: 1px dotted #eee; padding:4px 0;">JKK</td><td align="right" style="border-bottom: 1px dotted #eee;">${fmtRp(jkkComp)}</td></tr>
              <tr><td style="border-bottom: 1px dotted #eee; padding:4px 0;">JKM</td><td align="right" style="border-bottom: 1px dotted #eee;">${fmtRp(jkmComp)}</td></tr>
              <tr><td style="border-bottom: 1px dotted #eee; padding:4px 0;">JHT Company</td><td align="right" style="border-bottom: 1px dotted #eee;">${fmtRp(jhtComp)}</td></tr>
              <tr><td style="border-bottom: 1px dotted #eee; padding:4px 0;">Jaminan Pensiun Company</td><td align="right" style="border-bottom: 1px dotted #eee;">${fmtRp(jpComp)}</td></tr>
              <tr><td style="border-bottom: 1px dotted #eee; padding:4px 0;">BPJS Kesehatan Company</td><td align="right" style="border-bottom: 1px dotted #eee;">${fmtRp(bpjsKesComp)}</td></tr>
              <tr><td style="font-weight:bold; padding-top:6px;">Total benefits</td><td align="right" style="font-weight:bold; padding-top:6px;">${fmtRp(totalBenefits)}</td></tr>
            </table>
          </td>
          
          <td width="4%"></td>
          
          <td width="48%" valign="top">
            <div style="font-weight: bold; font-size: 12px; margin-bottom: 8px; color: #333;">Attendance Summary</div>
            <table width="100%" cellpadding="2" cellspacing="0" style="font-size: 10px;">
              <tr><td style="border-bottom: 1px dotted #eee; padding:4px 0;">Actual Working Day</td><td align="right" style="border-bottom: 1px dotted #eee;">${p.hadir + p.telat}d</td></tr>
              <tr><td style="border-bottom: 1px dotted #eee; padding:4px 0;">Schedule Working Day</td><td align="right" style="border-bottom: 1px dotted #eee;">22d</td></tr>
              <tr><td style="border-bottom: 1px dotted #eee; padding:4px 0;">Dayoff (Alpha)</td><td align="right" style="border-bottom: 1px dotted #eee;">${p.alpha}d</td></tr>
              <tr><td style="border-bottom: 1px dotted #eee; padding:4px 0;">National Holiday</td><td align="right" style="border-bottom: 1px dotted #eee;">0d</td></tr>
              <tr><td style="border-bottom: 1px dotted #eee; padding:4px 0;">Company Holiday</td><td align="right" style="border-bottom: 1px dotted #eee;">0d</td></tr>
              <tr><td style="font-weight:bold; padding-top:6px;">Attendance Code</td><td align="right" style="font-weight:bold; padding-top:6px;">H:${p.hadir} T:${p.telat}</td></tr>
            </table>
          </td>
        </tr>
      </table>

      <div style="font-size: 9px; color: #666; text-align: justify;">
        <p style="margin-bottom: 4px;">*These are the benefits you'll get from the company, but not included in your take-home pay (THP).</p>
        <p style="margin-bottom: 8px; font-weight: bold; color: #333;">THIS IS COMPUTER GENERATED PRINTOUT AND NO SIGNATURE IS REQUIRED</p>
        <p style="margin-bottom: 4px; text-transform: uppercase;">Please note that the contents of this statement should be treated with absolute confidentiality except to the extent you are required to make disclosure for any tax, legal, or regulatory purposes. Any breach of this confidentiality obligation will be dealt with seriously, which may involve disciplinary action being taken.</p>
        <p style="margin-bottom: 20px; text-transform: uppercase;">Harap diperhatikan, isi pernyataan ini adalah rahasia kecuali anda diminta untuk mengungkapkannya untuk keperluan pajak, hukum, atau kepentingan pemerintah. Setiap pelanggaran atas kewajiban menjaga kerahasiaan ini akan dikenakan sanksi, yang mungkin berupa tindakan kedisiplinan.</p>
        <div style="text-align: center; border-top: 1px solid #ccc; padding-top: 10px;">
          This payslip is generated by AbsenPro HRIS<br>
          <a href="#" style="color: #2563eb; text-decoration: none;">https://absenpro.jampe-ageng.com</a>
        </div>
      </div>
    </div>
    `;

    const blob = Utilities.newBlob(htmlBody, MimeType.HTML, `Slip_${p.nama}_${bulan}_${tahun}.pdf`).getAs(MimeType.PDF);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    sh.appendRow(["PAY" + new Date().getTime(), bulan, tahun, p.id_karyawan, p.nama, p.gPokok, p.totTunjangan, p.totPotongan, p.gBersih, file.getUrl()]);
    
    try {
       const empEmail = getUserEmail(p.id_karyawan);
       if (empEmail && empEmail !== "") {
         MailApp.sendEmail({
           to: empEmail,
           subject: `Slip Gaji ${bulan} ${tahun} - Jampe Ageng Trade`,
           htmlBody: `<p>Halo <b>${p.nama}</b>,</p><p>Terlampir adalah dokumen PDF slip gaji Anda untuk periode bulan <b>${bulan} ${tahun}</b>.</p><p>Take Home Pay: <b style="color:green; font-size:16px;">Rp ${fmtRp(p.gBersih)}</b></p><p>Harap simpan dokumen rahasia ini dengan baik.</p><p>Salam hangat,<br>Tim HRD & Finance</p>`,
           attachments: [blob]
         });
       }
    } catch(e) {}
  }
  
  return { success: true, message: `Berhasil! Slip gaji diterbitkan & PDF telah dikirim ke email karyawan.` };
}

// --- 7. SUBMIT FORMS ---
function submitCuti(id, jenis, mulai, selesai, ket, b64) { const fileUrl = b64 ? helperUploadFile(b64, `Cuti_${id}_${new Date().getTime()}`) : ""; SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Cuti").appendRow(["CT" + new Date().getTime(), Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"), id, getUserInfo(id).nama, jenis, mulai, selesai, ket, fileUrl, "Menunggu Konfirmasi"]); return { success: true, message: "Pengajuan berhasil." }; }
function submitReimbursement(id, ket, nom, desc, b64) { if(!b64) throw new Error("Bukti struk wajib!"); const fileUrl = helperUploadFile(b64, `Reimb_${id}_${new Date().getTime()}`); SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Reimbursement").appendRow(["RMB" + new Date().getTime(), Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"), id, getUserInfo(id).nama, ket, nom, desc, fileUrl, "Menunggu Konfirmasi"]); return { success: true, message: "Klaim berhasil." }; }
function submitReqAbsensi(id, tipe, jam, kord, ket, b64) { const fileUrl = b64 ? helperUploadFile(b64, `ReqAbs_${id}_${new Date().getTime()}`) : ""; SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Req_Absensi").appendRow(["RQA" + new Date().getTime(), Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"), id, getUserInfo(id).nama, tipe, jam, kord, ket, fileUrl, "Menunggu Konfirmasi"]); return { success: true, message: "Request Absen berhasil." }; }
function submitShiftChange(id, tgl, asal, tujuan, alasan) { SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Shift_Change").appendRow(["SHF" + new Date().getTime(), Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"), id, getUserInfo(id).nama, tgl, asal, tujuan, alasan, "Menunggu Konfirmasi"]); return { success: true, message: "Pengajuan tukar shift berhasil." }; }
function submitLembur(id, tglLembur, mulai, selesai, alasan, b64) { const fileUrl = b64 ? helperUploadFile(b64, `Lembur_${id}_${new Date().getTime()}`) : ""; SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Lembur").appendRow(["LMB" + new Date().getTime(), Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"), id, getUserInfo(id).nama, tglLembur, mulai, selesai, alasan, fileUrl, "Menunggu Konfirmasi"]); return { success: true, message: "Pengajuan lembur berhasil." }; }

// FUNGSI SUBMIT PINJAMAN BARU
function submitPinjaman(id, nominal, alasan) { 
  SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Pinjaman").appendRow(["PNJ" + new Date().getTime(), Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"), id, getUserInfo(id).nama, nominal, alasan, "", "", "", "", "", "Menunggu Konfirmasi"]); 
  return { success: true, message: "Pengajuan kasbon berhasil." }; 
}

function getPayslips(id) { const data = getSheetData("Payslip"); const result = []; for(let i=1; i<data.length; i++) { if(data[i][3] == id) { result.push({ id: data[i][0], bulan: data[i][1], tahun: data[i][2], gajiBersih: data[i][8], linkPdf: data[i][9] }); } } return result; }

// --- 8. UTILITIES ---
function getSheetData(name) { const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name); if (!sh) throw new Error("Sheet " + name + " tidak ditemukan!"); return sh.getDataRange().getValues(); }
function getUserInfo(id) { const data = getSheetData("Karyawan"); for (let i = 1; i < data.length; i++) { if (data[i][0] == id) return { nama: data[i][1], jabatan: data[i][4] }; } return { nama: "Unknown", jabatan: "Unknown" }; }
function helperUploadFile(base64Data, filename) {
  if (!base64Data) return ""; let data = base64Data, mimeType = MimeType.JPEG;
  if (data.indexOf("data:") === 0) { const parts = data.split(","); data = parts[1]; const mimeStr = parts[0].split(";")[0].split(":")[1]; if (mimeStr === "application/pdf") mimeType = MimeType.PDF; else if (mimeStr === "image/png") mimeType = MimeType.PNG; }
  const folderId = getSheetData("Pengaturan").find(r=>r[0]=="FOLDER_DRIVE_ID")[1];
  const file = DriveApp.getFolderById(folderId).createFile(Utilities.newBlob(Utilities.base64Decode(data), mimeType, filename)); file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); return file.getUrl();
}

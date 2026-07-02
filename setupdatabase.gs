/**
 * ABSENPRO - JAMPE AGENG TRADE
 * Developer: Bot AppScript Ultimate
 * FILE 1: SETUP DATABASE (V5.6 + Modul Pinjaman/Kasbon)
 */

// --- 1. SETUP DATABASE OTOMATIS ---
function setupDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Helper untuk membuat sheet beserta header dan styling
  const createSheet = (name, headers, bg) => { 
    let sh = ss.getSheetByName(name); 
    if (!sh) { 
      sh = ss.insertSheet(name); 
      sh.appendRow(headers); 
      sh.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground(bg).setFontColor("white"); 
    } 
    return sh; 
  };

  // 1. Sheet Karyawan
  createSheet("Karyawan", ["ID_Karyawan", "Nama", "Username", "PIN", "Jabatan", "Role", "Gaji_Pokok", "Uang_Makan", "Uang_Transport", "Tarif_Lembur", "Email", "Kuota_Cuti"], "#10b981")
    .appendRow(["EMP001", "Administrator Utama", "admin", "123456", "HRD", "Admin", 0, 0, 0, 0, "", 12]); 
  
  // 2. Sheet Absensi
  createSheet("Absensi", ["ID_Absen", "Tanggal", "ID_Karyawan", "Nama", "Jam Masuk", "Foto Masuk", "Jam Pulang", "Foto Pulang", "Status"], "#10b981");
  
  // 3. Sheet Pengaturan
  let shSetting = createSheet("Pengaturan", ["Key", "Value", "Keterangan"], "#10b981");
  if (shSetting.getLastRow() === 1) {
    let folderId = "";
    try { 
      const folders = DriveApp.getFoldersByName("AbsenPro_Photos"); 
      folderId = folders.hasNext() ? folders.next().getId() : DriveApp.createFolder("AbsenPro_Photos").getId(); 
    } catch(e) { 
      folderId = "ERROR"; 
    }
    shSetting.appendRow(["JAM_BATAS_MASUK", "08:00", "Batas waktu absen masuk"]);
    shSetting.appendRow(["JAM_BATAS_PULANG", "17:00", "Batas waktu absen pulang"]);
    shSetting.appendRow(["RADIUS", "50", "Radius absen GPS (meter)"]);
    shSetting.appendRow(["KOORDINAT", '[{"nama":"Kantor Pusat","coords":"-6.200, 106.816"}]', "Lokasi Koordinat GPS"]);
    shSetting.appendRow(["POTONGAN_TELAT", "15000", "Denda jika Terlambat (Rp)"]);
    shSetting.appendRow(["POTONGAN_ALPHA", "50000", "Denda jika Tidak Hadir (Rp)"]);
    shSetting.appendRow(["PENGUMUMAN_DASHBOARD", "Selamat datang di AbsenPro! Silakan absen tepat waktu.", "Teks Pengumuman"]);
    shSetting.appendRow(["FOLDER_DRIVE_ID", folderId, "ID Folder Google Drive"]);
  } else {
    const keys = shSetting.getDataRange().getValues().map(r => r[0]);
    if(!keys.includes("POTONGAN_TELAT")) shSetting.appendRow(["POTONGAN_TELAT", "15000", "Denda jika Terlambat (Rp)"]);
    if(!keys.includes("POTONGAN_ALPHA")) shSetting.appendRow(["POTONGAN_ALPHA", "50000", "Denda jika Tidak Hadir (Rp)"]);
    if(!keys.includes("PENGUMUMAN_DASHBOARD")) shSetting.appendRow(["PENGUMUMAN_DASHBOARD", "", "Teks Pengumuman"]);
  }

  // 4. Sheet Modul Karyawan
  createSheet("Reimbursement", ["ID_Reimb", "Tanggal", "ID_Karyawan", "Nama", "Kategori", "Nominal", "Keterangan", "Bukti_Drive", "Status"], "#3b82f6");
  createSheet("Cuti", ["ID_Cuti", "Tgl_Pengajuan", "ID_Karyawan", "Nama", "Jenis_Cuti", "Tgl_Mulai", "Tgl_Selesai", "Keterangan", "Bukti_Drive", "Status"], "#3b82f6");
  createSheet("Payslip", ["ID_Payslip", "Bulan", "Tahun", "ID_Karyawan", "Nama", "Gaji_Pokok", "Total_Tunjangan", "Total_Potongan", "Gaji_Bersih", "Link_PDF"], "#3b82f6");
  createSheet("Req_Absensi", ["ID_Req", "Tanggal", "ID_Karyawan", "Nama", "Tipe_Absen", "Jam", "Lokasi_Koordinat", "Keterangan", "Bukti_Drive", "Status"], "#3b82f6");
  createSheet("Shift_Change", ["ID_Shift", "Tgl_Pengajuan", "ID_Karyawan", "Nama", "Tgl_Jadwal", "Shift_Asal", "Shift_Tujuan", "Alasan", "Status"], "#3b82f6");
  createSheet("Lembur", ["ID_Lembur", "Tgl_Pengajuan", "ID_Karyawan", "Nama", "Tgl_Lembur", "Jam_Mulai", "Jam_Selesai", "Alasan", "Bukti_Drive", "Status"], "#3b82f6");
  
  // 5. SHEET BARU: PINJAMAN / KASBON
  createSheet("Pinjaman", ["ID_Pinjaman", "Tgl_Pengajuan", "ID_Karyawan", "Nama", "Nominal_Pinjam", "Alasan", "Tenor_Bulan", "Cicilan_Per_Bulan", "Sisa_Hutang", "Mulai_Bulan", "Mulai_Tahun", "Status"], "#f59e0b");
  
  return "Setup Database Selesai!";
}

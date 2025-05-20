
// Aplikasi Laporan Harian - Backend Google Apps Script

// Konfigurasi
const SPREADSHEET_ID = "16ZxTAYF9HNi8la1ESPQ6fn2a6tjnSDecog2S02oHQ0A"; // Ganti dengan ID spreadsheet Anda

// Fungsi utama untuk menjalankan aplikasi web
function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Sistem Laporan Harian")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// Fungsi untuk menyertakan file eksternal
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Fungsi otentikasi
function validateLogin(username, password, role) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    
    // Lewati baris header
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username && 
          data[i][1] === password && 
          data[i][2] === role && 
          data[i][3] === true) { // Status aktif
        return {
          success: true,
          userData: {
            username: data[i][0],
            name: data[i][4],
            role: data[i][2],
            department: data[i][5]
          }
        };
      }
    }
    
    return { 
      success: false, 
      error: "Kredensial tidak valid atau akun tidak aktif." 
    };
  } catch (error) {
    return { 
      success: false, 
      error: "Kesalahan login: " + error.toString() 
    };
  }
}

// Fungsi manajemen laporan
function submitReport(reportData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Reports");
    
    if (!sheet) {
      sheet = ss.insertSheet("Reports");
      sheet.appendRow([
        "ID Laporan", 
        "Username", 
        "Tanggal", 
        "Judul",
        "Tugas Selesai", 
        "Kemajuan", 
        "Tantangan", 
        "Rencana Besok",
        "Catatan Tambahan",
        "Status",
        "Komentar Admin",
        "Waktu Pengiriman"
      ]);
    }
    
    const reportId = Utilities.getUuid();
    const timestamp = new Date();
    
    sheet.appendRow([
      reportId,
      reportData.username,
      reportData.date,
      reportData.title,
      reportData.tasksCompleted,
      reportData.progress,
      reportData.challenges,
      reportData.nextDayPlan,
      reportData.additionalNotes,
      "Menunggu", // Status default
      "", // Komentar admin kosong
      timestamp
    ]);
    
    return { 
      success: true, 
      message: "Laporan berhasil dikirim.",
      reportId: reportId
    };
  } catch (error) {
    return { 
      success: false, 
      error: "Kesalahan mengirim laporan: " + error.toString() 
    };
  }
}

function getUserReports(username) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Reports");
    
    if (!sheet) {
      return { success: true, reports: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    const reports = [];
    
    // Lewati baris header
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username) {
        reports.push({
          reportId: data[i][0],
          username: data[i][1],
          date: data[i][2],
          title: data[i][3],
          tasksCompleted: data[i][4],
          progress: data[i][5],
          challenges: data[i][6],
          nextDayPlan: data[i][7],
          additionalNotes: data[i][8],
          status: data[i][9],
          adminComments: data[i][10],
          timestamp: data[i][11]
        });
      }
    }
    
    // Urutkan laporan berdasarkan tanggal (terbaru lebih dulu)
    reports.sort((a, b) => new Date(b.date) - new Date(a.date));
    
    return { success: true, reports: reports };
  } catch (error) {
    return { 
      success: false, 
      error: "Kesalahan mengambil laporan: " + error.toString() 
    };
  }
}

function getAllReports() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Reports");
    
    if (!sheet) {
      return { success: true, reports: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    const reports = [];
    
    // Lewati baris header
    for (let i = 1; i < data.length; i++) {
      reports.push({
        reportId: data[i][0],
        username: data[i][1],
        date: data[i][2],
        title: data[i][3],
        tasksCompleted: data[i][4],
        progress: data[i][5],
        challenges: data[i][6],
        nextDayPlan: data[i][7],
        additionalNotes: data[i][8],
        status: data[i][9],
        adminComments: data[i][10],
        timestamp: data[i][11]
      });
    }
    
    // Urutkan laporan berdasarkan tanggal (terbaru lebih dulu)
    reports.sort((a, b) => new Date(b.date) - new Date(a.date));
    
    return { success: true, reports: reports };
  } catch (error) {
    return { 
      success: false, 
      error: "Kesalahan mengambil laporan: " + error.toString() 
    };
  }
}

function updateReportStatus(reportId, status, adminComments) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Reports");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === reportId) {
        sheet.getRange(i + 1, 10).setValue(status);
        sheet.getRange(i + 1, 11).setValue(adminComments);
        return { 
          success: true, 
          message: "Status laporan berhasil diperbarui." 
        };
      }
    }
    
    return { 
      success: false, 
      error: "Laporan tidak ditemukan." 
    };
  } catch (error) {
    return { 
      success: false, 
      error: "Kesalahan memperbarui status laporan: " + error.toString() 
    };
  }
}

// Fungsi manajemen karyawan (hanya Admin)
function getAllEmployees() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    const employees = [];
    
    // Lewati baris header
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === "employee") {
        employees.push({
          username: data[i][0],
          isActive: data[i][3],
          name: data[i][4],
          department: data[i][5],
          email: data[i][6]
        });
      }
    }
    
    return { success: true, employees: employees };
  } catch (error) {
    return { 
      success: false, 
      error: "Kesalahan mengambil data karyawan: " + error.toString() 
    };
  }
}

function addEmployee(employeeData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Users");
    
    if (!sheet) {
      sheet = ss.insertSheet("Users");
      sheet.appendRow([
        "Username", 
        "Password", 
        "Peran", 
        "StatusAktif",
        "Nama", 
        "Departemen", 
        "Email"
      ]);
    }
    
    // Periksa apakah username sudah ada
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === employeeData.username) {
        return { 
          success: false, 
          error: "Username sudah digunakan." 
        };
      }
    }
    
    sheet.appendRow([
      employeeData.username,
      employeeData.password,
      "employee",
      true,
      employeeData.name,
      employeeData.department,
      employeeData.email
    ]);
    
    return { 
      success: true, 
      message: "Karyawan berhasil ditambahkan." 
    };
  } catch (error) {
    return { 
      success: false, 
      error: "Kesalahan menambahkan karyawan: " + error.toString() 
    };
  }
}

function updateEmployeeStatus(username, isActive) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username && data[i][2] === "employee") {
        sheet.getRange(i + 1, 4).setValue(isActive);
        return { 
          success: true, 
          message: isActive ? "Karyawan diaktifkan." : "Karyawan dinonaktifkan." 
        };
      }
    }
    
    return { 
      success: false, 
      error: "Karyawan tidak ditemukan." 
    };
  } catch (error) {
    return { 
      success: false, 
      error: "Kesalahan memperbarui status karyawan: " + error.toString() 
    };
  }
}

// Fungsi statistik dashboard
function getDashboardStats(username, role) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const reportsSheet = ss.getSheetByName("Reports");
    
    if (!reportsSheet) {
      return { 
        success: true, 
        stats: {
          totalReports: 0,
          pendingReports: 0,
          approvedReports: 0,
          rejectedReports: 0,
          weeklySubmissions: [0, 0, 0, 0, 0, 0, 0] 
        }
      };
    }
    
    const reportsData = reportsSheet.getDataRange().getValues();
    let totalReports = 0;
    let pendingReports = 0;
    let approvedReports = 0;
    let rejectedReports = 0;
    let weeklySubmissions = [0, 0, 0, 0, 0, 0, 0]; // Min-Sen
    
    const today = new Date();
    const oneWeekAgo = new Date(today);
    oneWeekAgo.setDate(today.getDate() - 7);
    
    // Lewati baris header
    for (let i = 1; i < reportsData.length; i++) {
      const reportDate = new Date(reportsData[i][2]);
      const reportUsername = reportsData[i][1];
      const reportStatus = reportsData[i][9];
      
      // Filter berdasarkan username untuk karyawan, hitung semua untuk admin
      if (role === "admin" || reportUsername === username) {
        totalReports++;
        
        if (reportStatus === "Menunggu") {
          pendingReports++;
        } else if (reportStatus === "Disetujui") {
          approvedReports++;
        } else if (reportStatus === "Ditolak") {
          rejectedReports++;
        }
        
        // Periksa apakah laporan berada dalam seminggu terakhir
        if (reportDate >= oneWeekAgo && reportDate <= today) {
          const dayOfWeek = reportDate.getDay(); // 0 = Minggu, 6 = Sabtu
          weeklySubmissions[dayOfWeek]++;
        }
      }
    }
    
    return { 
      success: true, 
      stats: {
        totalReports: totalReports,
        pendingReports: pendingReports,
        approvedReports: approvedReports,
        rejectedReports: rejectedReports,
        weeklySubmissions: weeklySubmissions
      }
    };
  } catch (error) {
    return { 
      success: false, 
      error: "Kesalahan mengambil statistik dashboard: " + error.toString() 
    };
  }
}

// Inisialisasi spreadsheet dengan data uji
function initializeApp() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Buat sheet Users jika belum ada
    let usersSheet = ss.getSheetByName("Users");
    if (!usersSheet) {
      usersSheet = ss.insertSheet("Users");
      usersSheet.appendRow([
        "Username", 
        "Password", 
        "Peran", 
        "StatusAktif",
        "Nama", 
        "Departemen", 
        "Email"
      ]);
      
      // Tambahkan admin dan karyawan uji
      usersSheet.appendRow([
        "admin",
        "admin123",
        "admin",
        true,
        "Pengguna Admin",
        "Manajemen",
        "admin@example.com"
      ]);
      
      usersSheet.appendRow([
        "karyawan",
        "karyawan123",
        "employee",
        true,
        "Karyawan Contoh",
        "Pengembangan",
        "karyawan@example.com"
      ]);
    }
    
    // Buat sheet Reports jika belum ada
    let reportsSheet = ss.getSheetByName("Reports");
    if (!reportsSheet) {
      reportsSheet = ss.insertSheet("Reports");
      reportsSheet.appendRow([
        "ID Laporan", 
        "Username", 
        "Tanggal", 
        "Judul",
        "Tugas Selesai", 
        "Kemajuan", 
        "Tantangan", 
        "Rencana Besok",
        "Catatan Tambahan",
        "Status",
        "Komentar Admin",
        "Waktu Pengiriman"
      ]);
      
      // Tambahkan contoh laporan
      const today = new Date();
      reportsSheet.appendRow([
        Utilities.getUuid(),
        "karyawan",
        Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        "Contoh Laporan Harian",
        "Menyelesaikan tugas 1 dan tugas 2",
        "Proyek A 50% selesai",
        "Mengalami masalah dengan integrasi API",
        "Melanjutkan pekerjaan pada Proyek A",
        "Tidak ada",
        "Menunggu",
        "",
        today
      ]);
    }
    
    return { 
      success: true, 
      message: "Aplikasi berhasil diinisialisasi dengan data contoh." 
    };
  } catch (error) {
    return { 
      success: false, 
      error: "Kesalahan inisialisasi aplikasi: " + error.toString() 
    };
  }
}
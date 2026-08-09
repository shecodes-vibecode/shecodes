// Memuat Halaman Utama
function doGet() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('NutriFit Tracker');
  
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

// Format Tanggal Hari Ini (DD/MM/YYYY)
function getTanggalHariIni() {
  var d = new Date();
  var day = ("0" + d.getDate()).slice(-2);
  var month = ("0" + (d.getMonth() + 1)).slice(-2);
  var year = d.getFullYear();
  return day + "/" + month + "/" + year;
}

// Pencarian Nutrisi Otomatis
function cariNutrisiOtomatis(nama) {
  if (!nama) return { sukses: false, pesan: "Nama makanan tidak boleh kosong." };
  
  var n = nama.toLowerCase().trim();
  
  var database = {
    'nasi padang': { kalori: 650, protein: 25, karbo: 65, lemak: 30 },
    'rendang': { kalori: 460, protein: 26, karbo: 8, lemak: 37 },
    'nasi goreng': { kalori: 450, protein: 12, karbo: 55, lemak: 18 },
    'nasi uduk': { kalori: 380, protein: 8, karbo: 52, lemak: 15 },
    'nasi kuning': { kalori: 350, protein: 7, karbo: 48, lemak: 14 },
    'nasi putih': { kalori: 204, protein: 4, karbo: 44, lemak: 0.4 },
    'nasi merah': { kalori: 110, protein: 2.6, karbo: 23, lemak: 0.9 },
    'lontong sayur': { kalori: 360, protein: 9, karbo: 50, lemak: 13 },
    'bubur ayam': { kalori: 290, protein: 10, karbo: 45, lemak: 8 },
    'mie ayam': { kalori: 400, protein: 14, karbo: 50, lemak: 14 },
    'mie goreng': { kalori: 380, protein: 9, karbo: 54, lemak: 14 },
    'mie instan': { kalori: 350, protein: 8, karbo: 46, lemak: 15 },
    'bakso': { kalori: 350, protein: 16, karbo: 35, lemak: 12 },
    'kwetiau goreng': { kalori: 470, protein: 13, karbo: 58, lemak: 20 },
    'bihun goreng': { kalori: 310, protein: 6, karbo: 48, lemak: 11 },
    'soto ayam': { kalori: 310, protein: 18, karbo: 25, lemak: 14 },
    'soto daging': { kalori: 350, protein: 22, karbo: 18, lemak: 20 },
    'rawon': { kalori: 330, protein: 20, karbo: 12, lemak: 22 },
    'ayam goreng': { kalori: 260, protein: 24, karbo: 2, lemak: 17 },
    'ayam bakar': { kalori: 220, protein: 25, karbo: 4, lemak: 11 },
    'ayam geprek': { kalori: 380, protein: 22, karbo: 15, lemak: 25 },
    'sate ayam': { kalori: 340, protein: 20, karbo: 15, lemak: 22 },
    'sate kambing': { kalori: 380, protein: 24, karbo: 8, lemak: 28 },
    'telur rebus': { kalori: 75, protein: 6, karbo: 0.5, lemak: 5 },
    'telur dadar': { kalori: 110, protein: 7, karbo: 1, lemak: 9 },
    'telur ceplok': { kalori: 92, protein: 6.3, karbo: 0.4, lemak: 7 },
    'ikan goreng': { kalori: 200, protein: 20, karbo: 2, lemak: 12 },
    'ikan bakar': { kalori: 160, protein: 22, karbo: 1, lemak: 7 },
    'tahu goreng': { kalori: 78, protein: 8, karbo: 2, lemak: 5 },
    'tempe goreng': { kalori: 118, protein: 7, karbo: 9, lemak: 7 },
    'tempe bacem': { kalori: 120, protein: 6, karbo: 13, lemak: 5 },
    'gado-gado': { kalori: 318, protein: 12, karbo: 38, lemak: 14 },
    'pecel': { kalori: 280, protein: 10, karbo: 32, lemak: 12 },
    'capcay': { kalori: 150, protein: 6, karbo: 18, lemak: 6 },
    'sayur sop': { kalori: 90, protein: 3, karbo: 14, lemak: 2 },
    'sayur asam': { kalori: 80, protein: 2, karbo: 15, lemak: 1.5 },
    'gorengan': { kalori: 140, protein: 2, karbo: 16, lemak: 8 },
    'martabak telur': { kalori: 320, protein: 12, karbo: 24, lemak: 19 },
    'martabak manis': { kalori: 270, protein: 6, karbo: 38, lemak: 11 },
    'roti tawar': { kalori: 80, protein: 3, karbo: 15, lemak: 1 },
    'pisang goreng': { kalori: 150, protein: 1.5, karbo: 25, lemak: 5.5 },
    'es teh manis': { kalori: 90, protein: 0, karbo: 22, lemak: 0 },
    'kopi susu': { kalori: 130, protein: 3, karbo: 18, lemak: 5 },
    'jus alpukat': { kalori: 230, protein: 2, karbo: 28, lemak: 13 },
    'susu sapi': { kalori: 150, protein: 8, karbo: 12, lemak: 8 }
  };

  for (var kunci in database) {
    if (n.includes(kunci)) {
      var item = database[kunci];
      return {
        sukses: true,
        kalori: item.kalori,
        protein: item.protein,
        karbo: item.karbo,
        lemak: item.lemak
      };
    }
  }

  return {
    sukses: false,
    pesan: "Makanan tidak ditemukan di kamus lokal. Silakan isi angka nutrisi secara manual."
  };
}

// Menyimpan Makanan
function simpanMakanan(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Log_Makanan");
    if (!sheet) {
      sheet = ss.insertSheet("Log_Makanan");
      sheet.appendRow(["Email", "Tanggal", "Nama Makanan", "Kalori", "Protein", "Karbo", "Lemak"]);
    }
    
    var email = Session.getActiveUser().getEmail() || "User";
    var tgl = getTanggalHariIni();
    
    sheet.appendRow([
      email, tgl, data.nama, 
      Number(data.kalori) || 0, 
      Number(data.protein) || 0, 
      Number(data.karbo) || 0, 
      Number(data.lemak) || 0
    ]);
    
    return { sukses: true, rekap: getRingkasanHariIni() };
  } catch(e) {
    return { sukses: false, error: e.toString() };
  }
}

// Menyimpan Aktivitas
function simpanAktivitas(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Log_Aktivitas");
    if (!sheet) {
      sheet = ss.insertSheet("Log_Aktivitas");
      sheet.appendRow(["Email", "Tanggal", "Jenis Aktivitas", "Durasi", "Kalori Terbakar"]);
    }
    
    var email = Session.getActiveUser().getEmail() || "User";
    var tgl = getTanggalHariIni();
    
    sheet.appendRow([
      email, tgl, data.jenis, data.durasi, 
      Number(data.kalori) || 0
    ]);
    
    return { sukses: true, rekap: getRingkasanHariIni() };
  } catch(e) {
    return { sukses: false, error: e.toString() };
  }
}

// Menambah Air Minum
function tambahAir(jumlah) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Log_Air");
    if (!sheet) {
      sheet = ss.insertSheet("Log_Air");
      sheet.appendRow(["Email", "Tanggal", "Jumlah Air (ml)"]);
    }
    
    var email = Session.getActiveUser().getEmail() || "User";
    var tgl = getTanggalHariIni();
    
    sheet.appendRow([email, tgl, Number(jumlah) || 0]);
    
    return { sukses: true, rekap: getRingkasanHariIni() };
  } catch(e) {
    return { sukses: false, error: e.toString() };
  }
}

// Mengambil Ringkasan Hari Ini
function getRingkasanHariIni() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tglHariIni = getTanggalHariIni();
  
  var totalKalori = 0, totalProtein = 0, totalKarbo = 0, totalLemak = 0, totalBakar = 0, totalAir = 0;

  var sheetMakanan = ss.getSheetByName("Log_Makanan");
  if (sheetMakanan && sheetMakanan.getLastRow() > 1) {
    var dataMakanan = sheetMakanan.getRange(2, 1, sheetMakanan.getLastRow() - 1, 7).getValues();
    for (var i = 0; i < dataMakanan.length; i++) {
      if (formatDate(dataMakanan[i][1]) === tglHariIni) {
        totalKalori += Number(dataMakanan[i][3]) || 0;
        totalProtein += Number(dataMakanan[i][4]) || 0;
        totalKarbo += Number(dataMakanan[i][5]) || 0;
        totalLemak += Number(dataMakanan[i][6]) || 0;
      }
    }
  }

  var sheetAktivitas = ss.getSheetByName("Log_Aktivitas");
  if (sheetAktivitas && sheetAktivitas.getLastRow() > 1) {
    var dataAktivitas = sheetAktivitas.getRange(2, 1, sheetAktivitas.getLastRow() - 1, 5).getValues();
    for (var j = 0; j < dataAktivitas.length; j++) {
      if (formatDate(dataAktivitas[j][1]) === tglHariIni) {
        totalBakar += Number(dataAktivitas[j][4]) || 0;
      }
    }
  }

  var sheetAir = ss.getSheetByName("Log_Air");
  if (sheetAir && sheetAir.getLastRow() > 1) {
    var dataAir = sheetAir.getRange(2, 1, sheetAir.getLastRow() - 1, 3).getValues();
    for (var k = 0; k < dataAir.length; k++) {
      if (formatDate(dataAir[k][1]) === tglHariIni) {
        totalAir += Number(dataAir[k][2]) || 0;
      }
    }
  }

  return {
    kalori: Math.round(totalKalori),
    protein: Math.round(totalProtein * 10) / 10,
    karbo: Math.round(totalKarbo * 10) / 10,
    lemak: Math.round(totalLemak * 10) / 10,
    bakar: Math.round(totalBakar),
    air: Math.round(totalAir)
  };
}

function formatDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    var day = ("0" + val.getDate()).slice(-2);
    var month = ("0" + (val.getMonth() + 1)).slice(-2);
    var year = val.getFullYear();
    return day + "/" + month + "/" + year;
  }
  return val.toString().trim();
}
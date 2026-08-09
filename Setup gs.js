function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheets = ["Log_Makanan", "Log_Aktivitas", "Log_Air"];
  var headers = {
    "Log_Makanan": ["Email", "Tanggal", "Nama Makanan", "Kalori", "Protein", "Karbo", "Lemak"],
    "Log_Aktivitas": ["Email", "Tanggal", "Jenis Aktivitas", "Durasi", "Kalori Terbakar"],
    "Log_Air": ["Email", "Tanggal", "Jumlah Air (ml)"]
  };

  sheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers[name]);
    }
  });
  
  Logger.log("Setup Berhasil! Lembar kerja (sheet) telah disiapkan.");
}
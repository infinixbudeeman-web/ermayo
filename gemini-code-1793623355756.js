function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('ระบบสรุปยอดอุบัติเหตุ อำเภอมายอ')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Accidents");
  if (!sheet) {
    sheet = ss.insertSheet("Accidents");
    sheet.appendRow([
      "ID", "Timestamp", "Date", "Subdistrict", "Location", 
      "Gender", "AgeGroup", "CrashType", "VehicleType", 
      "MinorInjured", "SevereInjured", "Dead", "Reporter"
    ]);
  }
  return sheet;
}

function addAccident(data) {
  var sheet = getSheet();
  var id = "ACC-" + Date.now();
  var timestamp = new Date();
  sheet.appendRow([
    id, 
    timestamp, 
    data.date, 
    data.subdistrict, 
    data.location,
    data.gender, 
    data.ageGroup, 
    data.crashType, 
    data.vehicleType,
    data.minorInjured, 
    data.severeInjured, 
    data.dead, 
    data.reporter
  ]);
  return { success: true, message: "บันทึกข้อมูลสำเร็จ" };
}

function getAccidents() {
  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  
  var headers = rows[0];
  var data = [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
    data.push(obj);
  }
  return data;
}

function deleteAccident(id) {
  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] == id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "ลบข้อมูลสำเร็จ" };
    }
  }
  return { success: false, message: "ไม่พบข้อมูลที่ต้องการลบ" };
}
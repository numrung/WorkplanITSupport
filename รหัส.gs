function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('IT Sura Support')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 📌 ดึงข้อมูลเหตุการณ์จาก Google Sheets
function getEvents() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  var data = sheet.getDataRange().getValues();
  var events = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      events.push({
        id: i,
        title: data[i][0],
        start: data[i][1],
        end: data[i][2]
      });
    }
  }
  return events;
}

// 📌 ค้นหาบรรทัดของวันที่ในคอลัมน์ E
function findRowByDate(date) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  var data = sheet.getDataRange().getValues();
  var formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");

  for (var i = 1; i < data.length; i++) {
    var sheetDate = Utilities.formatDate(new Date(data[i][4]), Session.getScriptTimeZone(), "yyyy-MM-dd"); // คอลัมน์ E (index 4)
    if (formattedDate === sheetDate) {
      return i + 1; // เนื่องจาก index เริ่มที่ 0 แต่ Google Sheets ใช้แถวเริ่มที่ 1
    }
  }
  return -1; // ไม่พบวันที่
}

// 📌 เพิ่มเหตุการณ์ไปที่บรรทัดเดียวกับวันที่ในคอลัมน์ E
function addEvent(title, startDate, endDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  var row = findRowByDate(startDate);

  if (row === -1) {
    return "ไม่สามารถเพิ่มเหตุการณ์! วันที่ไม่มีอยู่ในระบบ";
  }

  sheet.getRange(row, 1).setValue(title);      // คอลัมน์ A = ชื่อเหตุการณ์
  sheet.getRange(row, 2).setValue(startDate);    // คอลัมน์ B = วันที่เริ่มต้น
  sheet.getRange(row, 3).setValue(endDate);        // คอลัมน์ C = วันที่สิ้นสุด

  return "เพิ่มเหตุการณ์สำเร็จ!";
}

// 📌 ล้างค่าของเหตุการณ์ในคอลัมน์ A และ B โดยใช้วันที่เริ่มต้นในการค้นหาแถว
function deleteEventByStartDate(startDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  var row = findRowByDate(startDate);

  if (row === -1) {
    return "ไม่พบวันที่ในระบบ";
  }

  // ล้างค่าในคอลัมน์ A และ B ของแถวนั้น (คอลัมน์อื่นคงไว้)
  sheet.getRange(row, 1, 1, 2).clearContent();
  return "ล้างค่าของเหตุการณ์สำเร็จ!";
}


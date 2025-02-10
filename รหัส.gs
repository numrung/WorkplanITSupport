function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('IT Sura Support')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getEvents() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var events = data.slice(1).map((row, i) => row[0] && row[1] ? {
    id: i + 1,
    title: row[0],
    start: formatDate(row[1]),
    end: row[2] ? formatDate(row[2]) : formatDate(row[1]) // ใช้วันที่เริ่มต้นหากไม่มีวันที่สิ้นสุด
  } : null).filter(Boolean);
  
  return events;
}

function formatDate(date) {
  return typeof date === "string" ? date : Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function addEvent(title, startDate, endDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  var data = sheet.getDataRange().getValues();
  
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    var sheetStartDate = formatDate(data[i][4]); // คอลัมน์ E (Index 4)
    if (sheetStartDate === startDate) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    // หากไม่พบแถวที่ตรงกับวันที่เริ่มต้น ให้เพิ่มแถวใหม่
    rowIndex = sheet.getLastRow() + 1;
  }
  
  // หากไม่มีวันที่สิ้นสุด ให้ใช้วันที่เริ่มต้นเป็นวันที่สิ้นสุด
  endDate = endDate || startDate;
  
  // บันทึกเหตุการณ์ลงในแถวที่พบหรือแถวใหม่
  sheet.getRange(rowIndex, 1).setValue(title); // คอลัมน์ A
  sheet.getRange(rowIndex, 2).setValue(startDate); // คอลัมน์ B
  sheet.getRange(rowIndex, 3).setValue(endDate); // คอลัมน์ C
  sheet.getRange(rowIndex, 5).setValue(startDate); // คอลัมน์ E (วันที่เริ่มต้น)

  return "เพิ่มเหตุการณ์เสร็จสิ้น!";
}

function getEventDetails(eventId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  var data = sheet.getDataRange().getValues();
  var eventDetails = {};

  for (var i = 1; i < data.length; i++) {
    if (i + 1 === eventId) {  // ใช้ index แทน ID จริง
      eventDetails.description = data[i][2] || "ไม่มีข้อมูล";  // ข้อมูลจากคอลัมน์ C
      break;
    }
  }

  return eventDetails;
}

function deleteEventByStartDate(startDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (formatDate(data[i][4]) === startDate) { // คอลัมน์ E
      // ลบข้อมูลในคอลัมน์ A, B, C
      sheet.getRange(i + 1, 1).setValue(""); // ลบชื่อเหตุการณ์ในคอลัมน์ A
      sheet.getRange(i + 1, 2).setValue(""); // ลบวันที่เริ่มต้นในคอลัมน์ B
      sheet.getRange(i + 1, 3).setValue(""); // ลบวันที่สิ้นสุดในคอลัมน์ C
      return "เหตุการณ์ถูกลบข้อมูลเรียบร้อย!";
    }
  }

  return "ไม่พบเหตุการณ์ที่ต้องการลบ!";
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

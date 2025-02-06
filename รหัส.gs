function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('IT Sura Support')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 📌 ดึงข้อมูลเหตุการณ์จาก Google Sheets (แก้ไขให้โหลดเร็วขึ้น)
function getEvents() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  if (!sheet) {
    Logger.log("❌ ไม่พบชีต Events");
    return [];
  }

  var data = sheet.getDataRange().getValues();
  var events = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1]) {  // ตรวจสอบว่ามีชื่อเหตุการณ์และวันที่
      events.push({
        id: i,
        title: data[i][0],  // คอลัมน์ A: ชื่อเหตุการณ์
        start: formatDate(data[i][1]),  // คอลัมน์ B: วันที่เริ่มต้น
        end: data[i][2] ? formatDate(data[i][2]) : formatDate(data[i][1]) // คอลัมน์ C: วันที่สิ้นสุด (ถ้าไม่มีให้ใช้ start)
      });
    }
  }

  Logger.log("✅ ข้อมูลเหตุการณ์ที่ดึงมา: " + JSON.stringify(events)); // ตรวจสอบข้อมูลที่ส่งไป
  return events;
}

// 📌 ฟังก์ชันช่วยจัดรูปแบบวันที่ให้เป็น YYYY-MM-DD
function formatDate(date) {
  if (typeof date === "string") return date; // ถ้าเป็น string อยู่แล้ว ให้คืนค่าเดิม
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

// 📌 ค้นหาบรรทัดของวันที่ในคอลัมน์ B (แก้ไขให้ทำงานเร็วขึ้น)
function findRowByDate(date) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  if (!sheet) {
    Logger.log("❌ ไม่พบชีต Events");
    return -1;
  }

  var data = sheet.getDataRange().getValues();
  var formattedDate = formatDate(date);

  for (var i = 1; i < data.length; i++) {
    var sheetDate = formatDate(data[i][1]); // คอลัมน์ B (index 1) = วันที่เริ่มต้น
    if (formattedDate === sheetDate) {
      Logger.log("✅ พบวันที่ในแถวที่: " + (i + 1));
      return i + 1; // เนื่องจาก Google Sheets เริ่มแถวที่ 1
    }
  }

  Logger.log("❌ ไม่พบวันที่ในระบบ");
  return -1; // ไม่พบวันที่
}

// 📌 เพิ่มเหตุการณ์ไปที่บรรทัดเดียวกับวันที่ในคอลัมน์ B
function addEvent(title, startDate, endDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  var row = findRowByDate(startDate);

  if (row === -1) {
    Logger.log("❌ ไม่สามารถเพิ่มเหตุการณ์! วันที่ไม่มีอยู่ในระบบ");
    return "ไม่สามารถเพิ่มเหตุการณ์! วันที่ไม่มีอยู่ในระบบ";
  }

  sheet.getRange(row, 1).setValue(title);  // คอลัมน์ A = ชื่อเหตุการณ์
  sheet.getRange(row, 2).setValue(startDate);  // คอลัมน์ B = วันที่เริ่มต้น
  sheet.getRange(row, 3).setValue(endDate || startDate); // คอลัมน์ C = วันที่สิ้นสุด (ถ้าไม่มีให้ใช้ start)

  Logger.log("✅ เพิ่มเหตุการณ์สำเร็จ: " + title);
  return "เพิ่มเหตุการณ์สำเร็จ!";
}

// 📌 ล้างค่าของเหตุการณ์โดยใช้วันที่เริ่มต้นในการค้นหาแถว
function deleteEventByStartDate(startDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  var row = findRowByDate(startDate);

  if (row === -1) {
    Logger.log("❌ ไม่พบวันที่ในระบบ");
    return "ไม่พบวันที่ในระบบ";
  }

  // ล้างค่าในคอลัมน์ A, B และ C ของแถวนั้น (คอลัมน์อื่นคงไว้)
  sheet.getRange(row, 1, 1, 3).clearContent();

  Logger.log("✅ ล้างค่าของเหตุการณ์สำเร็จในแถวที่: " + row);
  return "ล้างค่าของเหตุการณ์สำเร็จ!";
}

const SHEET_ID = '1vmuDt4MKQvY5SZyZwcmA5uukObAVj5VdGeyTXASKRf8';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ระบบติดตามเอกสารทางการศึกษา | ร.ร.วัดสุทธิวราราม')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getOrCreateSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
    }
  }
  return sheet;
}

function getRawData(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    return sheet ? sheet.getDataRange().getDisplayValues() : [];
  } catch (e) { return []; }
}

// 1. ตรวจสอบ Login
function checkLogin(user, pass) {
  const data = getRawData('USER');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(user).trim() && String(data[i][2]).trim() === String(pass).trim()) {
      let role = data[i][4] ? String(data[i][4]).trim().toUpperCase() : (String(user).toLowerCase() === 'admin' ? 'ADMIN' : 'STAFF');
      return { success: true, username: data[i][3], role: role };
    }
  }
  return { success: false, message: 'ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง' };
}

// 2. ดึงข่าวสาร
function getNews() {
  const data = getRawData('NEWS');
  if (data.length <= 1) return [];
  return data.slice(1).filter(row => String(row[3]).toUpperCase() === 'TRUE').map(row => row[2]).reverse();
}

// 3. ค้นหานักเรียนแบบ Live Search & ดึงยกห้อง
function searchStudentsLive(keyword, clsFilter = '', roomFilter = '') {
  const data = getRawData('STU_DATA');
  if (data.length <= 1) return { success: true, data: [] };
  
  let results = [];
  
  // กรณีเลือกดึงทั้งห้อง
  if (clsFilter && roomFilter) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(clsFilter).trim() && String(data[i][1]).trim() === String(roomFilter).trim()) {
        results.push({
          class: data[i][0], room: data[i][1], number: data[i][2],
          stdcode: String(data[i][3]).trim(), prefix: data[i][4], firstname: data[i][5], secondname: data[i][6],
          displayText: `${data[i][3]} ${data[i][4]}${data[i][5]} ${data[i][6]} - ม.${data[i][0]}/${data[i][1]} เลขที่ ${data[i][2]}`
        });
      }
    }
    return { success: true, data: results };
  }

  // ค้นหาตามคำค้นหา
  if (!keyword || keyword.trim().length < 2) return { success: true, data: [] };
  let kw = keyword.trim().toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    let stdCode = String(data[i][3]).trim();
    let fname = String(data[i][5]).trim();
    let lname = String(data[i][6]).trim();
    let fullName = (fname + ' ' + lname).toLowerCase();
    
    if (stdCode.includes(kw) || fullName.includes(kw)) {
      results.push({
        class: data[i][0], room: data[i][1], number: data[i][2],
        stdcode: stdCode, prefix: data[i][4], firstname: fname, secondname: lname,
        displayText: `${stdCode} ${data[i][4]}${fname} ${lname} - ม.${data[i][0]}/${data[i][1]} เลขที่ ${data[i][2]}`
      });
      if (results.length >= 20) break;
    }
  }
  return { success: true, data: results };
}

// 4. ตั้งค่าประเภทเอกสาร
function getDocTypes() {
  const props = PropertiesService.getScriptProperties();
  let types = props.getProperty('DOC_TYPES');
  if (!types) {
    const defaults = [
      "ปพ.1 - บ (ระเบียนแสดงผลการเรียน) ระดับชั้น ม.ต้น", 
      "ปพ.1 - พ (ระเบียนแสดงผลการเรียน) ระดับชั้น ม.ปลาย", 
      "ปพ.7 ใบรับรองการเป็นนักเรียน"
    ];
    props.setProperty('DOC_TYPES', JSON.stringify(defaults));
    return defaults;
  }
  return JSON.parse(types);
}

function saveDocType(newType) {
  const props = PropertiesService.getScriptProperties();
  let types = JSON.parse(props.getProperty('DOC_TYPES') || '[]');
  if (!types.includes(newType)) {
    types.push(newType);
    props.setProperty('DOC_TYPES', JSON.stringify(types));
  }
  return getDocTypes();
}

function deleteDocType(typeToRemove) {
  const props = PropertiesService.getScriptProperties();
  let types = JSON.parse(props.getProperty('DOC_TYPES') || '[]');
  types = types.filter(t => t !== typeToRemove);
  props.setProperty('DOC_TYPES', JSON.stringify(types));
  return getDocTypes();
}

// 5. บันทึกคำร้องแบบกลุ่ม
function submitBatchRequests(studentsList, reqDetails) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = getOrCreateSheet(ss, 'REQUESTS', ['ReqID', 'Timestamp', 'StdCode', 'Prefix', 'FName', 'LName', 'DocType', 'Semester', 'Qty', 'Status', 'Admin', 'Note']);
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    
    let createdIds = [];
    studentsList.forEach((std, idx) => {
      let reqId = 'REQ' + Utilities.formatDate(new Date(new Date().getTime() + idx * 10), "GMT+7", "yyMMddHHmmssSSS");
      sheet.appendRow([
        reqId, timestamp, std.stdcode, std.prefix, std.firstname, std.secondname, 
        reqDetails.docType, reqDetails.semesters || '-', reqDetails.qty, reqDetails.status, reqDetails.adminName, reqDetails.note || ''
      ]);
      logStatusHistory(reqId, reqDetails.status, reqDetails.adminName, reqDetails.note || '');
      createdIds.push(reqId);
    });
    
    return { success: true, message: `บันทึกคำร้องเรียบร้อยแล้วทั้งหมด ${studentsList.length} รายการ`, createdIds: createdIds };
  } catch (e) { return { success: false, message: e.message }; }
}

// 6. ดึงคำร้อง
function getRequests(filterStdCode = null) {
  const data = getRawData('REQUESTS');
  let result = [];
  for (let i = 1; i < data.length; i++) {
    if (filterStdCode && String(data[i][2]).trim() !== String(filterStdCode).trim() && String(data[i][0]).trim() !== String(filterStdCode).trim()) continue;
    result.push({
      reqid: data[i][0], timestamp: data[i][1], stdcode: data[i][2], prefix: data[i][3], fname: data[i][4], lname: data[i][5],
      doctype: data[i][6], semester: data[i][7], qty: data[i][8], status: data[i][9], admin: data[i][10], note: data[i][11] || ''
    });
  }
  return { success: true, data: result.reverse() };
}

// 7. สถิติ Dashboard
function getDashboardStats(startDateStr = null, endDateStr = null) {
  let reqs = getRequests().data;
  
  if (startDateStr && endDateStr) {
    let start = new Date(startDateStr);
    let end = new Date(endDateStr);
    end.setHours(23, 59, 59, 999);
    
    reqs = reqs.filter(r => {
      let parts = r.timestamp.split(' ')[0].split('/');
      let rDate = new Date(parts[2], parts[1] - 1, parts[0]);
      return rDate >= start && rDate <= end;
    });
  }

  let statusCounts = { 'ยื่นคำร้อง': 0, 'กำลังดำเนินการ': 0, 'ติดปัญหา': 0, 'เสร็จสมบูรณ์': 0 };
  let typeCounts = {};

  reqs.forEach(r => {
    if (r.status === 'ยื่นคำร้อง') statusCounts['ยื่นคำร้อง']++;
    else if (r.status.includes('อยู่ระหว่างดำเนินการ') || r.status.includes('รอรูป เพื่อทำการลงนาม')) statusCounts['กำลังดำเนินการ']++;
    else if (r.status.includes('ขัดข้อง')) statusCounts['ติดปัญหา']++;
    else if (r.status.includes('เสร็จสมบูรณ์')) statusCounts['เสร็จสมบูรณ์']++;

    let typeShort = r.doctype.split(' ')[0] + (r.doctype.includes('ม.ต้น') ? ' (ต้น)' : r.doctype.includes('ม.ปลาย') ? ' (ปลาย)' : '');
    typeCounts[typeShort] = (typeCounts[typeShort] || 0) + 1;
  });

  return {
    total: reqs.length,
    pending: statusCounts['ยื่นคำร้อง'],
    processing: statusCounts['กำลังดำเนินการ'],
    completed: statusCounts['เสร็จสมบูรณ์'],
    issue: statusCounts['ติดปัญหา'],
    chartStatus: { labels: Object.keys(statusCounts), data: Object.values(statusCounts) },
    chartTypes: { labels: Object.keys(typeCounts), data: Object.values(typeCounts) }
  };
}

// 8. Audit Log History
function logStatusHistory(reqId, newStatus, adminName, note = '') {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const logSheet = getOrCreateSheet(ss, 'Log', ['ReqID', 'Timestamp', 'Status', 'Admin', 'Note']);
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    logSheet.appendRow([reqId, timestamp, newStatus, adminName, note]);
  } catch(e) {}
}

function getHistory(reqId) {
  const data = getRawData('Log');
  let history = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(reqId).trim()) {
      history.push({ timestamp: data[i][1], status: data[i][2], admin: data[i][3], note: data[i][4] || '' });
    }
  }
  return { success: true, data: history.reverse() };
}

// 9. อัปเดตสถานะเดี่ยว
function updateStatus(reqId, newStatus, adminName, note = '') {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('REQUESTS');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(reqId)) {
        sheet.getRange(i + 1, 10).setValue(newStatus); 
        sheet.getRange(i + 1, 11).setValue(adminName);
        sheet.getRange(i + 1, 12).setValue(note);
        logStatusHistory(reqId, newStatus, adminName, note);
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบรหัสคำร้องนี้' };
  } catch (e) { return { success: false, message: e.message }; }
}

// 10. อัปเดตสถานะกลุ่ม (Bulk Update)
function bulkUpdateStatus(reqIds, newStatus, adminName, note = '') {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('REQUESTS');
    const data = sheet.getDataRange().getValues();
    let updatedCount = 0;

    for (let i = 1; i < data.length; i++) {
      let currentId = String(data[i][0]).trim();
      if (reqIds.includes(currentId)) {
        sheet.getRange(i + 1, 10).setValue(newStatus);
        sheet.getRange(i + 1, 11).setValue(adminName);
        if (note) sheet.getRange(i + 1, 12).setValue(note);
        logStatusHistory(currentId, newStatus, adminName, note);
        updatedCount++;
      }
    }
    return { success: true, count: updatedCount };
  } catch (e) { return { success: false, message: e.message }; }
}

// 11. Auto-Archiving
function archiveOldData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const mainSheet = ss.getSheetByName('REQUESTS');
    const archiveSheet = getOrCreateSheet(ss, 'Archive', ['ReqID', 'Timestamp', 'StdCode', 'Prefix', 'FName', 'LName', 'DocType', 'Semester', 'Qty', 'Status', 'Admin', 'Note']);
    
    const data = mainSheet.getDataRange().getValues();
    const now = new Date();
    const oneYearMs = 365 * 24 * 60 * 60 * 1000;
    let archivedCount = 0;

    for (let i = data.length - 1; i >= 1; i--) {
      let status = String(data[i][9]);
      let dateStr = String(data[i][1]);
      
      if (status.includes('เสร็จสมบูรณ์')) {
        let parts = dateStr.split(' ')[0].split('/');
        let reqDate = new Date(parts[2], parts[1] - 1, parts[0]);
        
        if ((now - reqDate) > oneYearMs) {
          archiveSheet.appendRow(data[i]);
          mainSheet.deleteRow(i + 1);
          archivedCount++;
        }
      }
    }
    return { success: true, count: archivedCount };
  } catch (e) { return { success: false, message: e.message }; }
}

/**
 * CPE File Upload — Google Apps Script Web App
 *
 * วิธี deploy:
 * 1. ไปที่ https://script.google.com → New project
 * 2. วางโค้ดทั้งหมดนี้ลงใน Code.gs
 * 3. Deploy → New deployment → Web app
 *      Execute as : Me
 *      Who has access : Anyone
 * 4. คัดลอก Web app URL มาใส่ใน APPS_SCRIPT_URL ใน cpe-form-system.html
 */

var ROOT_FOLDER  = 'CPE เอกสารต่ออายุ';
var SHEET_NAME   = 'ข้อมูลโครงการบริการวิชาการและ CPE';   // ชื่อ Google Sheet

/* ─── Main POST handler ─── */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // action: 'log' → บันทึกข้อมูลลง Sheet (ไม่อัปโหลดไฟล์)
    if (data.action === 'log') {
      return logToSheet(data);
    }

    // action: 'upload' (default) → อัปโหลดไฟล์ไปยัง Drive
    var b64     = data.fileBase64;
    var name    = data.fileName  || 'file';
    var mime    = data.mimeType  || 'application/octet-stream';
    var orgName = data.orgName   || 'ไม่ระบุหน่วยงาน';
    var docType = data.docType   || 'document';

    var root      = getOrCreate(ROOT_FOLDER, null);
    var orgFolder = getOrCreate(orgName, root);

    var bytes = Utilities.base64Decode(b64);
    var blob  = Utilities.newBlob(bytes, mime, name);
    var file  = orgFolder.createFile(blob);

    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return respond({ success:true, url:file.getUrl(), id:file.getId(), name:file.getName() });
  } catch(err) {
    return respond({ success:false, error:err.toString() });
  }
}

function doGet() {
  return respond({ status:'CPE Upload API running' });
}

/* ─── Log submission to Google Sheet ─── */
function logToSheet(data) {
  try {
    var ss = getOrCreateSheet();
    var sheet;

    if (data.type === 'renewal') {
      sheet = getOrCreateTab(ss, 'ต่ออายุ CPE',
        ['วันที่-เวลา','หน่วยงาน','หนังสือขอต่ออายุ (URL)','แบบ สภ.59 (URL)','หลักฐานชำระ (URL)','สถานะ']);
      sheet.appendRow([
        formatDate(data.submittedAt),
        data.orgName   || '',
        data.letterUrl || '',
        data.form59Url || '',
        data.paymentUrl|| '',
        data.status    || 'รอตรวจสอบ'
      ]);
    } else if (data.type === 'conference') {
      sheet = getOrCreateTab(ss, 'ประชุมวิชาการ',
        ['วันที่-เวลา','หน่วยงาน','ชื่อประชุม','วันจัด','ไฟล์ PDF','ไฟล์ Word','ผู้บรรยาย (ชื่อ | CV URL)']);
      var speakers = (data.speakers || []).map(function(s){
        return s.name + (s.cvUrl ? ' | '+s.cvUrl : '');
      }).join('\n');
      sheet.appendRow([
        formatDate(data.submittedAt),
        data.orgName    || '',
        data.confName   || '',
        data.confDate   || '',
        data.pdfUrl     || '',
        data.wordUrl    || '',
        speakers
      ]);
    } else if (data.type === 'academic') {
      sheet = getOrCreateTab(ss, 'โครงการบริการวิชาการ',
        ['วันที่บันทึก','ผู้บันทึก','ชื่อโครงการ','หน่วยงาน','ผู้รับผิดชอบ','ผู้ประสานงาน',
         'สถานที่','วันเริ่ม','วันสิ้นสุด','งบประมาณรวม (บาท)','แหล่งงบ',
         'รายรับรวม (บาท)','รายรับสุทธิ',
         'SDGs','ยุทธศาสตร์ มบ.','ยุทธศาสตร์ ม.','EdPEx',
         'วัตถุประสงค์','กลุ่มเป้าหมาย',
         'ผลที่คาดว่าจะได้รับ','KPI เชิงปริมาณ','KPI เชิงคุณภาพ','ปีแผน']);
      sheet.appendRow([
        formatDate(data.savedAt),
        data.ownerEmail   || '',
        data.projectName  || '',
        data.dept         || '',
        data.sig1Name     || '',
        data.coordinators || '',
        data.location     || '',
        data.startDate    || '',
        data.endDate      || '',
        data.budgetTotal  || '',
        data.budgetSource || '',
        data.incomeTotal  || '',
        data.incomeNetText|| '',
        data.sdgs         || '',
        data.buuStrategies|| '',
        data.mStrategies  || '',
        data.edpex        || '',
        data.objectives   || '',
        data.targets      || '',
        data.results      || '',
        data.kpiQty       || '',
        data.kpiQual      || '',
        data.planYear     || ''
      ]);
    }

    // ── แจ้งเตือน admin ทางอีเมล ──
    sendAdminEmail(data);

    return respond({ success:true });
  } catch(err) {
    return respond({ success:false, error:err.toString() });
  }
}

/* ─── Email notification ─── */
var ADMIN_EMAIL = 'thiraphong.ge@go.buu.ac.th';

function sendAdminEmail(data) {
  try {
    var formType, subject, body;
    var now = formatDate(data.submittedAt || data.savedAt);

    if (data.type === 'renewal') {
      formType = 'ต่ออายุ CPE';
      subject  = '[แจ้งเตือน] มีการส่งเอกสารต่ออายุ CPE — ' + (data.orgName || '');
      body     = 'มีการส่งเอกสารใหม่เข้ามาในระบบ\n\n'
               + 'ประเภทฟอร์ม : ' + formType + '\n'
               + 'หน่วยงาน   : ' + (data.orgName || '-') + '\n'
               + 'วันที่-เวลา : ' + now + '\n'
               + 'สถานะ      : ' + (data.status || 'รอตรวจสอบ') + '\n\n'
               + 'กรุณาตรวจสอบในระบบหรือ Google Sheet: "ข้อมูลโครงการบริการวิชาการและ CPE" → tab ต่ออายุ CPE';

    } else if (data.type === 'conference') {
      formType = 'ประชุมวิชาการ';
      subject  = '[แจ้งเตือน] มีการส่งข้อมูลประชุมวิชาการ — ' + (data.confName || '');
      body     = 'มีการส่งข้อมูลประชุมวิชาการใหม่เข้ามาในระบบ\n\n'
               + 'ประเภทฟอร์ม : ' + formType + '\n'
               + 'หน่วยงาน   : ' + (data.orgName || '-') + '\n'
               + 'ชื่อประชุม  : ' + (data.confName || '-') + '\n'
               + 'วันที่-เวลา : ' + now + '\n\n'
               + 'กรุณาตรวจสอบในระบบหรือ Google Sheet: "ข้อมูลโครงการบริการวิชาการและ CPE" → tab ประชุมวิชาการ';

    } else if (data.type === 'academic') {
      formType = 'แบบฟอร์มโครงการบริการวิชาการ';
      subject  = '[แจ้งเตือน] มีการบันทึกโครงการบริการวิชาการ — ' + (data.projectName || '');
      body     = 'มีการบันทึกโครงการบริการวิชาการใหม่เข้ามาในระบบ\n\n'
               + 'ประเภทฟอร์ม   : ' + formType + '\n'
               + 'ชื่อโครงการ   : ' + (data.projectName || '-') + '\n'
               + 'หน่วยงาน     : ' + (data.dept || '-') + '\n'
               + 'ผู้รับผิดชอบ  : ' + (data.sig1Name || '-') + '\n'
               + 'ผู้บันทึก     : ' + (data.ownerEmail || '-') + '\n'
               + 'วันที่-เวลา   : ' + now + '\n\n'
               + 'กรุณาตรวจสอบในระบบหรือ Google Sheet: "ข้อมูลโครงการบริการวิชาการและ CPE" → tab โครงการบริการวิชาการ';
    } else {
      return; // ไม่ส่งถ้าประเภทไม่รู้จัก
    }

    MailApp.sendEmail(ADMIN_EMAIL, subject, body);
  } catch(e) {
    // ไม่ให้ email error กระทบ main flow
  }
}

/* ─── Sheet helpers ─── */
function getOrCreateSheet() {
  var files = DriveApp.getFilesByName(SHEET_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  var ss = SpreadsheetApp.create(SHEET_NAME);
  return ss;
}

function getOrCreateTab(ss, tabName, headers) {
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d0e4f7');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function formatDate(iso) {
  if (!iso) return new Date().toLocaleString('th-TH',{timeZone:'Asia/Bangkok'});
  return new Date(iso).toLocaleString('th-TH',{timeZone:'Asia/Bangkok'});
}

/* ─── Drive helpers ─── */
function getOrCreate(name, parent) {
  var it = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent ? parent.createFolder(name) : DriveApp.createFolder(name);
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

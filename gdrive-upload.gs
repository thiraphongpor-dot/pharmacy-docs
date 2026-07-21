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

    // action: 'send_expert_email' → ส่งอีเมลให้ผู้ทรงคุณวุฒิ
    if (data.action === 'send_expert_email') {
      return sendExpertEmail(data);
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
      var confHdrs = ['วันที่-เวลา','หน่วยงาน','ชื่อประชุม','วันจัด','รายละเอียดโครงการ','CV วิทยากร','กำหนดการ'];
      sheet = getOrCreateTab(ss, 'ประชุมวิชาการ', confHdrs);
      // อัพเดต header ทุกครั้งเพื่อแก้ column เก่า (ไฟล์ PDF, ไฟล์ Word, ผู้บรรยาย)
      for (var ch = 0; ch < confHdrs.length; ch++) {
        sheet.getRange(1, ch + 1).setValue(confHdrs[ch]).setFontWeight('bold').setBackground('#d0e4f7');
      }
      // ลบ column ที่เกิน (ถ้ามี)
      if (sheet.getLastColumn() > confHdrs.length) {
        sheet.deleteColumns(confHdrs.length + 1, sheet.getLastColumn() - confHdrs.length);
      }
      sheet.appendRow([
        formatDate(data.submittedAt),
        data.orgName      || '',
        data.confName     || '',
        data.confDate     || '',
        data.projectUrl   || '',
        data.cvfileUrl    || '',
        data.scheduleUrl  || ''
      ]);
    } else if (data.type === 'org_register') {
      sheet = getOrCreateTab(ss, 'ลงทะเบียนหน่วยงาน',
        ['วันที่ยื่น','ผู้ยื่น','ชื่อหน่วยงาน','ที่อยู่','โทรศัพท์','โทรสาร','อีเมล','เว็บไซต์',
         'หัวหน้าหน่วยกิต','ตำแหน่ง','สถานะ','รหัสหน่วยงาน','วันที่ผ่านการรับรอง','วันหมดอายุ','วันที่อนุมัติ','ผู้อนุมัติ']);
      sheet.appendRow([
        formatDate(data.submittedAt),
        data.submittedBy  || '',
        data.orgName      || '',
        data.address      || '',
        data.phone        || '',
        data.fax          || '',
        data.email        || '',
        data.website      || '',
        data.headName     || '',
        data.headPosition || '',
        data.status       || 'รอตรวจสอบ',
        '', '', '', '', ''
      ]);

    } else if (data.type === 'org_register_approved') {
      // ค้นหาแถวของหน่วยงานใน Sheet แล้วอัพเดตข้อมูลการอนุมัติ
      sheet = getOrCreateTab(ss, 'ลงทะเบียนหน่วยงาน',
        ['วันที่ยื่น','ผู้ยื่น','ชื่อหน่วยงาน','ที่อยู่','โทรศัพท์','โทรสาร','อีเมล','เว็บไซต์',
         'หัวหน้าหน่วยกิต','ตำแหน่ง','สถานะ','รหัสหน่วยงาน','วันที่ผ่านการรับรอง','วันหมดอายุ','วันที่อนุมัติ','ผู้อนุมัติ']);
      // เติม header คอลัมน์ที่อาจยังไม่มี (กรณี Sheet เก่า)
      var hdrRange = sheet.getRange(1, 1, 1, 16).getValues()[0];
      var extraHdrs = ['รหัสหน่วยงาน','วันที่ผ่านการรับรอง','วันหมดอายุ','วันที่อนุมัติ','ผู้อนุมัติ'];
      for (var h = 0; h < extraHdrs.length; h++) {
        if (!hdrRange[11 + h]) {
          sheet.getRange(1, 12 + h).setValue(extraHdrs[h]).setFontWeight('bold').setBackground('#d0e4f7');
        }
      }
      // ค้นหาแถวที่ตรงกับชื่อหน่วยงาน แล้วอัพเดต
      var lastRow = sheet.getLastRow();
      var found   = false;
      for (var i = 2; i <= lastRow; i++) {
        if (sheet.getRange(i, 3).getValue() === (data.orgName || '')) {
          sheet.getRange(i, 11).setValue('อนุมัติแล้ว');
          sheet.getRange(i, 12).setValue(data.orgCode        || '');
          sheet.getRange(i, 13).setValue(data.registrationDate|| '');
          sheet.getRange(i, 14).setValue(data.expiryDate     || '');
          sheet.getRange(i, 15).setValue(formatDate(data.approvedAt));
          sheet.getRange(i, 16).setValue(data.approvedBy     || '');
          found = true;
          break;
        }
      }
      if (!found) {
        // ถ้าหาไม่เจอ (เช่น submit ก่อน deploy ใหม่) ให้ append แถวใหม่แทน
        sheet.appendRow([
          formatDate(data.approvedAt), data.submittedBy||'', data.orgName||'',
          '','','','','','','','อนุมัติแล้ว',
          data.orgCode||'', data.registrationDate||'', data.expiryDate||'',
          formatDate(data.approvedAt), data.approvedBy||''
        ]);
      }

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
    } else if (data.type === 'org_register') {
      formType = 'ลงทะเบียนหน่วยงาน CPE';
      subject  = '[แจ้งเตือน] มีคำขอลงทะเบียนหน่วยงานใหม่ — ' + (data.orgName || '');
      body     = 'มีคำขอลงทะเบียนหน่วยงานใหม่เข้ามาในระบบ\n\n'
               + 'ประเภท         : ' + formType + '\n'
               + 'ชื่อหน่วยงาน   : ' + (data.orgName || '-') + '\n'
               + 'ที่อยู่         : ' + (data.address || '-') + '\n'
               + 'โทรศัพท์       : ' + (data.phone || '-') + '\n'
               + 'อีเมล          : ' + (data.email || '-') + '\n'
               + 'เว็บไซต์       : ' + (data.website || '-') + '\n'
               + 'ผู้ยื่นคำขอ    : ' + (data.submittedBy || '-') + '\n'
               + 'วันที่-เวลา    : ' + now + '\n'
               + 'สถานะ          : ' + (data.status || 'รอตรวจสอบ') + '\n\n'
               + 'กรุณาเข้าระบบ CPE เพื่ออนุมัติหรือปฏิเสธคำขอในส่วน "เพิ่ม/จัดการหน่วยงาน (Admin)"';
    } else if (data.type === 'org_register_approved') {
      return; // admin เป็นคนทำเอง ไม่ต้องส่งอีเมลซ้ำ
    } else {
      return; // ไม่ส่งถ้าประเภทไม่รู้จัก
    }

    MailApp.sendEmail(ADMIN_EMAIL, subject, body);
  } catch(e) {
    // ไม่ให้ email error กระทบ main flow
  }
}

/* ─── Send expert email ─── */
function sendExpertEmail(data) {
  var expertEmails = data.expertEmails || [];
  if (expertEmails.length === 0) {
    return respond({ success: false, error: 'ไม่มีอีเมลผู้ทรงคุณวุฒิ' });
  }

  var dateRange = data.dateStart || '-';
  if (data.dateEnd && data.dateEnd !== data.dateStart) dateRange += ' – ' + data.dateEnd;

  var subject = '[CPE] ขอเรียนเชิญพิจารณาหน่วยกิต CPE: ' + (data.confName || 'ประชุมวิชาการ');

  var html = '<div style="font-family:\'Sarabun\',sans-serif;max-width:600px;margin:0 auto;color:#1a202c">';
  html += '<div style="background:linear-gradient(135deg,#4f46e5,#7c3aed);padding:24px 28px;border-radius:12px 12px 0 0">';
  html += '<div style="color:#fff;font-size:1.2rem;font-weight:700">🎓 คำขอพิจารณาหน่วยกิต CPE</div>';
  html += '<div style="color:rgba(255,255,255,.8);font-size:.85rem;margin-top:4px">คณะเภสัชศาสตร์ มหาวิทยาลัยบูรพา</div>';
  html += '</div>';
  html += '<div style="background:#fff;border:1px solid #e5e7eb;border-top:none;padding:24px 28px;border-radius:0 0 12px 12px">';
  html += '<p>เรียน ผู้ทรงคุณวุฒิ</p>';
  html += '<p>คณะเภสัชศาสตร์ มหาวิทยาลัยบูรพา ขอเรียนเชิญพิจารณาหน่วยกิต CPE สำหรับงานประชุมวิชาการดังต่อไปนี้</p>';

  html += '<table style="width:100%;border-collapse:collapse;margin:16px 0;background:#f8fafc;border-radius:8px;overflow:hidden">';
  html += '<tr><td style="padding:9px 14px;font-size:.82rem;color:#6b7280;width:130px;border-bottom:1px solid #e5e7eb">ชื่องานประชุม</td><td style="padding:9px 14px;font-size:.88rem;font-weight:700;border-bottom:1px solid #e5e7eb">' + (data.confName || '-') + '</td></tr>';
  html += '<tr><td style="padding:9px 14px;font-size:.82rem;color:#6b7280;border-bottom:1px solid #e5e7eb">หน่วยงาน</td><td style="padding:9px 14px;font-size:.88rem;border-bottom:1px solid #e5e7eb">' + (data.orgName || '-') + '</td></tr>';
  html += '<tr><td style="padding:9px 14px;font-size:.82rem;color:#6b7280;border-bottom:1px solid #e5e7eb">วันที่จัด</td><td style="padding:9px 14px;font-size:.88rem;border-bottom:1px solid #e5e7eb">' + dateRange + '</td></tr>';
  if (data.location) html += '<tr><td style="padding:9px 14px;font-size:.82rem;color:#6b7280;border-bottom:1px solid #e5e7eb">สถานที่</td><td style="padding:9px 14px;font-size:.88rem;border-bottom:1px solid #e5e7eb">' + data.location + '</td></tr>';
  html += '</table>';

  if (data.adminMessage) {
    html += '<div style="background:#eff6ff;border-left:4px solid #4f46e5;border-radius:0 8px 8px 0;padding:12px 16px;margin:16px 0">';
    html += '<div style="font-size:.78rem;color:#4338ca;font-weight:700;margin-bottom:4px">📝 ข้อความจากเจ้าหน้าที่</div>';
    html += '<div style="font-size:.85rem;color:#1e3a8a;white-space:pre-line">' + data.adminMessage + '</div>';
    html += '</div>';
  }

  html += '<div style="margin:20px 0"><div style="font-size:.82rem;font-weight:700;color:#374151;margin-bottom:10px">📎 เอกสารแนบ</div>';
  html += '<div style="display:flex;flex-direction:column;gap:8px">';
  if (data.projectUrl) html += '<a href="' + data.projectUrl + '" style="color:#4f46e5;text-decoration:none;font-size:.85rem">📋 รายละเอียดโครงการ</a>';
  if (data.cvfileUrl)  html += '<a href="' + data.cvfileUrl  + '" style="color:#4f46e5;text-decoration:none;font-size:.85rem">👤 CV วิทยากร</a>';
  if (data.scheduleUrl) html += '<a href="' + data.scheduleUrl + '" style="color:#4f46e5;text-decoration:none;font-size:.85rem">📅 กำหนดการ</a>';
  html += '</div></div>';

  html += '<hr style="border:none;border-top:1px solid #e5e7eb;margin:20px 0">';
  html += '<p style="font-size:.78rem;color:#9ca3af">อีเมลนี้ส่งโดยระบบบริหารหน่วยกิต CPE คณะเภสัชศาสตร์ มหาวิทยาลัยบูรพา<br>กรุณาอย่าตอบกลับอีเมลนี้โดยตรง</p>';
  html += '</div></div>';

  var sent = 0;
  var errors = [];
  expertEmails.forEach(function(email) {
    email = (email || '').trim();
    if (email && email.indexOf('@') > 0) {
      try {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: html,
          name: 'ระบบ CPE คณะเภสัชศาสตร์ มหาวิทยาลัยบูรพา'
        });
        sent++;
      } catch(mailErr) {
        errors.push(email + ': ' + mailErr.message);
      }
    }
  });

  if (errors.length > 0) {
    return respond({ success: false, sent: sent, error: errors.join('; ') });
  }
  return respond({ success: true, sent: sent });
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

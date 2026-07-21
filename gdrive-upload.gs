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

    // action: 'update_conference_row' → admin อัพเดต row ใน Sheet ด้วย confId
    if (data.action === 'update_conference_row') {
      return updateConferenceRow(data);
    }

    // action: 'notify_cpe_result' → admin ยืนยันผลการพิจารณา → ส่งอีเมลแจ้ง user + อัพเดต Sheet
    if (data.action === 'notify_cpe_result') {
      return notifyCpeResult(data);
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
      var confHdrs = [
        'วันที่-เวลาที่ส่ง','ผู้ส่ง','หน่วยงาน','ชื่อประชุม',
        'วันที่เริ่มงาน','วันที่สิ้นสุดงาน','สถานที่','หมายเหตุ',
        'รายละเอียดโครงการ','CV วิทยากร','กำหนดการ',
        'มติ CPE (PDF)','มติ CPE (Word)',
        'จำนวนผู้ทรงคุณวุฒิ','อีเมลผู้ทรงคุณวุฒิ','วันที่ส่งผู้ทรง',
        'สถานะ','รหัสประชุม','ผลการพิจารณา (Word)','วันที่ยืนยันผล'
      ];
      sheet = getOrCreateTab(ss, 'ประชุมวิชาการ', confHdrs);
      // อัพเดต header ทุกครั้ง
      for (var ch = 0; ch < confHdrs.length; ch++) {
        sheet.getRange(1, ch + 1).setValue(confHdrs[ch]).setFontWeight('bold').setBackground('#d0e4f7');
      }
      // ลบ column ที่เกิน (ถ้ามี)
      if (sheet.getLastColumn() > confHdrs.length) {
        sheet.deleteColumns(confHdrs.length + 1, sheet.getLastColumn() - confHdrs.length);
      }
      sheet.appendRow([
        formatDate(data.submittedAt),   // 1 วันที่-เวลา
        data.submittedBy  || '',        // 2 ผู้ส่ง
        data.orgName      || '',        // 3 หน่วยงาน
        data.confName     || '',        // 4 ชื่อประชุม
        data.dateStart    || '',        // 5 วันที่เริ่มงาน
        data.dateEnd      || '',        // 6 วันที่สิ้นสุดงาน
        data.location     || '',        // 7 สถานที่
        data.note         || '',        // 8 หมายเหตุ
        data.projectUrl   || '',        // 9 รายละเอียดโครงการ
        data.cvfileUrl    || '',        // 10 CV วิทยากร
        data.scheduleUrl  || '',        // 11 กำหนดการ
        '',                             // 12 มติ CPE PDF
        '',                             // 13 มติ CPE Word
        '',                             // 14 จำนวนผู้ทรงคุณวุฒิ
        '',                             // 15 อีเมลผู้ทรงคุณวุฒิ
        '',                             // 16 วันที่ส่งผู้ทรง
        'รอพิจารณา',                   // 17 สถานะ
        data.confId       || '',        // 18 รหัสประชุม
        '',                             // 19 ผลการพิจารณา (Word)
        ''                              // 20 วันที่ยืนยันผล
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

/* ─── Update conference row by confId ─── */
// คอลัมน์ใน tab ประชุมวิชาการ (18 คอลัมน์):
//  1 วันที่-เวลาที่ส่ง  2 ผู้ส่ง  3 หน่วยงาน  4 ชื่อประชุม
//  5 วันที่เริ่มงาน  6 วันที่สิ้นสุดงาน  7 สถานที่  8 หมายเหตุ
//  9 รายละเอียดโครงการ  10 CV วิทยากร  11 กำหนดการ
//  12 มติ CPE (PDF)  13 มติ CPE (Word)
//  14 จำนวนผู้ทรงคุณวุฒิ  15 อีเมลผู้ทรงคุณวุฒิ  16 วันที่ส่งผู้ทรง
//  17 สถานะ  18 รหัสประชุม
function updateConferenceRow(data) {
  var ss    = getOrCreateSheet();
  var sheet = ss.getSheetByName('ประชุมวิชาการ');
  if (!sheet) return respond({ success: false, error: 'ไม่พบ tab ประชุมวิชาการ' });

  var confId = data.confId || '';
  if (!confId) return respond({ success: false, error: 'ไม่มี confId' });

  // ค้นหา row ที่มี confId ใน column 18
  var lastRow = sheet.getLastRow();
  var found   = -1;
  for (var r = 2; r <= lastRow; r++) {
    if (sheet.getRange(r, 18).getValue() === confId) { found = r; break; }
  }
  if (found < 0) return respond({ success: false, error: 'ไม่พบ row confId: ' + confId });

  // อัพเดตเฉพาะ field ที่ส่งมา
  if (data.matiPdfUrl    !== undefined) sheet.getRange(found, 12).setValue(data.matiPdfUrl    || '');
  if (data.matiWordUrl   !== undefined) sheet.getRange(found, 13).setValue(data.matiWordUrl   || '');
  if (data.expertEmails  !== undefined) {
    var emails = data.expertEmails || [];
    sheet.getRange(found, 14).setValue(emails.length);
    sheet.getRange(found, 15).setValue(emails.join(', '));
  }
  if (data.sentAt        !== undefined) sheet.getRange(found, 16).setValue(data.sentAt        ? formatDate(data.sentAt) : '');
  if (data.status        !== undefined) sheet.getRange(found, 17).setValue(data.status        || '');
  if (data.resultWordUrl !== undefined) sheet.getRange(found, 19).setValue(data.resultWordUrl || '');
  if (data.confirmedAt   !== undefined) sheet.getRange(found, 20).setValue(data.confirmedAt   ? formatDate(data.confirmedAt) : '');

  return respond({ success: true });
}

/* ─── Notify user: CPE result confirmed ─── */
function notifyCpeResult(data) {
  var userEmail = data.userEmail || '';
  if (!userEmail || userEmail.indexOf('@') < 0) {
    return respond({ success: false, error: 'ไม่มีอีเมล user' });
  }

  var confName  = data.confName  || 'ประชุมวิชาการ';
  var orgName   = data.orgName   || '';
  var dateRange = data.dateStart || '-';
  if (data.dateEnd && data.dateEnd !== data.dateStart) dateRange += ' – ' + data.dateEnd;

  var subject = '[CPE] ผลการพิจารณารับรองหน่วยกิต CPE: ' + confName;

  var html = '<div style="font-family:\'Sarabun\',sans-serif;max-width:600px;margin:0 auto;color:#1a202c">';
  html += '<div style="background:linear-gradient(135deg,#059669,#10b981);padding:24px 28px;border-radius:12px 12px 0 0">';
  html += '<div style="color:#fff;font-size:1.2rem;font-weight:700">✅ ผลการพิจารณารับรองหน่วยกิต CPE</div>';
  html += '<div style="color:rgba(255,255,255,.8);font-size:.85rem;margin-top:4px">คณะเภสัชศาสตร์ มหาวิทยาลัยบูรพา</div>';
  html += '</div>';
  html += '<div style="background:#fff;border:1px solid #e5e7eb;border-top:none;padding:24px 28px;border-radius:0 0 12px 12px">';
  html += '<p>เรียน ผู้ส่งคำขอพิจารณาหน่วยกิต CPE</p>';
  html += '<p>คณะเภสัชศาสตร์ มหาวิทยาลัยบูรพา ได้พิจารณาผลการรับรองหน่วยกิต CPE สำหรับงานประชุมวิชาการดังต่อไปนี้เรียบร้อยแล้ว</p>';
  html += '<table style="width:100%;border-collapse:collapse;margin:16px 0;background:#f0fdf4;border-radius:8px;overflow:hidden">';
  html += '<tr><td style="padding:9px 14px;font-size:.82rem;color:#6b7280;width:160px;border-bottom:1px solid #d1fae5">ชื่องานประชุม</td><td style="padding:9px 14px;font-size:.88rem;font-weight:700;border-bottom:1px solid #d1fae5">' + confName + '</td></tr>';
  html += '<tr><td style="padding:9px 14px;font-size:.82rem;color:#6b7280;border-bottom:1px solid #d1fae5">หน่วยงาน</td><td style="padding:9px 14px;font-size:.88rem;border-bottom:1px solid #d1fae5">' + orgName + '</td></tr>';
  html += '<tr><td style="padding:9px 14px;font-size:.82rem;color:#6b7280">วันที่จัด</td><td style="padding:9px 14px;font-size:.88rem">' + dateRange + '</td></tr>';
  html += '</table>';

  if (data.resultWordUrl) {
    html += '<table style="width:100%;border-collapse:collapse;margin:20px 0">';
    html += '<tr><td colspan="2" style="padding:10px 0 8px;font-size:.88rem;font-weight:700;color:#1a202c;border-bottom:2px solid #059669">ผลการพิจารณา</td></tr>';
    html += '<tr><td style="padding:10px 14px 10px 0;font-size:.85rem;color:#374151;width:290px;border-bottom:1px solid #e5e7eb">เอกสารผลการพิจารณารับรองหน่วยกิต CPE</td>';
    html += '<td style="padding:10px 0;border-bottom:1px solid #e5e7eb"><a href="' + data.resultWordUrl + '" style="color:#4f46e5;font-weight:700">ดาวน์โหลดเอกสาร</a></td></tr>';
    html += '</table>';
  }

  html += '<p style="font-size:.85rem;color:#6b7280;margin-top:20px">หากมีข้อสงสัยกรุณาติดต่อ คณะเภสัชศาสตร์ มหาวิทยาลัยบูรพา</p>';
  html += '</div></div>';

  try {
    MailApp.sendEmail({ to: userEmail, subject: subject, htmlBody: html });
  } catch(e) { /* ไม่ให้ email error กระทบ main flow */ }

  // อัพเดต Sheet
  updateConferenceRow({
    confId:        data.confId        || '',
    resultWordUrl: data.resultWordUrl || '',
    confirmedAt:   data.confirmedAt,
    status:        'ยืนยันผลแล้ว'
  });

  return respond({ success: true });
}

/* ─── Send expert email ─── */
function sendExpertEmail(data) {
  var expertEmails = data.expertEmails || [];
  if (expertEmails.length === 0) {
    return respond({ success: false, error: 'ไม่มีอีเมลผู้ทรงคุณวุฒิ' });
  }

  var dateRange = data.dateStart || '-';
  if (data.dateEnd && data.dateEnd !== data.dateStart) dateRange += ' – ' + data.dateEnd;

  var sendTimeStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
  var subject = '[CPE] ขอเรียนเชิญพิจารณาหน่วยกิต CPE: ' + (data.confName || 'ประชุมวิชาการ') + ' (' + sendTimeStr + ')';

  var html = '<div style="font-family:\'Sarabun\',sans-serif;max-width:600px;margin:0 auto;color:#1a202c">';
  html += '<div style="background:linear-gradient(135deg,#4f46e5,#7c3aed);padding:24px 28px;border-radius:12px 12px 0 0">';
  html += '<div style="color:#fff;font-size:1.2rem;font-weight:700">🎓 คำขอพิจารณาหน่วยกิต CPE</div>';
  html += '<div style="color:rgba(255,255,255,.8);font-size:.85rem;margin-top:4px">คณะเภสัชศาสตร์ มหาวิทยาลัยบูรพา</div>';
  html += '</div>';
  html += '<div style="background:#fff;border:1px solid #e5e7eb;border-top:none;padding:24px 28px;border-radius:0 0 12px 12px">';
  html += '<p>เรียน ผู้ทรงคุณวุฒิ</p>';
  html += '<p>ขอความอนุเคราะห์ผู้ทรงคุณวุฒิพิจารณาหน่วยกิต CPE สำหรับงานประชุมวิชาการดังต่อไปนี้</p>';

  html += '<table style="width:100%;border-collapse:collapse;margin:16px 0;background:#f8fafc;border-radius:8px;overflow:hidden">';
  html += '<tr><td style="padding:9px 14px;font-size:.82rem;color:#6b7280;width:130px;border-bottom:1px solid #e5e7eb">ชื่องานประชุม</td><td style="padding:9px 14px;font-size:.88rem;font-weight:700;border-bottom:1px solid #e5e7eb">' + (data.confName || '-') + '</td></tr>';
  html += '<tr><td style="padding:9px 14px;font-size:.82rem;color:#6b7280;border-bottom:1px solid #e5e7eb">หน่วยงาน</td><td style="padding:9px 14px;font-size:.88rem;border-bottom:1px solid #e5e7eb">' + (data.orgName || '-') + '</td></tr>';
  html += '<tr><td style="padding:9px 14px;font-size:.82rem;color:#6b7280;border-bottom:1px solid #e5e7eb">วันที่จัด</td><td style="padding:9px 14px;font-size:.88rem;border-bottom:1px solid #e5e7eb">' + dateRange + '</td></tr>';
  if (data.location) html += '<tr><td style="padding:9px 14px;font-size:.82rem;color:#6b7280;border-bottom:1px solid #e5e7eb">สถานที่</td><td style="padding:9px 14px;font-size:.88rem;border-bottom:1px solid #e5e7eb">' + data.location + '</td></tr>';
  html += '</table>';

  // ─── กลุ่ม 1: เอกสารแนบ ───
  html += '<table style="width:100%;border-collapse:collapse;margin:20px 0 0">';
  html += '<tr><td colspan="2" style="padding:10px 0 8px;font-size:.88rem;font-weight:700;color:#1a202c;border-bottom:2px solid #4f46e5">เอกสารแนบ</td></tr>';
  if (data.projectUrl)  html += '<tr><td style="padding:10px 14px 10px 0;font-size:.85rem;color:#374151;width:290px;border-bottom:1px solid #e5e7eb">รายละเอียดโครงการ</td><td style="padding:10px 0;border-bottom:1px solid #e5e7eb"><a href="' + data.projectUrl  + '" style="color:#4f46e5;font-weight:700">เปิดเอกสาร</a></td></tr>';
  if (data.cvfileUrl)   html += '<tr><td style="padding:10px 14px 10px 0;font-size:.85rem;color:#374151;border-bottom:1px solid #e5e7eb">CV วิทยากร</td><td style="padding:10px 0;border-bottom:1px solid #e5e7eb"><a href="' + data.cvfileUrl   + '" style="color:#4f46e5;font-weight:700">เปิดเอกสาร</a></td></tr>';
  if (data.scheduleUrl) html += '<tr><td style="padding:10px 14px 10px 0;font-size:.85rem;color:#374151;border-bottom:1px solid #e5e7eb">กำหนดการ</td><td style="padding:10px 0;border-bottom:1px solid #e5e7eb"><a href="' + data.scheduleUrl + '" style="color:#4f46e5;font-weight:700">เปิดเอกสาร</a></td></tr>';
  html += '</table>';
  // ─── กลุ่ม 2: มติ CPE (แยกออกมา) ───
  if (data.matiPdfUrl || data.matiWordUrl) {
    html += '<table style="width:100%;border-collapse:collapse;margin:0 0 20px">';
    html += '<tr><td colspan="2" style="padding:16px 0 8px;font-size:.88rem;font-weight:700;color:#1a202c;border-bottom:2px solid #0891b2">มติผลการพิจารณารับรอง CPE</td></tr>';
    if (data.matiPdfUrl)  html += '<tr><td style="padding:10px 14px 10px 0;font-size:.85rem;color:#374151;width:290px;border-bottom:1px solid #e5e7eb">มติผลการพิจารณารับรอง CPE (PDF)</td><td style="padding:10px 0;border-bottom:1px solid #e5e7eb"><a href="' + data.matiPdfUrl  + '" style="color:#4f46e5;font-weight:700">เปิดเอกสาร</a></td></tr>';
    if (data.matiWordUrl) html += '<tr><td style="padding:10px 14px 10px 0;font-size:.85rem;color:#374151;width:290px;border-bottom:1px solid #e5e7eb">มติผลการพิจารณารับรอง CPE (Word)</td><td style="padding:10px 0;border-bottom:1px solid #e5e7eb"><a href="' + data.matiWordUrl + '" style="color:#4f46e5;font-weight:700">เปิดเอกสาร</a></td></tr>';
    html += '</table>';
  }
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

/**
 * نظام بلاغات أسئلة الاختبارات — Google Apps Script backend.
 * Configure Script Properties: SHEET_ID and FOLDER_ID (see README).
 * Deploy as Web App: Execute as Me; Who has access: Anyone.
 * IMPORTANT: Public web app; deploy only after an administrator accepts the security limitations.
 */
const HEADERS = [
  'رقم البلاغ', 'تاريخ الاستلام', 'صفة المبلّغ', 'الاسم', 'الرقم الجامعي/الوظيفي',
  'البريد الإلكتروني', 'رقم الهاتف', 'المرحلة', 'رمز المقرر', 'اسم المقرر', 'نوع الاختبار',
  'نوع السؤال', 'تحديد الشعبة/الاختبار', 'نص السؤال المتذكَّر', 'خيارات متذكَّرة',
  'نوع الإشكال المبلغ عنه', 'تفصيل الإشكال', 'رابط المرفق', 'حالة البلاغ',
  'نتيجة المراجعة العلمية', 'رأي المحاضر', 'المعالجة المعتمدة', 'رابط السؤال في النظام',
  'تاريخ التنفيذ', 'إعادة التصحيح', 'ملاحظات داخلية'
];
const MAX_FILE_BYTES = 2 * 1024 * 1024;
const ALLOWED_MIMES = ['image/jpeg', 'image/png', 'image/webp', 'application/pdf'];
const MAX_FIELD = {fullName:120,identifier:40,email:160,phone:25,examDetails:180,questionText:2000,rememberedOptions:1500,issueDetails:3000};

function doGet() { return HtmlService.createHtmlOutput('خدمة استقبال البلاغات تعمل. استخدم رابط النموذج المنشور لتقديم بلاغ.'); }

function doPost(e) {
  const p = (e && e.parameter) || {};
  const requestId = String(p.requestId || '').slice(0,100);
  try {
    if (p.website) throw new Error('تعذر قبول الطلب.');
    const props = PropertiesService.getScriptProperties();
    const sheetId = props.getProperty('SHEET_ID');
    const folderId = props.getProperty('FOLDER_ID');
    if (!sheetId || !folderId) throw new Error('خدمة استقبال البلاغات غير مهيأة بعد.');
    const needed = ['role','fullName','email','courseId','examType','questionType','questionText','issueType','issueDetails'];
    for (const key of needed) if (!String(p[key] || '').trim()) throw new Error('يرجى استكمال جميع البيانات المطلوبة.');
    for (const [key, length] of Object.entries(MAX_FIELD)) if (String(p[key] || '').length > length) throw new Error('أحد الحقول تجاوز الطول المسموح.');
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(p.email))) throw new Error('البريد الإلكتروني غير صالح.');
    const courseId = Number(p.courseId);
    if (!Number.isInteger(courseId) || courseId < 1 || courseId > 173) throw new Error('المقرر غير صالح.');
    const courseList = getCourses_();
    const c = courseList[courseId - 1];
    if (!c) throw new Error('المقرر غير موجود.');

    let fileBlob = null;
    if (p.fileBase64) {
      const mime = String(p.fileMime || '');
      if (!ALLOWED_MIMES.includes(mime)) throw new Error('نوع الملف غير مسموح.');
      if (String(p.fileBase64).length > Math.ceil(MAX_FILE_BYTES * 4 / 3) + 10) throw new Error('حجم الملف أكبر من المسموح.');
      const bytes = Utilities.base64Decode(String(p.fileBase64));
      if (bytes.length > MAX_FILE_BYTES) throw new Error('حجم الملف أكبر من المسموح.');
      const expectedMagic = {
        'application/pdf':[37,80,68,70], 'image/png':[137,80,78,71],
        'image/jpeg':[255,216,255], 'image/webp':[82,73,70,70]
      };
      const magic = expectedMagic[mime];
      if (!magic.every((b, i) => ((bytes[i] + 256) % 256) === b)) throw new Error('محتوى الملف لا يطابق نوعه.');
      if (mime === 'image/webp' && String.fromCharCode(...bytes.slice(8,12)) !== 'WEBP') throw new Error('ملف الصورة غير صالح.');
      const extension = {'application/pdf':'pdf','image/png':'png','image/jpeg':'jpg','image/webp':'webp'}[mime];
      const safeFileName = `attachment-${Utilities.getUuid()}.${extension}`;
      fileBlob = Utilities.newBlob(bytes, mime, safeFileName);
    }

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) throw new Error('الخدمة مشغولة حاليًا. حاول مجددًا.');
    let ref, fileUrl = '';
    try {
      const ss = SpreadsheetApp.openById(sheetId);
      let sheet = ss.getSheetByName('البلاغات');
      if (!sheet) sheet = ss.insertSheet('البلاغات');
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(HEADERS);
        sheet.setFrozenRows(1);
        sheet.getRange(1,1,1,HEADERS.length).setFontWeight('bold').setBackground('#dcece5');
      }
      const serial = sheet.getLastRow();
      ref = `EX-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd')}-${String(serial).padStart(5,'0')}`;
      if (fileBlob) {
        const file = DriveApp.getFolderById(folderId).createFile(fileBlob);
        file.setDescription(`مرفق البلاغ ${ref} — للاستخدام الداخلي فقط`);
        fileUrl = file.getUrl(); // Files inherit the folder's permission; never publish publicly.
      }
      sheet.appendRow([
        ref, new Date(), safe_(p.role), safe_(p.fullName), safe_(p.identifier), safe_(p.email), safe_(p.phone),
        safe_(c.level), safe_(c.code), safe_(c.name), safe_(p.examType), safe_(p.questionType),
        safe_(p.examDetails), safe_(p.questionText), safe_(p.rememberedOptions),
        safe_(p.issueType), safe_(p.issueDetails), fileUrl, 'جديد', '', '', '', '', '', '', ''
      ]);
      SpreadsheetApp.flush();
    } finally { lock.releaseLock(); }
    return respond_({kind:'exam-issue-result',requestId,ok:true,reference:ref});
  } catch(err) {
    console.error('Submission failure',err);
    return respond_({kind:'exam-issue-result',requestId,ok:false,message:String(err.message||'تعذر حفظ البلاغ')});
  }
}

function safe_(value) {
  const x = String(value || '').trim();
  // Prevent spreadsheet formula injection while retaining the original visible data.
  return /^[=+\-@\t\r]/.test(x) ? "'" + x : x;
}

function respond_(data) {
  // No personal data is posted back. The client checks the message origin and request nonce.
  const json = JSON.stringify(data).replace(/</g,'\\u003c');
  return HtmlService.createHtmlOutput('<!doctype html><html><body><script>parent.postMessage(' + json + ',"*");<\/script></body></html>')
    .setTitle('نتيجة إرسال البلاغ');
}

// عند ترقية جدول موجود، راجع ترتيب الأعمدة قبل استقبال بلاغات جديدة.
function initialize() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  const folderId = props.getProperty('FOLDER_ID');
  if (!id || !folderId) throw new Error('أدخل SHEET_ID و FOLDER_ID في خصائص السكربت أولًا.');
  const sheet = SpreadsheetApp.openById(id).getSheetByName('البلاغات') || SpreadsheetApp.openById(id).insertSheet('البلاغات');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS); sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,HEADERS.length).setFontWeight('bold').setBackground('#dcece5');
  }
  DriveApp.getFolderById(folderId).getName();
}

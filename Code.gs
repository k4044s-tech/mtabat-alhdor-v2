/**
 * Apps Script Backend لتطبيق "متابعة الحضور" (V2)
 * يُلصق هذا الكود في: الشيت "متابعة الحضور V29" ← Extensions ← Apps Script
 * بعد اللصق: Project Settings ← Script Properties ← أضف SHARED_SECRET بأي قيمة سرية،
 * ثم Deploy ← New deployment ← Web app ← Execute as: Me ← Who has access: Anyone.
 *
 * ⚠️ إلزامي: التطبيق لن يعمل إطلاقاً (كل الطلبات تُرفض عمداً) حتى تُضبط
 * SHARED_SECRET في Script Properties. هذا "فشل مغلق" مقصود: بما أن النشر
 * "Anyone has access"، لا يجوز أن يعمل أي إجراء بلا مصادقة بالخطأ.
 * بعد ضبط السر، الصقه أيضاً في الواجهة (تبويب الإعدادات ← رابط Apps Script).
 */

var SHEET_TEACHERS = 'المعلمون';
var SHEET_ATTENDANCE = 'الحضور';
var SHEET_ABSENCE = 'الغياب';
var SHEET_IMPORTS = 'الاستيرادات';
var SHEET_SETTINGS = 'الإعدادات';

var DEFAULT_SCHEDULE_PERIODS = [
  { id: 'summer', label: 'صيفي', workStart: '6:45', workEnd: '12:30', dateFrom: '', dateTo: '' },
  { id: 'winter', label: 'شتوي', workStart: '7:00', workEnd: '12:45', dateFrom: '', dateTo: '' },
  { id: 'ramadan', label: 'رمضان', workStart: '9:30', workEnd: '13:00', dateFrom: '', dateTo: '' },
];
var DEFAULT_SETTINGS = {
  schedulePeriods: DEFAULT_SCHEDULE_PERIODS,
  defaultPeriodId: 'summer',
  weekendDays: [5, 6],
  holidays: [],
  // تاريخ مباشرة الدوام لكل فئة: لا يُسجَّل غياب لأي موظف قبل تاريخ مباشرته (المعلمون عادةً
  // يباشرون بعد الإداريين). يمكن تجاوزه لموظف بعينه بعمود startDate في شيت المعلمين.
  roleStartDates: { 'معلم': '', 'إداري': '' },
};
var STAFF_ROLES = ['معلم', 'إداري'];
var WEEKDAY_AR_ = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];

function doPost(e) {
  var payload = JSON.parse(e.postData.contents);
  return handleRequest_(payload);
}

// أمنياً: GET مقصور على "ping" فقط (فحص اتصال بلا بيانات). كل الإجراءات الأخرى
// (قراءة أو كتابة) يجب أن تمر عبر POST حصراً — الواجهة نفسها لا تستخدم GET إطلاقاً،
// وروابط GET تُسجَّل في history المتصفح وسجلات الخادم وقد تُحمَّل بالخطأ (مثلاً
// كصورة مضمّنة في صفحة أخرى)، فلا داعٍ لتوسيع سطح الهجوم بلا فائدة فعلية.
function doGet(e) {
  var payload = {};
  for (var k in e.parameter) payload[k] = e.parameter[k];
  if (payload.action !== 'ping') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'GET غير مسموح لهذا الإجراء — استخدم POST' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return handleRequest_(payload);
}

function handleRequest_(payload) {
  var out;
  try {
    var secret = PropertiesService.getScriptProperties().getProperty('SHARED_SECRET');
    if (!secret) {
      // فشل مغلق: بدون SHARED_SECRET مضبوط في Script Properties نرفض كل الطلبات
      // بدل السماح بها بلا مصادقة. النشر على "Anyone has access" يجعل الرابط عاماً،
      // فيجب ألا يعمل أي إجراء قبل ضبط السر عمداً من قِبل مدير المدرسة.
      out = { ok: false, error: 'التطبيق غير مُهيَّأ: لم يُضبط SHARED_SECRET في Script Properties' };
    } else if (payload.token !== secret) {
      out = { ok: false, error: 'unauthorized' };
    } else {
      switch (payload.action) {
        case 'ping':
          out = { ok: true, time: Date.now() };
          break;
        case 'getTeachers':
          out = handleGetTeachers_();
          break;
        case 'getDay':
          out = handleGetDay_(payload);
          break;
        case 'importDay':
          out = handleImportDay_(payload);
          break;
        case 'importPermissions':
          out = handleImportPermissions_(payload);
          break;
        case 'importComprehensiveReport':
          out = handleImportComprehensiveReport_(payload);
          break;
        case 'getPermissionsRange':
          out = handleGetPermissionsRange_(payload);
          break;
        case 'deleteDay':
          out = handleDeleteDay_(payload);
          break;
        case 'getAbsenceRange':
          out = handleGetAbsenceRange_(payload);
          break;
        case 'getAttendanceRange':
          out = handleGetAttendanceRange_(payload);
          break;
        case 'getRangeBundle':
          out = handleGetRangeBundle_(payload);
          break;
        case 'getMonthlySummary':
          out = handleGetMonthlySummary_(payload);
          break;
        case 'getAnalytics':
          out = handleGetAnalytics_(payload);
          break;
        case 'listImports':
          out = handleListImports_();
          break;
        case 'getEmployeeReport':
          out = handleGetEmployeeReport_(payload);
          break;
        case 'getAnnualPenalties':
          out = handleGetAnnualPenalties_(payload);
          break;
        case 'getSettings':
          out = { ok: true, settings: getSettingsObj_() };
          break;
        case 'saveSettings':
          out = handleSaveSettings_(payload);
          break;
        case 'clearCache':
          // إبطال يدوي لذاكرة التخزين المؤقت — ضروري بعد أي تحديث لمنطق الحساب نفسه بالخادم
          // (لا بيانات جديدة)، لأن الكتابة العادية فقط هي ما يُبطل الكاش تلقائياً، لا رفع كود جديد.
          bumpCacheGen_();
          out = { ok: true };
          break;
        case 'previewRecalculate':
          out = handlePreviewRecalculate_(payload);
          break;
        case 'applyRecalculate':
          out = handleApplyRecalculate_(payload);
          break;
        case 'updateAbsence':
          out = handleUpdateAbsence_(payload);
          break;
        case 'bulkUpdateAbsence':
          out = handleBulkUpdateAbsence_(payload);
          break;
        case 'getAlerts':
          out = handleGetAlerts_();
          break;
        case 'markActionDone':
          out = handleMarkActionDone_(payload);
          break;
        case 'getActionStatuses':
          out = handleGetActionStatuses_();
          break;
        case 'getAllTeachers':
          out = handleGetAllTeachers_();
          break;
        case 'saveTeachers':
          out = handleSaveTeachers_(payload);
          break;
        case 'deleteTeacher':
          out = handleDeleteTeacher_(payload);
          break;
        case 'reapplyPermissions':
          out = handleReapplyPermissions_(payload);
          break;
        case 'getFieldNotes':
          out = handleGetFieldNotes_(payload);
          break;
        case 'saveFieldNote':
          out = handleSaveFieldNote_(payload);
          break;
        case 'deleteFieldNote':
          out = handleDeleteFieldNote_(payload);
          break;
        case 'authInfo':
          out = handleAuthInfo_();
          break;
        case 'register':
          out = handleRegister_(payload);
          break;
        case 'login':
          out = handleLogin_(payload);
          break;
        case 'verifySession':
          out = handleVerifySession_(payload);
          break;
        case 'forgotPassword':
          out = handleForgotPassword_(payload);
          break;
        case 'resetPassword':
          out = handleResetPassword_(payload);
          break;
        case 'changePassword':
          out = handleChangePassword_(payload);
          break;
        case 'listUsers':
          out = handleListUsers_(payload);
          break;
        case 'updateUser':
          out = handleUpdateUser_(payload);
          break;
        case 'exportAll':
          out = handleExportAll_();
          break;
        default:
          out = { ok: false, error: 'unknown action: ' + payload.action };
      }
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// عناوين الشيتات الأساسية — يُنشأ أي شيت ناقص تلقائياً بها، فيكفي لصق هذا الكود في شيت
// جديد فارغ تماماً ليعمل التطبيق (أو لتُعاد عليه عمليات نسخة محلية جديدة عند المزامنة).
function defaultHeadersFor_(name) {
  var m = {};
  m[SHEET_TEACHERS] = ['civil', 'num', 'name', 'active', 'role', 'startDate'];
  m[SHEET_ATTENDANCE] = ['id', 'date', 'civil', 'num', 'name', 'checkIn', 'checkOut', 'hoursRaw', 'lateMinutes', 'earlyMinutes', 'noCheckout', 'schedStart', 'schedEnd', 'schedTolerance', 'schedLabel', 'schedId', 'updatedAt'];
  m[SHEET_ABSENCE] = ['id', 'date', 'civil', 'num', 'name', 'classification', 'reason', 'note', 'updatedAt'];
  m[SHEET_IMPORTS] = ['date', 'importedAt', 'fileName', 'schedStart', 'schedEnd', 'schedLabel', 'presentCount', 'absentCount', 'updatedAt'];
  m[SHEET_SETTINGS] = ['key', 'value', 'updatedAt'];
  return m[name] || null;
}

function getSheet_(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    var headers = defaultHeadersFor_(name);
    if (headers) return getOrCreateSheet_(name, headers);
    throw new Error('sheet not found: ' + name);
  }
  return sheet;
}

// ==================== تخزين مؤقت (Cache) لتسريع جلب البيانات ====================
// كل قراءة نطاق (getAttendanceRange وغيرها) تفتح الشيت وتقرأه فعلياً في كل استدعاء،
// وهذا هو السبب الرئيسي للبطء عند تكرار فتح نفس الشاشة أو التنقل بين التبويبات بلا
// تغيّر فعلي بالبيانات. نُخزّن نتيجة كل قراءة (حتى ٦ ساعات) في CacheService
// (خدمة Google سريعة جداً، لا تفتح الشيت إطلاقاً)، ونُبطل كل النسخ المخزَّنة فوراً بمجرد
// حدوث أي كتابة فعلية (استيراد/حذف/تعديل غياب/تحديث حالة إجراء) عبر رقم "جيل" الكاش —
// فيستحيل أن يرى المستخدم بيانات قديمة بعد أي تعديل، مهما طالت مدة الصلاحية.
function cacheGen_() {
  return PropertiesService.getScriptProperties().getProperty('CACHE_GEN') || '0';
}
function bumpCacheGen_() {
  var props = PropertiesService.getScriptProperties();
  var v = (Number(props.getProperty('CACHE_GEN')) || 0) + 1;
  props.setProperty('CACHE_GEN', String(v));
}
// ينفّذ computeFn ويخزّن نتيجتها، أو يعيد النتيجة المخزَّنة سابقاً إن وُجدت لنفس المفتاح
// وجيل الكاش الحالي. أي عطل في الكاش نفسه (مفتاح طويل، حجم كبير، خدمة معطّلة) يُتجاهل
// بأمان وتُحسب النتيجة من الشيت مباشرة — الكاش تسريع اختياري لا يعتمد عليه صحة البيانات.
function cached_(actionKey, params, computeFn) {
  var cache = CacheService.getScriptCache();
  var rawKey = actionKey + ':' + JSON.stringify(params);
  var key = 'v' + cacheGen_() + ':' + (rawKey.length > 200
    ? Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, rawKey))
    : rawKey);
  try {
    var cachedJson = cache.get(key);
    if (cachedJson) return JSON.parse(cachedJson);
  } catch (e) { /* تجاهل عطل القراءة من الكاش */ }

  var result = computeFn();

  try {
    var json = JSON.stringify(result);
    // المدة هنا لا تؤثر على صحة البيانات إطلاقاً — أي كتابة فعلية (استيراد/حذف/تعديل)
    // تُبطل كل النسخ المخزَّنة فوراً عبر رقم الجيل أعلاه بغضّ النظر عن هذه المدة، فرفعناها
    // للحد الأقصى المسموح به من Google (٦ ساعات) بدل ٩٠ ثانية فقط — لا خسارة في الصحة
    // ومكسب حقيقي في السرعة لبقية الجلسة.
    if (json.length < 95000) cache.put(key, json, 21600);
  } catch (e) { /* تجاهل عطل الكتابة في الكاش (حجم كبير جداً مثلاً) — النتيجة صحيحة على أي حال */ }

  return result;
}

// أعمدة تحتاج تنسيقاً خاصاً عند القراءة، لأن Google Sheets قد يحوّل نصوصاً تشبه
// التاريخ/الوقت (مثل "2026-08-16" أو "6:45") تلقائياً إلى خلايا Date عند الكتابة —
// سواء كتبها هذا الكود أو أي إدخال يدوي سابق في الشيت.
var DATE_ONLY_COLUMNS = { date: 1 };
var TIME_12H_COLUMNS = { checkIn: 1, checkOut: 1 };
var TIME_24H_COLUMNS = { schedStart: 1, schedEnd: 1, hoursRaw: 1 };

function isDateValue_(v) {
  return Object.prototype.toString.call(v) === '[object Date]';
}

// يعيد قيمة الخلية كما ينبغي أن تظهر للتطبيق: نص عادي إذا حوّلتها Sheets تلقائياً
// إلى Date/Time، أو القيمة كما هي غير ذلك.
function formatCellForOutput_(header, value) {
  if (!isDateValue_(value)) return value;
  var tz = Session.getScriptTimeZone();
  if (DATE_ONLY_COLUMNS[header]) return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  if (TIME_12H_COLUMNS[header]) return Utilities.formatDate(value, tz, 'hh:mm a');
  if (TIME_24H_COLUMNS[header]) return Utilities.formatDate(value, tz, 'H:mm');
  return Utilities.formatDate(value, tz, "yyyy-MM-dd'T'HH:mm:ss");
}

function sheetToObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = formatCellForOutput_(headers[j], values[i][j]);
    }
    rows.push(row);
  }
  return rows;
}

// أسرع من sheetToObjects_ للبحث ضمن نطاق تاريخي: يقرأ عمود التاريخ وحده أولاً (رخيص)
// بدل كل الأعمدة، ثم يقرأ فقط الكتلة المحصورة بين أول وآخر صف مطابق للنطاق — بدل
// قراءة وتنسيق كل صفوف الشيت (بما فيها تواريخ خارج النطاق المطلوب) في كل استدعاء.
// هذا هو أهم عامل في سرعة الملخص/التحليلات/سجل الغياب/التقارير مع نمو البيانات
// عبر السنوات الدراسية، لأنه يوقف نمو زمن الاستجابة مع كِبر تاريخ الشيت الكامل ويجعله
// متناسباً فقط مع حجم الفترة المطلوبة فعلياً.
function getRowsForDateRange_(sheet, from, to) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var dateCol = headers.indexOf('date') + 1;
  if (dateCol === 0) {
    return sheetToObjects_(sheet).filter(function (r) {
      var d = normalizeDate_(r.date);
      return d >= from && d <= to;
    });
  }

  var dateValues = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues();
  var matchRows = [];
  for (var i = 0; i < dateValues.length; i++) {
    var d = normalizeDate_(dateValues[i][0]);
    if (d >= from && d <= to) matchRows.push(i + 2);
  }
  if (!matchRows.length) return [];

  var minRow = matchRows[0];
  var maxRow = matchRows[matchRows.length - 1];
  var block = sheet.getRange(minRow, 1, maxRow - minRow + 1, lastCol).getValues();
  var matchSet = {};
  matchRows.forEach(function (r) { matchSet[r] = true; });

  var out = [];
  for (var i = 0; i < block.length; i++) {
    var rowNum = minRow + i;
    if (!matchSet[rowNum]) continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = formatCellForOutput_(headers[j], block[i][j]);
    out.push(obj);
  }
  return out;
}

function getRowsForDate_(sheet, date) {
  return getRowsForDateRange_(sheet, date, date);
}

function normalizeDate_(v) {
  if (isDateValue_(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}

// يُرجع قيم المصفوفة بلا تكرار، بترتيب أول ظهور — تُستخدم لتلخيص عيّنات التشخيص (مثل
// السجلات المدنية غير المطابقة) بلا الاعتماد على Set حتى يبقى الأسلوب موحَّداً مع بقية الملف.
function uniqueValues_(arr) {
  var seen = {}, out = [];
  arr.forEach(function (v) {
    var k = String(v);
    if (seen[k]) return;
    seen[k] = true;
    out.push(v);
  });
  return out;
}

function isActive_(v) {
  return v === true || String(v).toLowerCase() === 'true';
}

// يقرأ "H:MM" أو "HH:MM AM/PM" ويرجّع عدد الدقائق من منتصف الليل، أو null إن تعذّر
function parseTimeToMinutes_(s) {
  if (!s) return null;
  var m = String(s).trim().match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM|am|pm)?$/);
  if (!m) return null;
  var h = parseInt(m[1], 10);
  var min = parseInt(m[2], 10);
  var ap = m[3] ? m[3].toUpperCase() : null;
  if (ap === 'PM' && h < 12) h += 12;
  if (ap === 'AM' && h === 12) h = 0;
  return h * 60 + min;
}

function formatMinutesAsHM_(totalMin) {
  if (totalMin < 0) totalMin += 24 * 60;
  var h = Math.floor(totalMin / 60);
  var m = totalMin % 60;
  return h + ':' + (m < 10 ? '0' : '') + m;
}

// أعمدة نصية تبدو كتاريخ/وقت — تُفرَض عليها تنسيقة "نص عادي" قبل الكتابة حتى لا
// تحوّلها Sheets تلقائياً إلى خلية Date (وإلا نفقد الصيغة الأصلية عند القراءة لاحقاً).
// أُضيفت أيضاً classification/reason/note/fileName لأنها نصوص حرة يُحتمل أن تبدأ
// بـ "=" (مثلاً "=HYPERLINK(...)")، فتُفسَّرها Sheets كصيغة بدل نص عادي إن لم تُفرض
// عليها تنسيقة "نص" — وهذا يحمي من حقن الصيغ (formula injection).
var FORCE_TEXT_COLUMNS = {
  date: 1, checkIn: 1, checkOut: 1, hoursRaw: 1,
  schedStart: 1, schedEnd: 1, schedLabel: 1, schedId: 1,
  classification: 1, reason: 1, note: 1, fileName: 1,
  requestId: 1, fromDate: 1, toDate: 1, fromTime: 1, toTime: 1, type: 1, status: 1,
  name: 1, civil: 1, startDate: 1, role: 1,
  username: 1, email: 1, salt: 1, hash: 1, resetHash: 1,
  location: 1, verifiedBy: 1, witness: 1, createdBy: 1, punchKind: 1, kind: 1,
};

// upsert بمطابقة عمود المعرّف (idColName) — يحدّث الصف الموجود أو يضيف صفاً جديداً
// يضيف تلقائياً أي عمود ناقص من "names" في نهاية صف العناوين إن لم يكن موجوداً بالفعل —
// ترقية آمنة لشيت قديم بعمود جديد (مثل presentCount/absentCount) بلا تدخل يدوي من المستخدم.
function ensureHeaders_(sheet, names) {
  var lastCol = sheet.getLastColumn();
  var headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var missing = names.filter(function (n) { return headers.indexOf(n) === -1; });
  if (!missing.length) return;
  sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
}

function upsertRows_(sheet, idColName, records) {
  if (!records.length) return;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol = headers.indexOf(idColName);
  if (idCol === -1) throw new Error('id column not found: ' + idColName);

  var idToRow = {};
  for (var i = 1; i < data.length; i++) idToRow[normalizeDate_(data[i][idCol])] = i + 1;
  var nextRow = data.length + 1;

  var textColIndexes = [];
  headers.forEach(function (h, idx) {
    if (FORCE_TEXT_COLUMNS[h]) textColIndexes.push(idx + 1); // 1-based
  });

  records.forEach(function (rec) {
    var id = normalizeDate_(rec[idColName]);
    var targetRow = idToRow[id];
    if (!targetRow) {
      targetRow = nextRow;
      idToRow[id] = targetRow;
      nextRow++;
    }
    textColIndexes.forEach(function (colIdx) {
      sheet.getRange(targetRow, colIdx).setNumberFormat('@');
    });
    var rowArr = headers.map(function (h) {
      return Object.prototype.hasOwnProperty.call(rec, h) ? rec[h] : '';
    });
    sheet.getRange(targetRow, 1, 1, headers.length).setValues([rowArr]);
  });
}

// نسخة مُحسَّنة من upsertRows_ للكتابات الكبيرة (مئات/آلاف السجلات دفعة واحدة، كما في
// استيراد التقرير الشامل لعدة أسابيع أو أشهر): تقرأ الشيت مرة واحدة، تدمج التحديثات
// والإضافات في مصفوفة بالذاكرة، ثم تكتب الشيت كاملاً بنداء setValues واحد — بدل
// upsertRows_ التي تفتح نطاقاً (Range) منفصلاً لكل صف (وأخرى لكل عمود نصي بذلك الصف)،
// فتصبح بطيئة جداً (آلاف نداءات API) مع آلاف السجلات دفعة واحدة.
function batchUpsertRows_(sheet, idColName, records) {
  if (!records.length) return;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol = headers.indexOf(idColName);
  if (idCol === -1) throw new Error('id column not found: ' + idColName);

  var idToRow = {}; // id -> فهرس الصف داخل data (0-based، يشمل صف العناوين)
  for (var i = 1; i < data.length; i++) idToRow[normalizeDate_(data[i][idCol])] = i;

  records.forEach(function (rec) {
    var id = normalizeDate_(rec[idColName]);
    var rowArr = headers.map(function (h) {
      return Object.prototype.hasOwnProperty.call(rec, h) ? rec[h] : '';
    });
    if (Object.prototype.hasOwnProperty.call(idToRow, id)) {
      data[idToRow[id]] = rowArr;
    } else {
      data.push(rowArr);
      idToRow[id] = data.length - 1;
    }
  });

  // نص عادي إجبارياً لكل الأعمدة الحسّاسة قبل الكتابة (نفس حماية upsertRows_ من حقن الصيغ) —
  // عمود كامل بنداء واحد بدل خلية بخلية، فهي أيضاً أسرع بكثير على هذا الحجم من البيانات.
  var textColIndexes = [];
  headers.forEach(function (h, idx) { if (FORCE_TEXT_COLUMNS[h]) textColIndexes.push(idx + 1); });
  // صفوف قديمة حوّلتها Sheets سابقاً إلى خلايا Date تُقرأ هنا كائنات Date — نعيدها نصاً
  // بصيغتها الصحيحة قبل كتابتها في عمود "نص"، وإلا قد تُخزَّن كنص طويل مثل
  // "Sun Aug 16 2026 00:00:00 GMT+0300" فلا تطابق صيغة yyyy-MM-dd عند الفلترة لاحقاً.
  for (var r = 1; r < data.length; r++) {
    textColIndexes.forEach(function (c) {
      var v = data[r][c - 1];
      if (isDateValue_(v)) data[r][c - 1] = formatCellForOutput_(headers[c - 1], v);
    });
  }
  textColIndexes.forEach(function (colIdx) {
    sheet.getRange(1, colIdx, data.length, 1).setNumberFormat('@');
  });
  sheet.getRange(1, 1, data.length, headers.length).setValues(data);
}

// حذف دفعي لصفوف بمعرّفاتها (idSet: كائن id -> true) — قراءة واحدة، تصفية، ثم كتابة واحدة،
// بدل deleteRow() لكل صف (deleteRowsByIds_) التي تبطؤ جداً مع مئات الصفوف في الاستيراد الشامل.
function batchDeleteIds_(sheet, idSet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  var idCol = data[0].indexOf('id');
  if (idCol === -1) return;
  var kept = data.filter(function (row, i) { return i === 0 || !idSet[normalizeDate_(row[idCol])]; });
  if (kept.length === data.length) return;
  sheet.getRange(1, 1, data.length, data[0].length).clearContent();
  sheet.getRange(1, 1, kept.length, data[0].length).setValues(kept);
}

function deleteRowsByIds_(sheet, ids) {
  if (!ids.length) return;
  var idSet = {};
  ids.forEach(function (id) { idSet[id] = true; });
  var data = sheet.getDataRange().getValues();
  var idCol = data[0].indexOf('id');
  for (var i = data.length - 1; i >= 1; i--) {
    if (idSet[normalizeDate_(data[i][idCol])]) sheet.deleteRow(i + 1);
  }
}

// يحذف كل صفوف تاريخ معيّن من الحضور/الغياب/الاستيرادات — لتصحيح استيراد خاطئ
function handleDeleteDay_(payload) {
  var date = String(payload.date);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // حذف دفعي (قراءة، تصفية، إعادة كتابة مرة واحدة) بدل deleteRow() لكل صف على حدة —
    // يوم كامل قد يعني ٣٠-٥٠+ معلّماً، فكانت هذه العملية تُصدر بهذا العدد من نداءات API
    // منفصلة لكل شيت بدل اثنين فقط (مسح ثم إعادة كتابة).
    [SHEET_ATTENDANCE, SHEET_ABSENCE, SHEET_IMPORTS].forEach(function (name) {
      var sheet = getSheet_(name);
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) return;
      var headers = data[0];
      var dateCol = headers.indexOf('date');
      var kept = data.filter(function (row, idx) {
        return idx === 0 || normalizeDate_(row[dateCol]) !== date;
      });
      if (kept.length === data.length) return; // لا صفوف بهذا التاريخ بهذا الشيت — لا حاجة للكتابة
      sheet.getRange(1, 1, data.length, headers.length).clearContent();
      sheet.getRange(1, 1, kept.length, headers.length).setValues(kept);
    });
    return { ok: true, date: date };
  } finally {
    bumpCacheGen_();
    lock.releaseLock();
  }
}

function handleGetTeachers_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('teachers_v1');
  if (cached) return JSON.parse(cached);

  var teachers = sheetToObjects_(getSheet_(SHEET_TEACHERS)).filter(function (t) {
    return isActive_(t.active);
  });
  var result = { ok: true, teachers: teachers };
  cache.put('teachers_v1', JSON.stringify(result), 300); // 5 دقائق — الروستر نادراً ما يتغيّر
  return result;
}

function handleGetDay_(payload) {
  var date = String(payload.date);
  var attendance = getRowsForDate_(getSheet_(SHEET_ATTENDANCE), date);
  attendance = annotateWithPermissions_(attendance, date, date);
  var absence = getRowsForDate_(getSheet_(SHEET_ABSENCE), date);
  var importInfo = getRowsForDate_(getSheet_(SHEET_IMPORTS), date)[0] || null;
  return { ok: true, date: date, attendance: attendance, absence: absence, importInfo: importInfo };
}

function handleImportDay_(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var date = String(payload.date);
    var schedStart = payload.schedStart || '';
    var schedEnd = payload.schedEnd || '';
    var schedLabel = payload.schedLabel || '';
    var schedTolerance = Number(payload.schedTolerance || 0);
    var schedId = payload.schedId || '';
    var rows = payload.rows || [];

    var teachers = sheetToObjects_(getSheet_(SHEET_TEACHERS));
    var teacherByCivil = {};
    teachers.forEach(function (t) {
      teacherByCivil[String(t.civil)] = t;
    });

    // نحتفظ بتصنيف/سبب/ملاحظة الغياب اليدوية الموجودة مسبقاً لهذا التاريخ عند إعادة الاستيراد
    var existingAbsenceByCivil = {};
    getRowsForDate_(getSheet_(SHEET_ABSENCE), date).forEach(function (a) {
      existingAbsenceByCivil[String(a.civil)] = a;
    });

    var schedStartMin = parseTimeToMinutes_(schedStart);
    var schedEndMin = parseTimeToMinutes_(schedEnd);
    var now = Date.now();
    var presentCivils = {};
    var attendanceRecords = [];
    var skippedUnmatched = []; // صفوف بسجل مدني لا يطابق أي معلّم بالروستر — نُبلّغ عنها بدل إسقاطها بصمت
    var noFingerprintCivils = []; // معلمون وردوا بالملف بلا أي بصمة — يُعامَلون غائبين

    rows.forEach(function (r) {
      var civil = String(r.civil);
      var t = teacherByCivil[civil];
      if (!t) { skippedUnmatched.push(civil); return; }
      // صف بلا أي بصمة (لا حضور ولا انصراف) ليس حضوراً: تقرير البصمة يسرد أحياناً كل الموظفين
      // ومنهم الغائب والمجاز بخانات فارغة — كان يُحسب "حاضراً" فلا يُسجَّل غيابه ولا تُربط إجازته.
      if (!hasFingerprint_(r)) { noFingerprintCivils.push(civil); return; }
      presentCivils[civil] = true;

      var checkInMin = parseTimeToMinutes_(r.checkIn);
      var checkOutMin = r.checkOut ? parseTimeToMinutes_(r.checkOut) : null;
      var lateMinutes = checkInMin != null && schedStartMin != null
        ? Math.max(0, checkInMin - schedStartMin - schedTolerance)
        : 0;
      var earlyMinutes = checkOutMin != null && schedEndMin != null
        ? Math.max(0, schedEndMin - checkOutMin)
        : 0;
      var noCheckout = checkOutMin == null;
      var hoursRaw = checkInMin != null && checkOutMin != null
        ? formatMinutesAsHM_(checkOutMin - checkInMin)
        : '';

      attendanceRecords.push({
        id: date + '__' + civil,
        date: date,
        civil: civil,
        num: t.num,
        name: t.name,
        checkIn: r.checkIn || '',
        checkOut: r.checkOut || '',
        hoursRaw: hoursRaw,
        lateMinutes: lateMinutes,
        earlyMinutes: earlyMinutes,
        noCheckout: noCheckout,
        schedStart: schedStart,
        schedEnd: schedEnd,
        schedTolerance: schedTolerance,
        schedLabel: schedLabel,
        schedId: schedId,
        updatedAt: now,
      });
    });

    var daySettings = getSettingsObj_();
    var skipAbsenceMarking = isNonWorkingDay_(date, daySettings);
    // استئذانات اليوم: غياب يغطيه استئذان "مقبول" ليوم كامل يُصنَّف "بعذر" فوراً عند الاستيراد —
    // بدون هذا كان الربط يحدث فقط لحظة استيراد ملف الاستئذانات، فأي يوم يُستورد بعده يبقى
    // "لم يُدخل عذر" رغم وجود إجازته المعتمدة.
    var dayPermsByKey = buildPermissionsByCivilDate_(date, date);
    var absenceRecords = [];
    teachers.forEach(function (t) {
      if (!isActive_(t.active)) return;
      var civil = String(t.civil);
      if (presentCivils[civil]) return;
      if (skipAbsenceMarking) return; // يوم عطلة/نهاية أسبوع: لا نُسجّل غياباً لمن لم يظهر بالملف
      if (beforeStart_(t, date, daySettings)) return; // لم يباشر بعد (ما زال في إجازته قبل تاريخ مباشرته)
      var existing = existingAbsenceByCivil[civil];
      var classification = existing ? existing.classification : '';
      var reason = existing ? existing.reason : '';
      if (!classification) {
        var fullDay = acceptedFullDayPermission_(dayPermsByKey, civil, date);
        if (fullDay) { classification = 'بعذر'; reason = mapPermissionTypeToReason_(fullDay.type); }
      }
      absenceRecords.push({
        id: date + '__' + civil,
        date: date,
        civil: civil,
        num: t.num,
        name: t.name,
        classification: classification,
        reason: reason,
        note: existing ? existing.note : '',
        updatedAt: now,
      });
    });

    upsertRows_(getSheet_(SHEET_ATTENDANCE), 'id', attendanceRecords);
    upsertRows_(getSheet_(SHEET_ABSENCE), 'id', absenceRecords);

    // احذف أي سجل غياب/حضور سابق لنفس اليوم أصبح غير صحيح الآن (مثلاً معلّم كان غائباً وحضر اليوم)
    var absentCivilsNow = {};
    absenceRecords.forEach(function (a) { absentCivilsNow[a.civil] = true; });
    var staleAbsenceIds = Object.keys(existingAbsenceByCivil)
      .filter(function (c) { return !absentCivilsNow[c]; })
      .map(function (c) { return date + '__' + c; });
    deleteRowsByIds_(getSheet_(SHEET_ABSENCE), staleAbsenceIds);
    // وبالمثل: معلّم كان "حاضراً" في استيراد سابق لنفس اليوم ولم يعد موجوداً في الملف الجديد
    // يُحذف سجل حضوره القديم — وإلا بقي حاضراً وغائباً معاً فتُحسب له أيام حضور زائدة
    // (ويُضاف تأخيره/انصرافه القديم) في الملخص والتحليلات والتقارير.
    var existingAttendance = getRowsForDate_(getSheet_(SHEET_ATTENDANCE), date);
    var staleAttendanceIds = existingAttendance
      .filter(function (r) { return !presentCivils[String(r.civil)]; })
      .map(function (r) { return date + '__' + String(r.civil); });
    deleteRowsByIds_(getSheet_(SHEET_ATTENDANCE), staleAttendanceIds);
    var importsSheet_ = getSheet_(SHEET_IMPORTS);
    ensureHeaders_(importsSheet_, ['presentCount', 'absentCount']);
    upsertRows_(importsSheet_, 'date', [
      {
        date: date,
        importedAt: new Date().toISOString(),
        fileName: payload.fileName || '',
        schedStart: schedStart,
        schedEnd: schedEnd,
        schedLabel: schedLabel,
        presentCount: attendanceRecords.length,
        absentCount: absenceRecords.length,
        updatedAt: now,
      },
    ]);

    return {
      ok: true,
      date: date,
      presentCount: attendanceRecords.length,
      absentCount: absenceRecords.length,
      noFingerprintCount: noFingerprintCivils.length, skippedUnmatchedCount: skippedUnmatched.length,
      skippedUnmatchedCivils: skippedUnmatched.slice(0, 20), // عيّنة كافية للتشخيص بلا إثقال الاستجابة
    };
  } finally {
    bumpCacheGen_();
    lock.releaseLock();
  }
}

// ==================== استيراد تقرير شامل (فترة طويلة، كل الموظفين دفعة واحدة) ====================
// ملف "تقرير الحضور والانصراف لجميع الموظفين" من برنامج حضوري — يغطي أسابيع أو أشهر في ملف
// واحد، بخلاف importDay الذي يستورد يوماً واحداً فقط. الواجهة تُحلّل الملف وترسل صفوفاً
// منظَّفة مسبقاً بالشكل {civil, date, kind, checkIn?, checkOut?, noCheckout?, reason?} حيث
// kind هو present أو absent أو absent_excused (الشرح الكامل في الواجهة عند قراءة الملف).
// التأخير/الانصراف المبكر يُحسبان هنا من البصمة الفعلية ودوام ذلك التاريخ وفق الإعدادات
// الحالية — لا من أي عمود جاهز بالملف — بنفس منطق importDay تماماً، فتبقى الأرقام متسقة
// مهما كان مصدر الاستيراد. تستخدم batchUpsertRows_ (كتابة دفعة واحدة) بدل upsertRows_
// (صف بصف) لأن حجم البيانات هنا قد يبلغ آلاف الصفوف عبر عشرات الأيام دفعة واحدة.
function handleImportComprehensiveReport_(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var rows = payload.rows || [];
    if (!rows.length) return { ok: true, datesProcessed: 0, attendanceWritten: 0, absenceWritten: 0, skippedUnmatchedCount: 0 };

    var teachers = sheetToObjects_(getSheet_(SHEET_TEACHERS));
    var teacherByCivil = {};
    teachers.forEach(function (t) { teacherByCivil[String(t.civil)] = t; });

    var settings = getSettingsObj_();

    var dateSet = {};
    rows.forEach(function (r) { dateSet[String(r.date)] = true; });
    var dateList = Object.keys(dateSet).sort();
    var from = dateList[0], to = dateList[dateList.length - 1];
    // نفس آلية importDay: استئذان "مقبول" ليوم كامل يُصنَّف الغياب "بعذر" بسببه تلقائياً
    var permsByKey = (from && to) ? buildPermissionsByCivilDate_(from, to) : {};

    // نحتفظ بتصنيف/سبب/ملاحظة الغياب اليدوية الموجودة مسبقاً لهذه الأيام عند إعادة الاستيراد
    var existingAbsenceByKey = {};
    if (from && to) {
      getRowsForDateRange_(getSheet_(SHEET_ABSENCE), from, to).forEach(function (a) {
        existingAbsenceByKey[normalizeDate_(a.date) + '__' + String(a.civil)] = a;
      });
    }

    var now = Date.now();
    var attendanceRecords = [];
    var absenceRecords = [];
    var skippedUnmatched = [];
    var schedCache = {};
    function scheduleFor(date) {
      if (!schedCache[date]) schedCache[date] = resolveScheduleForDate_(date, settings);
      return schedCache[date];
    }

    rows.forEach(function (r) {
      var civil = String(r.civil);
      var date = String(r.date);
      var t = teacherByCivil[civil];
      if (!t) { skippedUnmatched.push(civil); return; }
      if (isNonWorkingDay_(date, settings)) return; // حماية ثانية: عطلة حسب الإعدادات الحالية
      if (beforeStart_(t, date, settings)) return; // لم يباشر بعد بهذا التاريخ

      var sched = scheduleFor(date);
      var schedStartMin = parseTimeToMinutes_(sched.workStart);
      var schedEndMin = parseTimeToMinutes_(sched.workEnd);

      if (r.kind === 'present') {
        var checkInMin = r.checkIn ? parseTimeToMinutes_(r.checkIn) : null;
        var noCheckout = !!r.noCheckout || !r.checkOut;
        var checkOutMin = noCheckout ? null : parseTimeToMinutes_(r.checkOut);
        var lateMinutes = checkInMin != null && schedStartMin != null ? Math.max(0, checkInMin - schedStartMin) : 0;
        var earlyMinutes = checkOutMin != null && schedEndMin != null ? Math.max(0, schedEndMin - checkOutMin) : 0;
        var hoursRaw = checkInMin != null && checkOutMin != null ? formatMinutesAsHM_(checkOutMin - checkInMin) : '';
        attendanceRecords.push({
          id: date + '__' + civil, date: date, civil: civil, num: t.num, name: t.name,
          checkIn: r.checkIn || '', checkOut: noCheckout ? '' : (r.checkOut || ''),
          hoursRaw: hoursRaw, lateMinutes: lateMinutes, earlyMinutes: earlyMinutes, noCheckout: noCheckout,
          schedStart: sched.workStart, schedEnd: sched.workEnd, schedTolerance: 0,
          schedLabel: sched.label, schedId: sched.id, updatedAt: now,
        });
      } else {
        var existing = existingAbsenceByKey[date + '__' + civil];
        var classification = existing ? existing.classification : '';
        var reason = existing ? existing.reason : '';
        if (!classification) {
          if (r.kind === 'absent_excused') {
            classification = 'بعذر';
            reason = r.reason || 'إجازة';
          } else {
            var fullDay = acceptedFullDayPermission_(permsByKey, civil, date);
            if (fullDay) { classification = 'بعذر'; reason = mapPermissionTypeToReason_(fullDay.type); }
          }
        }
        absenceRecords.push({
          id: date + '__' + civil, date: date, civil: civil, num: t.num, name: t.name,
          classification: classification, reason: reason, note: existing ? existing.note : '', updatedAt: now,
        });
      }
    });

    // من صار حاضراً بالملف يُحذف سجل غيابه القديم لنفس اليوم، ومن صار غائباً يُحذف حضوره —
    // نفس ما يفعله importDay، وإلا بقي الشخص حاضراً وغائباً معاً فتتضخم أرقامه بالتقارير.
    // يقتصر على (التاريخ + السجل المدني) الواردين بالملف فقط؛ لا يمسّ أي سجل آخر.
    var presentIds = {}, absentIds = {};
    attendanceRecords.forEach(function (r) { presentIds[r.id] = true; });
    absenceRecords.forEach(function (r) { absentIds[r.id] = true; });
    batchDeleteIds_(getSheet_(SHEET_ABSENCE), presentIds);
    batchDeleteIds_(getSheet_(SHEET_ATTENDANCE), absentIds);

    batchUpsertRows_(getSheet_(SHEET_ATTENDANCE), 'id', attendanceRecords);
    batchUpsertRows_(getSheet_(SHEET_ABSENCE), 'id', absenceRecords);

    // سجل الاستيرادات: صف واحد لكل تاريخ مشمول فعلياً بالملف، بعدد الحاضرين/الغائبين الحقيقي
    var presentCountByDate = {}, absentCountByDate = {};
    attendanceRecords.forEach(function (r) { presentCountByDate[r.date] = (presentCountByDate[r.date] || 0) + 1; });
    absenceRecords.forEach(function (r) { absentCountByDate[r.date] = (absentCountByDate[r.date] || 0) + 1; });
    var importsSheet_ = getSheet_(SHEET_IMPORTS);
    ensureHeaders_(importsSheet_, ['presentCount', 'absentCount']);
    var importRecords = dateList.filter(function (d) { return presentCountByDate[d] || absentCountByDate[d]; }).map(function (d) {
      var sched = scheduleFor(d);
      return {
        date: d, importedAt: new Date().toISOString(), fileName: payload.fileName || 'تقرير شامل',
        schedStart: sched.workStart, schedEnd: sched.workEnd, schedLabel: sched.label,
        presentCount: presentCountByDate[d] || 0, absentCount: absentCountByDate[d] || 0, updatedAt: now,
      };
    });
    batchUpsertRows_(importsSheet_, 'date', importRecords);

    return {
      ok: true,
      datesProcessed: importRecords.length,
      attendanceWritten: attendanceRecords.length,
      absenceWritten: absenceRecords.length,
      skippedUnmatchedCount: skippedUnmatched.length,
      skippedUnmatchedDistinctCount: uniqueValues_(skippedUnmatched).length,
      // القائمة كاملة بلا اقتصاص — الواجهة تعرض عيّنة منها فقط، لكنها تحتاج الكاملة لبناء
      // سجلات معلمين جدد منها عند الطلب (زر "إضافة الموظفين غير المطابقين").
      skippedUnmatchedCivils: uniqueValues_(skippedUnmatched),
      // عيّنة من السجلات المدنية الفعلية بالروستر — تُعرض بالواجهة بجانب الأرقام غير المطابقة
      // حتى يقارن المستخدم الشكل (عدد الخانات، وجود أصفار بادئة...) بنفسه دون تخمين من هنا.
      rosterSampleCivils: teachers.slice(0, 5).map(function (t) { return String(t.civil); }),
      rosterCount: teachers.length,
    };
  } finally {
    bumpCacheGen_();
    lock.releaseLock();
  }
}

// ==================== الاستئذانات (ملف "تقرير الاستئذانات" من برنامج حضوري) ====================
// طلب استئذان "مقبول" بلا وقت محدد (أو يمتد لعدة أيام) هو غياب معتمَد رسمياً بسبب معروف —
// يُصنَّف يوم الغياب المطابق تلقائياً. طلب "مقبول" بوقت محدد (نفس اليوم) هو إذن للتغيّب
// جزءاً من الدوام — يُستخدم لاحقاً لاستثناء تأخير/انصراف مبكر متقاطع معه زمنياً بالكامل.
var SHEET_PERMISSIONS = 'الاستئذانات';
var SHEET_PERMISSIONS_HEADERS = ['requestId', 'civil', 'name', 'fromTime', 'toTime', 'fromDate', 'toDate', 'status', 'type'];

// يحوّل نوع الاستئذان الرسمي لأقرب سبب غياب معروف بالنظام؛ يُرجع النص الأصلي إن لم يوجد
// تطابق فيظهر كخيار إضافي بقائمة الأسباب تلقائياً (نفس آلية أي سبب غير معروف مسبقاً).
function mapPermissionTypeToReason_(type) {
  var t = String(type || '').trim();
  var map = {
    'اجازة- سنوية': 'الإجازة العادية', 'اجازة - سنوية': 'الإجازة العادية', 'اجازة سنوية': 'الإجازة العادية',
    'شخصي - طارئ': 'الإجازة الاضطرارية', 'شخصي- طارئ': 'الإجازة الاضطرارية',
    'شخصي - طبي': 'الإجازة المرضية', 'شخصي- طبي': 'الإجازة المرضية',
    'رسمية - تدريب': 'تدريب', 'رسمية- تدريب': 'تدريب',
    'رسمية - مهمة عمل': 'مهمة عمل', 'رسمية- مهمة عمل': 'مهمة عمل',
  };
  return map[t] || t;
}

function handleImportPermissions_(payload) {
  var rows = payload.rows || [];
  if (!rows.length) return { ok: true, imported: 0, classifiedDays: 0 };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var permSheet = getOrCreateSheet_(SHEET_PERMISSIONS, SHEET_PERMISSIONS_HEADERS);
    var records = rows.map(function (r) {
      return {
        requestId: String(r.requestId), civil: String(r.civil), name: r.name || '',
        fromTime: r.fromTime || '', toTime: r.toTime || '',
        fromDate: r.fromDate || '', toDate: r.toDate || '',
        status: r.status || '', type: r.type || '',
      };
    });
    upsertRows_(permSheet, 'requestId', records);

    // تصنيف تلقائي كامل: كل طلب "مقبول" بلا وقت محدد يُطبَّق فوراً على أيام الغياب
    // المطابقة، ويستبدل أي تصنيف حالي دون سؤال (بحسب المطلوب صراحةً).
    var absSheet = getSheet_(SHEET_ABSENCE);
    var absData = absSheet.getDataRange().getValues();
    var absHeaders = absData[0];
    var idCol = absHeaders.indexOf('id');
    var classCol = absHeaders.indexOf('classification');
    var reasonCol = absHeaders.indexOf('reason');
    var updatedAtCol = absHeaders.indexOf('updatedAt');
    var now = Date.now();
    // فهرس معرّف← رقم صف يُبنى مرة واحدة (O(n)) بدل مسح كامل الشيت خطياً لكل يوم لكل طلب
    // استئذان — كان هذا O(الطلبات × الأيام × صفوف الغياب)، أصبح O(صفوف الغياب + الطلبات × الأيام).
    var absIdToRow = {};
    for (var ai = 1; ai < absData.length; ai++) absIdToRow[normalizeDate_(absData[ai][idCol])] = ai + 1;

    var classifiedDays = 0;
    var malformedDatePermissions = []; // طلبات "مقبول" بتاريخ غير قياسي (yyyy-MM-dd) — لا تُصنَّف صامتة
    var isoDateRe = /^\d{4}-\d{2}-\d{2}$/;
    records.forEach(function (r) {
      if (String(r.status).trim() !== 'مقبول') return;
      if (isPersonalPermission_(r.type)) return; // الشخصي لا يُعتمد ليوم كامل
      if (String(r.fromTime || '').trim() || String(r.toTime || '').trim()) return; // استئذان جزئي بوقت محدد — ليس غياب يوم كامل
      if (!r.fromDate || !r.toDate) return;
      if (!isoDateRe.test(r.fromDate) || !isoDateRe.test(r.toDate)) { malformedDatePermissions.push(r.requestId); return; }
      var reason = mapPermissionTypeToReason_(r.type);
      // نحسب عدد الأيام عبر daysBetweenISO_ ونتنقّل بـ addDaysISO_ (نص "yyyy-MM-dd" مباشرة)
      // بدل new Date(نص التاريخ) + setDate() — الأخيرة تُفسَّر بتوقيت UTC دوماً عند التحويل
      // من نص ISO مباشرة، ما قد يُزيح اليوم بخطأ صامت خارج فرق +3 المطابق للسعودية تحديداً.
      var spanDays = daysBetweenISO_(r.fromDate, r.toDate);
      for (var dayOffset = 0; dayOffset <= spanDays; dayOffset++) {
        var dateStr = addDaysISO_(r.fromDate, dayOffset);
        var absId = dateStr + '__' + r.civil;
        var row = absIdToRow[absId];
        if (row) {
          absSheet.getRange(row, classCol + 1).setNumberFormat('@').setValue('بعذر');
          absSheet.getRange(row, reasonCol + 1).setNumberFormat('@').setValue(reason);
          absSheet.getRange(row, updatedAtCol + 1).setValue(now);
          classifiedDays++;
        }
      }
    });

    return { ok: true, imported: records.length, classifiedDays: classifiedDays, malformedDateCount: malformedDatePermissions.length };
  } finally {
    bumpCacheGen_();
    lock.releaseLock();
  }
}

// يضيف علامتي excusedLate/excusedEarly لكل صف حضور — صحيح فقط إذا وُجد استئذان "مقبول"
// بوقت محدد لنفس اليوم يغطّي كامل فجوة التأخير (من بداية الدوام حتى الحضور الفعلي) أو
// كامل فجوة الانصراف المبكر (من الانصراف الفعلي حتى نهاية الدوام) — تغطية جزئية لا تكفي.
// يبني خريطة (السجل المدني|التاريخ) ← كل طلبات الاستئذان (أي حالة، يوم كامل أو جزئي بوقت
// محدد) التي تغطي ذلك اليوم — أساس مشترك لكل من: عمود "الاستئذان" بجداول الغياب والتأخر،
// واستثناء التأخير/الانصراف المبكر الفعلي من قوائمهما.
function buildPermissionsByCivilDate_(from, to) {
  var rows = sheetToObjects_(getOrCreateSheet_(SHEET_PERMISSIONS, SHEET_PERMISSIONS_HEADERS));
  var map = {};
  rows.forEach(function (p) {
    var f = normalizeDate_(p.fromDate);
    var t = normalizeDate_(p.toDate || p.fromDate);
    if (!f || !t || t < from || f > to) return;
    // نفس إصلاح UTC/محلي أعلاه: تنقّل بالنص عبر addDaysISO_ بدل new Date(نص) + setDate().
    var spanDays = daysBetweenISO_(f, t);
    for (var dayOffset = 0; dayOffset <= spanDays; dayOffset++) {
      var dateStr = addDaysISO_(f, dayOffset);
      if (dateStr < from || dateStr > to) continue;
      var key = String(p.civil) + '|' + dateStr;
      (map[key] || (map[key] = [])).push(p);
    }
  });
  return map;
}

// ==================== ربط الغياب بالبصمة والاستئذانات (مصدر واحد للحقيقة) ====================
// القاعدة: المعلّم "حاضر" في يوم عمل فقط إن كانت له بصمة (حضور أو انصراف). غير ذلك "غائب".
// والغائب الذي يغطيه استئذان "مقبول" ليوم كامل (إجازة/مهمة/تدريب…) يُصنَّف "بعذر" بسببه
// تلقائياً — أياً كان ترتيب استيراد ملفات البصمة وملف الاستئذانات.
// تاريخ مباشرة الموظف الفعلي: تاريخه الخاص إن وُجد، وإلا تاريخ مباشرة فئته من الإعدادات
function staffRole_(t) {
  var r = String((t && t.role) || '').trim();
  return r === 'إداري' ? 'إداري' : 'معلم';
}
function staffStartDate_(t, settings) {
  var own = t && t.startDate ? normalizeDate_(t.startDate) : '';
  if (own && /^\d{4}-\d{2}-\d{2}$/.test(own)) return own;
  var map = (settings && settings.roleStartDates) || {};
  var d = map[staffRole_(t)] || '';
  return /^\d{4}-\d{2}-\d{2}$/.test(d) ? d : '';
}
function beforeStart_(t, date, settings) {
  var st = staffStartDate_(t, settings);
  return !!(st && date < st);
}

function hasFingerprint_(r) {
  return !!(String(r.checkIn || '').trim() || String(r.checkOut || '').trim());
}

// الاستئذان الشخصي (طبي/طارئ) لا يُعتمد ليوم كامل حسب الآليات المنظمة لحضوري (رابعاً/٢) —
// فلا يُعذر به غياب يوم كامل تلقائياً حتى لو ظهر "مقبول"؛ يبقى للمراجعة اليدوية.
function isPersonalPermission_(type) {
  return /^\s*شخص/.test(String(type || ''));
}

// استئذان "مقبول" بلا وقت محدد (إجازة/مهمة ليوم كامل أو أكثر) يغطي (السجل المدني، التاريخ)
function acceptedFullDayPermission_(permsByKey, civil, date) {
  var perms = permsByKey[String(civil) + '|' + date] || [];
  for (var i = 0; i < perms.length; i++) {
    var p = perms[i];
    if (String(p.status).trim() === 'مقبول' && !String(p.fromTime || '').trim() && !String(p.toTime || '').trim()
        && !isPersonalPermission_(p.type)) return p;
  }
  return null;
}

// إعادة ربط الاستئذانات بالغياب لفترة: كل يوم غياب "غير مصنّف" يغطيه استئذان مقبول ليوم
// كامل يُصنَّف "بعذر" بسببه. لا يلمس أي يوم صنّفه المستخدم يدوياً.
function handleReapplyPermissions_(payload) {
  var from = payload.from || '0000-01-01';
  var to = payload.to || '9999-12-31';
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // (أ) إصلاح بيانات قديمة: سجلات "حضور" بلا أي بصمة تُحوَّل إلى غياب
    var converted = convertEmptyAttendanceToAbsence_(from, to);
    // (ب) حذف غياب مسجّل قبل تاريخ مباشرة الموظف (يحدث عند تحديد/تعديل تاريخ المباشرة لاحقاً)
    var removedBeforeStart = removeAbsencesBeforeStart_(from, to);
    // (ج) حذف غياب مسجّل في يوم صار عطلة (إجازة أُضيفت للتقويم بعد استيراد ذلك اليوم)
    var removedNonWorking = removeAbsencesOnNonWorkingDays_(from, to);
    var permsByKey = buildPermissionsByCivilDate_(from, to);
    var absSheet = getSheet_(SHEET_ABSENCE);
    var data = absSheet.getDataRange().getValues();
    var h = data[0];
    var dateCol = h.indexOf('date'), civilCol = h.indexOf('civil'), classCol = h.indexOf('classification');
    var reasonCol = h.indexOf('reason'), updatedAtCol = h.indexOf('updatedAt');
    var now = Date.now();
    var fixed = 0;
    for (var i = 1; i < data.length; i++) {
      var d = normalizeDate_(data[i][dateCol]);
      if (!d || d < from || d > to) continue;
      if (String(data[i][classCol] || '').trim()) continue;
      var p = acceptedFullDayPermission_(permsByKey, data[i][civilCol], d);
      if (!p) continue;
      absSheet.getRange(i + 1, classCol + 1).setNumberFormat('@').setValue('بعذر');
      absSheet.getRange(i + 1, reasonCol + 1).setNumberFormat('@').setValue(mapPermissionTypeToReason_(p.type));
      if (updatedAtCol >= 0) absSheet.getRange(i + 1, updatedAtCol + 1).setValue(now);
      fixed++;
    }
    return { ok: true, classifiedDays: fixed, convertedToAbsence: converted, removedBeforeStart: removedBeforeStart, removedNonWorking: removedNonWorking };
  } finally {
    bumpCacheGen_();
    lock.releaseLock();
  }
}

// يحذف سجلات الحضور الفارغة (بلا حضور ولا انصراف) في الفترة ويُنشئ مكانها سجل غياب — إن كان
// يوم عمل وللمعلّم سجل نشط ولا يوجد له غياب مسجّل لنفس اليوم. يُرجع عدد ما حُوِّل.
function convertEmptyAttendanceToAbsence_(from, to) {
  var attSheet = getSheet_(SHEET_ATTENDANCE);
  var empty = getRowsForDateRange_(attSheet, from, to).filter(function (r) { return !hasFingerprint_(r); });
  if (!empty.length) return 0;
  var settings = getSettingsObj_();
  var teacherByCivil = {};
  sheetToObjects_(getSheet_(SHEET_TEACHERS)).forEach(function (t) { teacherByCivil[String(t.civil)] = t; });
  var absSheet = getSheet_(SHEET_ABSENCE);
  var existingAbs = {};
  getRowsForDateRange_(absSheet, from, to).forEach(function (a) { existingAbs[normalizeDate_(a.date) + '__' + String(a.civil)] = true; });
  var now = Date.now();
  var newAbs = [];
  var dropIds = [];
  empty.forEach(function (r) {
    var d = normalizeDate_(r.date);
    var civil = String(r.civil);
    dropIds.push(d + '__' + civil);
    var t = teacherByCivil[civil];
    if (!t || !isActive_(t.active) || isNonWorkingDay_(d, settings)) return;
    if (beforeStart_(t, d, settings)) return;
    if (existingAbs[d + '__' + civil]) return;
    newAbs.push({ id: d + '__' + civil, date: d, civil: civil, num: t.num, name: t.name,
      classification: '', reason: '', note: '', updatedAt: now });
  });
  deleteRowsByIds_(attSheet, dropIds);
  if (newAbs.length) upsertRows_(absSheet, 'id', newAbs);
  // تحديث أعداد الحضور/الغياب في سجل الاستيرادات للأيام المتأثرة
  var days = {};
  empty.forEach(function (r) { days[normalizeDate_(r.date)] = true; });
  Object.keys(days).forEach(function (d) { refreshImportCounts_(d); });
  return newAbs.length;
}

// الاستيراد اليومي والشامل لا يسجّلان غياباً في يوم عطلة أصلاً، لكن إجازة تُضاف للتقويم بعد
// استيراد أيامها (مثل اليوم الوطني) تترك غياب كل الموظفين فيها قائماً. هنا نحذفه بأثر رجعي.
// سجلات الحضور الفعلي (من بصم فعلاً يوم العطلة) لا تُمسّ. إن لم يبقَ لليوم أي حضور يُحذف سجل
// استيراده أيضاً فيظهر بالتقويم "عطلة" لا "مستورد"، وإلا تُحدَّث أعداده فقط.
function removeAbsencesOnNonWorkingDays_(from, to) {
  var settings = getSettingsObj_();
  var absSheet = getSheet_(SHEET_ABSENCE);
  var idSet = {}, days = {}, count = 0;
  getRowsForDateRange_(absSheet, from, to).forEach(function (a) {
    var d = normalizeDate_(a.date);
    if (!isNonWorkingDay_(d, settings)) return;
    idSet[d + '__' + String(a.civil)] = true; days[d] = true; count++;
  });
  if (!count) return 0;
  batchDeleteIds_(absSheet, idSet);

  var attSheet = getSheet_(SHEET_ATTENDANCE);
  var emptyDays = {};
  Object.keys(days).forEach(function (d) {
    if (getRowsForDate_(attSheet, d).length) refreshImportCounts_(d);
    else emptyDays[d] = true;
  });
  if (Object.keys(emptyDays).length) {
    var impSheet = getSheet_(SHEET_IMPORTS);
    var data = impSheet.getDataRange().getValues();
    var dCol = data[0].indexOf('date');
    var kept = data.filter(function (row, i) { return i === 0 || !emptyDays[normalizeDate_(row[dCol])]; });
    if (kept.length !== data.length) {
      impSheet.getRange(1, 1, data.length, data[0].length).clearContent();
      impSheet.getRange(1, 1, kept.length, data[0].length).setValues(kept);
    }
  }
  return count;
}

function removeAbsencesBeforeStart_(from, to) {
  var settings = getSettingsObj_();
  var teacherByCivil = {};
  sheetToObjects_(getSheet_(SHEET_TEACHERS)).forEach(function (t) { teacherByCivil[String(t.civil)] = t; });
  var absSheet = getSheet_(SHEET_ABSENCE);
  var ids = [];
  var days = {};
  getRowsForDateRange_(absSheet, from, to).forEach(function (a) {
    var t = teacherByCivil[String(a.civil)];
    var d = normalizeDate_(a.date);
    if (t && beforeStart_(t, d, settings)) { ids.push(d + '__' + String(a.civil)); days[d] = true; }
  });
  if (!ids.length) return 0;
  deleteRowsByIds_(absSheet, ids);
  Object.keys(days).forEach(function (d) { refreshImportCounts_(d); });
  return ids.length;
}

function refreshImportCounts_(date) {
  var impSheet = getSheet_(SHEET_IMPORTS);
  var data = impSheet.getDataRange().getValues();
  var h = data[0];
  var dCol = h.indexOf('date'), pCol = h.indexOf('presentCount'), aCol = h.indexOf('absentCount');
  if (dCol < 0 || pCol < 0 || aCol < 0) return;
  for (var i = 1; i < data.length; i++) {
    if (normalizeDate_(data[i][dCol]) !== date) continue;
    impSheet.getRange(i + 1, pCol + 1).setValue(getRowsForDate_(getSheet_(SHEET_ATTENDANCE), date).length);
    impSheet.getRange(i + 1, aCol + 1).setValue(getRowsForDate_(getSheet_(SHEET_ABSENCE), date).length);
  }
}

// عند وجود أكثر من طلب لنفس اليوم (نادر)، نُبرز "مقبول" أولاً، ثم "تحت الإجراء"، ثم "مرفوض" —
// أكثر الحالات أهمية للإداري عند مراجعة سجل الغياب أو التأخر.
function pickPrimaryPermission_(perms) {
  if (!perms || !perms.length) return null;
  var order = { 'مقبول': 0, 'تحت الإجراء': 1, 'مرفوض': 2 };
  return perms.slice().sort(function (a, b) {
    var oa = order[a.status] != null ? order[a.status] : 3;
    var ob = order[b.status] != null ? order[b.status] : 3;
    return oa - ob;
  })[0];
}

// سماحية بالدقائق عند تصنيف الاستئذان كـ"متعلق بالتأخير" أو "متعلق بالانصراف المبكر" —
// الاستئذانات نادراً ما تُسجَّل بمطابقة تامة لدقيقة بداية/نهاية الدوام بالضبط (كما في مثال
// حقيقي: استئذان ينتهي ١٢:٤١ ونهاية الدوام ١٢:٤٥ — أربع دقائق فرق لا تُبطل كونه استئذان انصراف).
var PERMISSION_EDGE_TOLERANCE_MIN = 20;

// استئذان "متعلق بالتأخير" هو ما يبدأ عند/قبل بداية الدوام تقريباً (يُبرِّر تأخّر الحضور)،
// و"متعلق بالانصراف المبكر" هو ما ينتهي عند/بعد نهاية الدوام تقريباً (يُبرِّر انصرافاً مبكراً).
// استئذان منتصف اليوم (يبدأ بعد بداية الدوام بوضوح وينتهي قبل نهايته بوضوح) لا علاقة له
// بأيّهما، فلا يظهر بشارة أي منهما — وهذا هو بالضبط الخطأ الذي وقعتُ فيه سابقاً: عرض استئذان
// انصراف بعد الظهر كأنه شارة مرتبطة بتأخّر الصباح لنفس الشخص لمجرد وقوعه بنفس اليوم.
function permissionRelevantToLate_(p, schedStartMin) {
  var pFrom = parseTimeToMinutes_(p.fromTime);
  return pFrom != null && schedStartMin != null && pFrom <= schedStartMin + PERMISSION_EDGE_TOLERANCE_MIN;
}
function permissionRelevantToEarly_(p, schedEndMin) {
  var pTo = parseTimeToMinutes_(p.toTime);
  return pTo != null && schedEndMin != null && pTo >= schedEndMin - PERMISSION_EDGE_TOLERANCE_MIN;
}

function annotateWithPermissions_(rows, from, to) {
  if (!rows.length) return rows;
  var permsByKey = buildPermissionsByCivilDate_(from, to);
  if (!Object.keys(permsByKey).length) return rows;

  return rows.map(function (r) {
    var key = String(r.civil) + '|' + normalizeDate_(r.date);
    var perms = permsByKey[key];
    if (!perms) return r;

    var copy = {};
    for (var k in r) copy[k] = r[k];
    var schedStartMin = parseTimeToMinutes_(r.schedStart);
    var schedEndMin = parseTimeToMinutes_(r.schedEnd);
    var checkInMin = parseTimeToMinutes_(r.checkIn);
    var checkOutMin = parseTimeToMinutes_(r.checkOut);

    // شارتان منفصلتان — واحدة لسياق التأخير وأخرى لسياق الانصراف المبكر — لا شارة واحدة
    // عامة، حتى لا يظهر استئذان انصراف بعد الظهر كأنه متعلق بتأخّر الصباح والعكس.
    var latePerm = perms.filter(function (p) { return p.fromTime && p.toTime && permissionRelevantToLate_(p, schedStartMin); });
    var earlyPerm = perms.filter(function (p) { return p.fromTime && p.toTime && permissionRelevantToEarly_(p, schedEndMin); });
    if (latePerm.length) {
      var lp = pickPrimaryPermission_(latePerm);
      copy.latePermissionStatus = lp.status;
      copy.latePermissionType = lp.type;
      var lpFrom = parseTimeToMinutes_(lp.fromTime), lpTo = parseTimeToMinutes_(lp.toTime);
      if (lpFrom != null && lpTo != null) copy.latePermissionDurationMin = lpTo - lpFrom;
    }
    if (earlyPerm.length) {
      var ep = pickPrimaryPermission_(earlyPerm);
      copy.earlyPermissionStatus = ep.status;
      copy.earlyPermissionType = ep.type;
    }

    // الاستثناء الفعلي: نسبيّ، لا "الكل أو لا شيء" — نطرح بالضبط عدد الدقائق التي يغطّيها
    // الاستئذان المعتمد من الفجوة (بداية الدوام حتى الحضور، أو الانصراف حتى نهاية الدوام)،
    // ونُبقي الباقي كتأخير/انصراف حقيقي محسوب. استئذان يغطي ٦٥ من أصل ٧٥ دقيقة تأخير يترك
    // ١٠ دقائق فقط محسوبة، بدل إسقاط الكل لمجرد عدم التطابق الحرفي التام.
    var approvedIntraDay = perms.filter(function (p) { return p.status === 'مقبول' && p.fromTime && p.toTime; });
    if (approvedIntraDay.length) {
      var origLate = Number(r.lateMinutes) || 0;
      var origEarly = Number(r.earlyMinutes) || 0;
      var excusedLateAmount = 0, excusedEarlyAmount = 0;
      approvedIntraDay.forEach(function (p) {
        var pFrom = parseTimeToMinutes_(p.fromTime);
        var pTo = parseTimeToMinutes_(p.toTime);
        if (pFrom == null || pTo == null) return;
        if (origLate > 0 && schedStartMin != null && checkInMin != null) {
          var os = Math.max(schedStartMin, pFrom), oe = Math.min(checkInMin, pTo);
          if (oe > os) excusedLateAmount += (oe - os);
        }
        if (origEarly > 0 && checkOutMin != null && schedEndMin != null) {
          var os2 = Math.max(checkOutMin, pFrom), oe2 = Math.min(schedEndMin, pTo);
          if (oe2 > os2) excusedEarlyAmount += (oe2 - os2);
        }
      });
      excusedLateAmount = Math.min(excusedLateAmount, origLate);
      excusedEarlyAmount = Math.min(excusedEarlyAmount, origEarly);
      if (excusedLateAmount > 0) {
        copy.excusedLateMinutes = excusedLateAmount;
        copy.lateMinutes = origLate - excusedLateAmount;
      }
      if (excusedEarlyAmount > 0) {
        copy.excusedEarlyMinutes = excusedEarlyAmount;
        copy.earlyMinutes = origEarly - excusedEarlyAmount;
      }
      // تبقى علامتا excusedLate/excusedEarly (يعتمد عليهما تبويب "سجل الانصراف" حالياً) صحيحتين
      // فقط عند تغطية كاملة (الباقي = صفر) — لا تغيير في ذلك الجزء تحديداً.
      copy.excusedLate = origLate > 0 && excusedLateAmount >= origLate;
      copy.excusedEarly = origEarly > 0 && excusedEarlyAmount >= origEarly;
    }

    return copy;
  });
}

// نسخة مبسَّطة لشيت الغياب: تُضيف شارة "الاستئذان" فقط (أي طلب مرتبط بنفس اليوم) بلا حساب
// استثناء — الغياب له آلية تصنيف تلقائي منفصلة أصلاً عند الاستيراد.
function annotatePermissionBadgeOnly_(rows, from, to) {
  if (!rows.length) return rows;
  var permsByKey = buildPermissionsByCivilDate_(from, to);
  if (!Object.keys(permsByKey).length) return rows;
  return rows.map(function (r) {
    var key = String(r.civil) + '|' + normalizeDate_(r.date);
    var perms = permsByKey[key];
    if (!perms) return r;
    var primary = pickPrimaryPermission_(perms);
    var copy = {};
    for (var k in r) copy[k] = r[k];
    copy.permissionStatus = primary.status;
    copy.permissionType = primary.type;
    // علامة للواجهة: يوجد استئذان مقبول ليوم كامل يغطي هذا الغياب
    if (acceptedFullDayPermission_(permsByKey, r.civil, normalizeDate_(r.date))) copy.permissionFullDay = true;
    return copy;
  });
}

// كشف طلبات الاستئذان لفترة — لتبويب "الاستئذانات" (عرض/بحث/فلترة/طباعة). الشيت صغير
// نسبياً (مئات الصفوف بحد أقصى بالعام) فقراءته كاملاً ثم الفلترة بالذاكرة كافية تماماً،
// بلا حاجة لتحسين النطاق المستخدَم مع شيتي الحضور/الغياب الأكبر حجماً بكثير.
function handleGetPermissionsRange_(payload) {
  var from = String(payload.from);
  var to = String(payload.to);
  return cached_('getPermissionsRange', { from: from, to: to }, function () {
    var allRows = sheetToObjects_(getOrCreateSheet_(SHEET_PERMISSIONS, SHEET_PERMISSIONS_HEADERS));
    // آخر تاريخ مُدخل فعلياً بكل بيانات الاستئذانات (بغضّ النظر عن الفترة المعروضة حالياً)
    // — يساعد على معرفة مدى حداثة آخر ملف تم استيراده دون الحاجة لتوسيع الفترة يدوياً.
    var lastEnteredDate = '';
    allRows.forEach(function (r) {
      var d = normalizeDate_(r.toDate || r.fromDate);
      if (d && d > lastEnteredDate) lastEnteredDate = d;
    });

    var rows = allRows.filter(function (r) {
      var f = normalizeDate_(r.fromDate);
      var t = normalizeDate_(r.toDate || r.fromDate);
      return f && f <= to && t >= from;
    });
    rows.sort(function (a, b) { return normalizeDate_(b.fromDate).localeCompare(normalizeDate_(a.fromDate)); });
    return { ok: true, from: from, to: to, rows: rows, lastEnteredDate: lastEnteredDate };
  });
}

function handleGetAbsenceRange_(payload) {
  var from = String(payload.from);
  var to = String(payload.to);
  return cached_('getAbsenceRange', { from: from, to: to }, function () {
    var rows = getRowsForDateRange_(getSheet_(SHEET_ABSENCE), from, to);
    rows.sort(function (a, b) {
      return normalizeDate_(a.date) < normalizeDate_(b.date) ? -1 : 1;
    });
    rows = annotatePermissionBadgeOnly_(rows, from, to);
    return { ok: true, from: from, to: to, rows: rows };
  });
}

// مثل handleGetAbsenceRange_ لكن لشيت الحضور — تُستخدم لبناء تقارير الفترة (كشوفات
// المعلمين، كشف التأخر للجميع، الكشف الشهري التجميعي) بمرور واحد على البيانات الخام
// بدل عدة طلبات منفصلة لكل معلّم أو كل يوم.
function handleGetAttendanceRange_(payload) {
  var from = String(payload.from);
  var to = String(payload.to);
  return cached_('getAttendanceRange', { from: from, to: to }, function () {
    var rows = getRowsForDateRange_(getSheet_(SHEET_ATTENDANCE), from, to);
    rows.sort(function (a, b) {
      return normalizeDate_(a.date) < normalizeDate_(b.date) ? -1 : 1;
    });
    rows = annotateWithPermissions_(rows, from, to);
    return { ok: true, from: from, to: to, rows: rows };
  });
}

// يُرجع الحضور والغياب معاً لنفس الفترة بطلب خادم واحد بدل طلبين منفصلين — تستخدمه
// الشاشات التي تحتاج الاثنين معاً (مثل "تقرير المعلمين") لتقليل عدد رحلات الشبكة، وكل
// رحلة إلى Apps Script لها كلفة زمنية ملموسة بحد ذاتها بغضّ النظر عن حجم البيانات.
function handleGetRangeBundle_(payload) {
  var from = String(payload.from);
  var to = String(payload.to);
  return cached_('getRangeBundle', { from: from, to: to }, function () {
    var attendance = getRowsForDateRange_(getSheet_(SHEET_ATTENDANCE), from, to);
    attendance = annotateWithPermissions_(attendance, from, to);
    var absence = getRowsForDateRange_(getSheet_(SHEET_ABSENCE), from, to);
    attendance.sort(function (a, b) { return normalizeDate_(a.date) < normalizeDate_(b.date) ? -1 : 1; });
    absence.sort(function (a, b) { return normalizeDate_(a.date) < normalizeDate_(b.date) ? -1 : 1; });
    return { ok: true, from: from, to: to, attendance: attendance, absence: absence };
  });
}

// ملخص شهري لكل معلّم مرتّب تنازلياً حسب "المخالفات" (غياب + تأخير + انصراف مبكر) — لشاشة الملخص والترتيب
function handleGetMonthlySummary_(payload) {
  var from = String(payload.from);
  var to = String(payload.to);
  return cached_('getMonthlySummary', { from: from, to: to }, function () {

  var teachers = sheetToObjects_(getSheet_(SHEET_TEACHERS));
  var daily = getRowsForDateRange_(getSheet_(SHEET_ATTENDANCE), from, to);
  daily = annotateWithPermissions_(daily, from, to);
  var absences = getRowsForDateRange_(getSheet_(SHEET_ABSENCE), from, to);
  var imports = getRowsForDateRange_(getSheet_(SHEET_IMPORTS), from, to);

  function blankEntry(civil, num, name, active) {
    return {
      civil: civil, num: num, name: name, active: active,
      presentDays: 0, absenceDays: 0, excused: 0, unexcused: 0, unclassified: 0,
      lateDays: 0, lateMinutes: 0, earlyDays: 0, earlyMinutes: 0, noCheckoutDays: 0,
      excusedLateMinutes: 0, excusedEarlyMinutes: 0,
    };
  }

  var byCivil = {};
  teachers.forEach(function (t) {
    byCivil[String(t.civil)] = blankEntry(String(t.civil), t.num, t.name, isActive_(t.active));
  });

  daily.forEach(function (r) {
    var civil = String(r.civil);
    var e = byCivil[civil] || (byCivil[civil] = blankEntry(civil, r.num, r.name, true));
    e.presentDays++;
    var late = Number(r.lateMinutes) || 0;
    var early = Number(r.earlyMinutes) || 0;
    // تُستثنى دقائق التأخير/الانصراف المبكر المغطّاة باستئذان معتمد بالكامل من أرقام
    // المخالفات بكل مكان (لا فقط جدولَي التأخر/الانصراف)، وتُحسب على حدة كـ"مستثنى باستئذان"
    // بدل حذفها بصمت — حتى تبقى مرئية وواضحة أينما ظهر مجموع التأخير أو الانصراف.
    // lateMinutes/earlyMinutes هنا صافية بالفعل بعد annotateWithPermissions_ (طُرح منها ما
    // يغطّيه الاستئذان نسبياً) — نجمعها كما هي، ونجمع الجزء المُستثنى فعلياً على حدة للعرض.
    if (late > 0) { e.lateDays++; e.lateMinutes += late; }
    if (early > 0) { e.earlyDays++; e.earlyMinutes += early; }
    e.excusedLateMinutes += Number(r.excusedLateMinutes) || 0;
    e.excusedEarlyMinutes += Number(r.excusedEarlyMinutes) || 0;
    if (String(r.noCheckout) === 'true') e.noCheckoutDays++;
  });

  absences.forEach(function (a) {
    var civil = String(a.civil);
    var e = byCivil[civil] || (byCivil[civil] = blankEntry(civil, a.num, a.name, true));
    e.absenceDays++;
    if (a.classification === 'بعذر') e.excused++;
    else if (a.classification === 'بدون عذر') e.unexcused++;
    else e.unclassified++;
  });

  var list = Object.keys(byCivil).map(function (k) { return byCivil[k]; });
  var workDaysCount = imports.length;
  var sumSettings = getSettingsObj_();
  var teacherByCivilSum = {};
  teachers.forEach(function (t) { teacherByCivilSum[String(t.civil)] = t; });
  var importDates = imports.map(function (im) { return normalizeDate_(im.date); });
  list.forEach(function (e) {
    // أيام العمل المحسوبة لكل موظف تبدأ من تاريخ مباشرته — لا تُخفّض نسبة حضوره أيامٌ قبلها
    var t = teacherByCivilSum[String(e.civil)];
    var st = t ? staffStartDate_(t, sumSettings) : '';
    var myWorkDays = st ? importDates.filter(function (d) { return d >= st; }).length : workDaysCount;
    e.workDays = myWorkDays;
    e.startDate = st;
    // الغياب بعذر (إجازة معتمدة) ليس مخالفة — المخالفات: غياب بدون عذر أو غير مصنّف + أيام تأخر + أيام انصراف مبكر
    e.violations = e.unexcused + e.unclassified + e.lateDays + e.earlyDays;
    e.attendanceRate = myWorkDays > 0 ? Math.min(1, e.presentDays / myWorkDays) : null;
  });
  list.sort(function (a, b) { return (b.violations - a.violations) || ((a.num || 0) - (b.num || 0)); });

  var totals = list.reduce(function (acc, e) {
    acc.absenceDays += e.absenceDays;
    acc.excused += e.excused;
    acc.unexcused += e.unexcused;
    acc.unclassified += e.unclassified;
    acc.lateDays += e.lateDays;
    acc.lateMinutes += e.lateMinutes;
    acc.earlyDays += e.earlyDays;
    acc.earlyMinutes += e.earlyMinutes;
    acc.noCheckoutDays += e.noCheckoutDays;
    acc.excusedLateMinutes += e.excusedLateMinutes;
    acc.excusedEarlyMinutes += e.excusedEarlyMinutes;
    return acc;
  }, { absenceDays: 0, excused: 0, unexcused: 0, unclassified: 0, lateDays: 0, lateMinutes: 0, earlyDays: 0, earlyMinutes: 0, noCheckoutDays: 0, excusedLateMinutes: 0, excusedEarlyMinutes: 0 });

  return { ok: true, from: from, to: to, list: list, totals: totals, workDays: imports.length };
  });
}

function classifyLate_(minutes) {
  if (!minutes || minutes <= 0) return null;
  if (minutes <= 10) return 'بسيط';
  if (minutes <= 30) return 'متوسط';
  return 'كبير';
}

// اتجاه يومي + شدة التأخير + أكثر أسباب الغياب تكراراً خلال فترة — لشاشة التحليلات
function handleGetAnalytics_(payload) {
  var from = String(payload.from);
  var to = String(payload.to);
  return cached_('getAnalytics', { from: from, to: to }, function () {

  var daily = getRowsForDateRange_(getSheet_(SHEET_ATTENDANCE), from, to);
  daily = annotateWithPermissions_(daily, from, to);
  var absences = getRowsForDateRange_(getSheet_(SHEET_ABSENCE), from, to);

  // تجميع مسبق حسب التاريخ بمرور واحد (O(n)) بدل .filter() لكل تاريخ داخل حلقة
  // (كان O(عدد الأيام × عدد الصفوف) وأصبح بطيئاً فعلياً مع نمو الفترة المطلوبة)
  var dailyByDate = {}, absByDate = {};
  daily.forEach(function (r) {
    var d = normalizeDate_(r.date);
    (dailyByDate[d] || (dailyByDate[d] = [])).push(r);
  });
  absences.forEach(function (r) {
    var d = normalizeDate_(r.date);
    (absByDate[d] || (absByDate[d] = [])).push(r);
  });
  var dates = Object.keys(Object.assign({}, dailyByDate, absByDate)).sort();

  var dailyTrend = dates.map(function (date) {
    var dRows = dailyByDate[date] || [];
    var aRows = absByDate[date] || [];
    return {
      date: date,
      present: dRows.length,
      absent: aRows.length,
      late: dRows.filter(function (r) { return Number(r.lateMinutes) > 0; }).length,
      early: dRows.filter(function (r) { return Number(r.earlyMinutes) > 0; }).length,
      noCheckout: dRows.filter(function (r) { return String(r.noCheckout) === 'true'; }).length,
    };
  });

  var latenessSeverity = { 'بسيط': 0, 'متوسط': 0, 'كبير': 0 };
  daily.forEach(function (r) {
    var c = classifyLate_(Number(r.lateMinutes));
    if (c) latenessSeverity[c]++;
  });

  var reasonCounts = {};
  absences.forEach(function (a) {
    var key = (a.reason && String(a.reason).trim())
      || (a.classification === 'بعذر' ? 'بعذر (بدون سبب محدد)' : a.classification === 'بدون عذر' ? 'بدون عذر' : 'غير مصنف');
    reasonCounts[key] = (reasonCounts[key] || 0) + 1;
  });
  var absenceReasons = Object.keys(reasonCounts)
    .map(function (label) { return { label: label, value: reasonCounts[label] }; })
    .sort(function (a, b) { return b.value - a.value; });
  if (absenceReasons.length > 8) {
    var head = absenceReasons.slice(0, 7);
    var restTotal = absenceReasons.slice(7).reduce(function (s, x) { return s + x.value; }, 0);
    absenceReasons = head.concat([{ label: 'أخرى', value: restTotal }]);
  }

  // نسبة الحضور العامة للفترة + مقارنة بفترة سابقة مساوية في الطول — لبطاقات المؤشرات
  // العلوية في شاشة التحليلات الجديدة (سهم أعلى/أدنى مقابل الفترة السابقة)
  var totalPresent = daily.length;
  var totalAbsent = absences.length;
  var attendanceRate = (totalPresent + totalAbsent) ? totalPresent / (totalPresent + totalAbsent) : null;

  var periodDays = daysBetweenISO_(from, to) + 1;
  var prevTo = addDaysISO_(from, -1);
  var prevFrom = addDaysISO_(from, -periodDays);
  var prevPresent = 0, prevAbsent = 0, prevRate = null;
  if (periodDays > 0 && periodDays <= 366) {
    prevPresent = getRowsForDateRange_(getSheet_(SHEET_ATTENDANCE), prevFrom, prevTo).length;
    prevAbsent = getRowsForDateRange_(getSheet_(SHEET_ABSENCE), prevFrom, prevTo).length;
    prevRate = (prevPresent + prevAbsent) ? prevPresent / (prevPresent + prevAbsent) : null;
  }

  return {
    ok: true, from: from, to: to, dailyTrend: dailyTrend, latenessSeverity: latenessSeverity, absenceReasons: absenceReasons,
    totalPresent: totalPresent, totalAbsent: totalAbsent, attendanceRate: attendanceRate,
    previousPeriod: { from: prevFrom, to: prevTo, totalPresent: prevPresent, totalAbsent: prevAbsent, attendanceRate: prevRate },
  };
  });
}

// سجل كل الاستيرادات مع عدد الحاضرين/الغائبين لكل تاريخ — لشاشة التقارير.
// منذ هذا التحديث يُخزَّن العدّادان مباشرة في صف الاستيراد نفسه وقت الحفظ (handleImportDay_)
// فلا حاجة لقراءة شيتي الحضور/الغياب الكاملين هنا؛ نلجأ لحسابهما فقط للاستيرادات
// القديمة السابقة لهذا التحديث والتي لا تحمل العدّاد بعد (توافق خلفي).
function handleListImports_() {
  return cached_('listImports', {}, function () {
  var imports = sheetToObjects_(getSheet_(SHEET_IMPORTS));
  var needsLegacyCounts = imports.some(function (r) { return r.presentCount === '' || r.presentCount === undefined || r.presentCount === null; });

  var presentCountByDate = {}, absentCountByDate = {};
  if (needsLegacyCounts) {
    var attendance = sheetToObjects_(getSheet_(SHEET_ATTENDANCE));
    var absence = sheetToObjects_(getSheet_(SHEET_ABSENCE));
    attendance.forEach(function (r) {
      var d = normalizeDate_(r.date);
      presentCountByDate[d] = (presentCountByDate[d] || 0) + 1;
    });
    absence.forEach(function (r) {
      var d = normalizeDate_(r.date);
      absentCountByDate[d] = (absentCountByDate[d] || 0) + 1;
    });
  }

  var list = imports.map(function (r) {
    var d = normalizeDate_(r.date);
    var hasStored = r.presentCount !== '' && r.presentCount !== undefined && r.presentCount !== null;
    return {
      date: d,
      importedAt: r.importedAt,
      fileName: r.fileName,
      schedStart: r.schedStart,
      schedEnd: r.schedEnd,
      schedLabel: r.schedLabel,
      presentCount: hasStored ? Number(r.presentCount) : (presentCountByDate[d] || 0),
      absentCount: hasStored ? Number(r.absentCount) : (absentCountByDate[d] || 0),
    };
  });
  list.sort(function (a, b) { return b.date.localeCompare(a.date); });
  return { ok: true, list: list };
  });
}

// ==================== سلّم إجراءات التأخير (مرفق ٦) ====================
// "بيان بعدد دقائق التأخير التي تم رصدها على الموظفين والإجراء المتخذ بشأنها" — الإدارة العامة
// للتعليم / إدارة الموارد البشرية. ثماني درجات متسلسلة بنصوصها كما وردت في النموذج:
//   step         رقم الدرجة (م) في النموذج
//   minMinutes   الحد التراكمي لدقائق التأخير
//   threshold    نص عمود "عدد دقائق التأخير" كما في النموذج
//   action       نوع الإجراء (يُستخدم أيضاً مفتاحاً لحفظ "تم/لم يتم" — لا تغيّر نصوص ١-٧
//                حتى لا تضيع حالات التنفيذ المسجّلة سابقاً)
//   actionText   الإجراء المتخذ (نص النموذج)
//   deductionDays أيام الحسم المعادلة التي ينص عليها النموذج عند هذه الدرجة
// الحسم يتبع النموذج حرفياً: ٤٢٠ د = يوم، ٨٤٠ د = يومان، ١٦٨٠ د = ٤ أيام (فأكثر: يوم لكل ٤٢٠ د).
// ما بين حدّين يبقى على حسم الدرجة الأدنى — فمثلاً ٦٩٠ دقيقة = يوم واحد لا يومان.
// الدرجة ٨ (repeat) لها نفس حد الدرجة ٧ لكنها للإحالة المتكررة: تُستحق فقط بعد تنفيذ الدرجة ٧
// واستمرار التأخر بعدها (تراكم دقائق جديدة بعد تاريخ الرفع الأول).
var EARLY_DEDUCT_ACTION = 'رفع حسم الانصراف المبكر';
var LATE_ACTION_TIERS = [
  { step: 1, minMinutes: 30, threshold: '٣٠ دقيقة', action: 'تنبيه شفوي',
    actionText: 'نظراً لتأخرك لذا وجب تنبيهك شفوياً', deductionDays: 0 },
  { step: 2, minMinutes: 60, threshold: '٦٠ دقيقة', action: 'تعهد خطي',
    actionText: 'نظراً لتأخرك لذا وجب تنبيهك خطياً', deductionDays: 0 },
  { step: 3, minMinutes: 120, threshold: '١٢٠ دقيقة', action: 'تعهد خطي (١)',
    actionText: 'أتعهد بعدم التأخر وإذا تكرر أتحمل ما يترتب على تأخيري من إجراءات', deductionDays: 0 },
  { step: 4, minMinutes: 240, threshold: '٢٤٠ دقيقة', action: 'تعهد خطي (٢)',
    actionText: 'أتعهد للمرة الثانية بعدم التأخر وإذا تكرر أتحمل ما يترتب على تأخيري من إجراءات', deductionDays: 0 },
  { step: 5, minMinutes: 420, threshold: '٤٢٠ دقيقة بما يعادل يوم', action: 'لفت نظر',
    actionText: 'نلفت نظرك لتكرار تأخرك علماً أنه سبق تنبيهك وأخذ تعهد عليك مع رفع الغياب للحسم', deductionDays: 1 },
  { step: 6, minMinutes: 840, threshold: '٨٤٠ دقيقة بما يعادل يومان', action: 'مساءلة ولفت نظر',
    actionText: 'مساءلة الموظف/ـة خطياً من قبل مدير/ة الجهة والرفع للمكتب المختص ليلفت نظره/ها مع رفع الغياب للحسم', deductionDays: 2 },
  { step: 7, minMinutes: 1680, threshold: '١٦٨٠ دقيقة بما يعادل ٤ أيام', action: 'الرفع لإدارة الموارد البشرية (وحدة متابعة دوام الموظفين)',
    actionText: 'الرفع لإدارة الموارد البشرية (وحدة متابعة دوام الموظفين) بما تم من إجراءات سابقة مع إرفاق صورة من مساءلة التأخير ورفع الغياب للحسم', deductionDays: 4 },
  { step: 8, minMinutes: 1680, threshold: '١٦٨٠ دقيقة بما يعادل ٤ أيام فأكثر', action: 'الرفع لإدارة الموارد البشرية مع إرفاق كافة الإجراءات',
    actionText: 'الرفع لإدارة الموارد البشرية (وحدة متابعة دوام الموظفين) مع إرفاق كافة الإجراءات التي سبق اتخاذها أعلاه مع إرفاق صورة من هذا النموذج ورفع الغياب للحسم', deductionDays: 4, repeat: true },
];

// أعلى درجة "بحد الدقائق" بلغها التراكم (الدرجات ١-٧؛ الدرجة ٨ تُحسم بسجل التنفيذ)
// بداية السنة المالية للتاريخ المعطى — أساس تراكم دقائق التأخير والانصراف المبكر وترحيلها
// حسب "الآليات والإجراءات المنظمة لتطبيق حضوري": من ٣١ ديسمبر إلى ٣٠ ديسمبر من السنة التالية.
function fiscalYearStart_(iso) {
  var y = Number(iso.slice(0, 4));
  var start = y + '-12-31';
  return iso >= start ? start : (y - 1) + '-12-31';
}
// إبقاء الاسم القديم متوافقاً
function schoolYearStart_(iso) { return fiscalYearStart_(iso); }

function lateActionTier_(cumulativeMinutes) {
  var best = null;
  LATE_ACTION_TIERS.forEach(function (t) {
    if (!t.repeat && cumulativeMinutes >= t.minMinutes) best = t;
  });
  return best;
}

// أيام الحسم: يوم عن كل ٤٢٠ دقيقة (٧ ساعات) أو مضاعفاتها — وما دون ذلك يُرحَّل ضمن السنة المالية
function lateDeductionDays_(cumulativeMinutes) {
  return Math.floor((Number(cumulativeMinutes) || 0) / 420);
}

function addDaysISO_(iso, delta) {
  var parts = iso.split('-').map(Number);
  var d = new Date(parts[0], parts[1] - 1, parts[2]);
  d.setDate(d.getDate() + delta);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function daysBetweenISO_(a, b) {
  var pa = a.split('-').map(Number), pb = b.split('-').map(Number);
  var da = new Date(pa[0], pa[1] - 1, pa[2]), db = new Date(pb[0], pb[1] - 1, pb[2]);
  return Math.round((db - da) / 86400000);
}

// يُرجع الحد التراكمي التالي الذي لم يُبلغه بعد (أعلى مباشرة من الحد الحالي)، أو null إن
// كان قد بلغ أعلى درجة أصلاً — يُستخدم لتنبيه "اقترب من درجة جزاء جديدة"
function nextLateTier_(cumulativeMinutes) {
  var next = null;
  for (var i = 0; i < LATE_ACTION_TIERS.length; i++) {
    var tier = LATE_ACTION_TIERS[i];
    if (!tier.repeat && tier.minMinutes > cumulativeMinutes) {
      if (!next || tier.minMinutes < next.minMinutes) next = tier;
    }
  }
  return next;
}

// يجمع كل ما يحتاج متابعة من المدير في نقطة واحدة — لبطاقة "بحاجة إلى متابعة" في لوحة اليوم:
// (١) يوم عمل بلا استيراد، (٢) غياب بلا تصنيف (آخر ٣٠ يوماً)، (٣) غياب متكرر بدون عذر (٣ أيام فأكثر، كامل السجل، يستثني "بعذر")،
// (٤) اقتراب معلّم من درجة جزاء تالية (خلال ٣٠ دقيقة)، (٥) بلا بصمة انصراف في آخر استيراد.
function handleGetAlerts_() {
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  return cached_('getAlerts', { today: today }, function () {
  var settings = getSettingsObj_();
  var teacherByCivil = {};
  sheetToObjects_(getSheet_(SHEET_TEACHERS)).filter(function (t) { return isActive_(t.active); })
    .forEach(function (t) { teacherByCivil[String(t.civil)] = t; });

  var imports = sheetToObjects_(getSheet_(SHEET_IMPORTS));
  var importDates = {};
  imports.forEach(function (r) { importDates[normalizeDate_(r.date)] = true; });

  // آخر يوم عمل (متجاوزاً العطل ونهاية الأسبوع) بحثاً عن استيراد ناقص — بحد أقصى أسبوعين للخلف
  var missingImportDate = null;
  for (var i = 0; i < 14; i++) {
    var d = addDaysISO_(today, -i);
    if (isNonWorkingDay_(d, settings)) continue;
    if (!importDates[d]) missingImportDate = d;
    break;
  }

  var from30 = addDaysISO_(today, -30);
  var absences30 = getRowsForDateRange_(getSheet_(SHEET_ABSENCE), from30, today);
  var unclassifiedAbsences = absences30
    .filter(function (r) { return !r.classification; })
    .map(function (r) { return { id: r.id, date: normalizeDate_(r.date), civil: String(r.civil), name: r.name }; })
    .sort(function (a, b) { return b.date.localeCompare(a.date); });

  // "غياب متكرر" يُحسب على كامل السجل المتوفر — لا نافذة الثلاثين يوماً — لأنه يمثّل نمطاً
  // تراكمياً يهمّ الإدارة معرفته بالكامل مهما امتدت الفترة (صيفي أو شتوي أو رمضان)، لا آخر
  // شهر فقط؛ معلّم غاب ١٣ يوماً منذ بداية الفصل يجب أن يظهر بـ١٣ هنا، لا يُقتطع منها ما سبق
  // آخر ٣٠ يوماً.
  var absencesAllTime = getRowsForDateRange_(getSheet_(SHEET_ABSENCE), '2000-01-01', today);
  var absenceCountByCivil = {};
  absencesAllTime.forEach(function (r) {
    // "بعذر" (إجازة مرضية موثّقة مثلاً) ليست نمطاً يستحق المتابعة الإدارية، حتى لو تكررت
    // أياماً كثيرة؛ فقط الغياب بدون عذر أو غير المصنَّف بعد (الذي قد يتحوّل لاحقاً لبدون
    // عذر) يُحسب هنا، وإلا ظهر معلّم بإجازة مرضية طويلة موثّقة كأنه "غياب متكرر بحاجة
    // متابعة" رغم التزامه الكامل بالإجراء الرسمي.
    if (r.classification === 'بعذر') return;
    var c = String(r.civil);
    absenceCountByCivil[c] = (absenceCountByCivil[c] || 0) + 1;
  });
  var repeatedAbsence = Object.keys(absenceCountByCivil)
    .filter(function (c) { return absenceCountByCivil[c] >= 3; })
    .map(function (c) {
      var t = teacherByCivil[c];
      return { civil: c, name: t ? t.name : '', num: t ? t.num : '', count: absenceCountByCivil[c] };
    })
    .sort(function (a, b) { return b.count - a.count; });

  // "اقترب من إجراء" يجب أن يُحسب على نفس أساس سلّم الإجراءات: التراكم منذ بداية العام
  // الدراسي — لا آخر ٣٠ يوماً — وإلا ظهر معلّم تجاوز الدرجة فعلاً على أنه "يقترب" منها.
  var yearFrom = schoolYearStart_(today);
  var dailyYear = getRowsForDateRange_(getSheet_(SHEET_ATTENDANCE), yearFrom, today);
  dailyYear = annotateWithPermissions_(dailyYear, yearFrom, today);
  var lateByCivil = {};
  dailyYear.forEach(function (r) {
    var mins = Number(r.lateMinutes) || 0;
    if (mins <= 0) return;
    var c = String(r.civil);
    lateByCivil[c] = (lateByCivil[c] || 0) + mins;
  });
  var approachingPenalty = Object.keys(lateByCivil)
    .map(function (c) {
      var total = lateByCivil[c];
      var next = nextLateTier_(total);
      if (!next || (next.minMinutes - total) > 30) return null;
      var t = teacherByCivil[c];
      return { civil: c, name: t ? t.name : '', num: t ? t.num : '', totalLateMinutes: total, remaining: next.minMinutes - total, nextAction: next.action };
    })
    .filter(Boolean)
    .sort(function (a, b) { return a.remaining - b.remaining; });

  var lastImportDate = null;
  imports.forEach(function (r) {
    var d = normalizeDate_(r.date);
    if (!lastImportDate || d > lastImportDate) lastImportDate = d;
  });
  var noCheckoutLast = [];
  if (lastImportDate) {
    noCheckoutLast = getRowsForDate_(getSheet_(SHEET_ATTENDANCE), lastImportDate)
      .filter(function (r) { return String(r.noCheckout) === 'true'; })
      .map(function (r) { return { civil: String(r.civil), name: r.name, num: r.num }; });
  }

  return {
    ok: true,
    today: today,
    missingImportDate: missingImportDate,
    unclassifiedAbsences: unclassifiedAbsences,
    repeatedAbsence: repeatedAbsence,
    interruptionCases: computeInterruptionCases_(today, settings, teacherByCivil),
    fieldPending: pendingFieldNotes_(today, teacherByCivil),
    approachingPenalty: approachingPenalty,
    noCheckoutLastImport: { date: lastImportDate, list: noCheckoutLast },
  };
  });
}

// كشف الجزاءات السنوية: تراكم دقائق التأخير لكل معلّم خلال فترة، والإجراء الرسمي المقابل
// (المرفق رقم ٦ الصادر عن الإدارة العامة للتعليم) + أيام الحسم المعادلة
var SHEET_ACTIONS = 'إجراءات_التأخير';
var SHEET_ACTIONS_HEADERS = ['key', 'civil', 'action', 'done', 'doneAt', 'minutes', 'updatedAt'];

// ينشئ الشيت تلقائياً بعناوينه إن لم يكن موجوداً بعد — لا حاجة لإعداد يدوي من المستخدم
function getOrCreateSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

// يسجّل هل نُفّذ الإجراء المطلوب لمعلّم معيّن أم لا — المفتاح هو (السجل المدني + نص
// الإجراء) وليس الفترة، لأن الإجراء يتبع الدرجة التي بلغها التراكم لا فترة تاريخ محددة:
// إن ارتفعت درجته لاحقاً يصبح إجراءً جديداً مختلف المفتاح يحتاج تنفيذاً من جديد.
// يسجّل "تم/لم يتم" لأي إجراء يُحدَّد بمفتاح — إما مفتاح تراكم درجة (السجل المدني + نص
// الإجراء، من كشف الاجراءات) أو مفتاح إشعار يوم واحد (من سجل الغياب/التأخر) يُمرَّر
// صراحةً في payload.key. القيمة نفسها تُقرأ لاحقاً عبر handleGetActionStatuses_.
function handleMarkActionDone_(payload) {
  var done = !!payload.done;
  var civil = payload.civil ? String(payload.civil) : '';
  var action = payload.tierAction ? String(payload.tierAction) : '';
  var key = payload.key ? String(payload.key) : (civil + '||' + action);
  var sheet = getOrCreateSheet_(SHEET_ACTIONS, SHEET_ACTIONS_HEADERS);
  ensureHeaders_(sheet, SHEET_ACTIONS_HEADERS); // شيتات قديمة بلا عمود minutes
  // minutes = تراكم دقائق التأخير لحظة التنفيذ — تحتاجه الدرجة ٨ لمعرفة هل استمر التأخر بعد الرفع
  var minutes = payload.minutes === undefined || payload.minutes === null || payload.minutes === '' ? '' : Number(payload.minutes);
  upsertRows_(sheet, 'key', [{
    key: key, civil: civil, action: action, done: done,
    doneAt: done ? (payload.doneAt ? String(payload.doneAt) : new Date().toISOString()) : '',
    minutes: done ? minutes : '', updatedAt: Date.now(),
  }]);
  bumpCacheGen_();
  return { ok: true };
}

// يُرجع كل حالات "تم/لم يتم" المسجَّلة دفعة واحدة (مفتاح ← {done, doneAt}) — تستخدمه
// شاشات سجل الغياب والتأخر لعرض عمود "الحالة" بجانب كل يوم دون طلب منفصل لكل صف.
function handleGetActionStatuses_() {
  return cached_('getActionStatuses', {}, function () {
    var rows = sheetToObjects_(getOrCreateSheet_(SHEET_ACTIONS, SHEET_ACTIONS_HEADERS));
    var statuses = {};
    rows.forEach(function (r) {
      statuses[r.key] = { done: String(r.done) === 'true', doneAt: r.doneAt || '' };
    });
    return { ok: true, statuses: statuses };
  });
}

function handleGetAnnualPenalties_(payload) {
  var from = String(payload.from);
  var to = String(payload.to);
  return cached_('getAnnualPenalties', { from: from, to: to }, function () {

  var teachers = sheetToObjects_(getSheet_(SHEET_TEACHERS)).filter(function (t) { return isActive_(t.active); });
  var daily = getRowsForDateRange_(getSheet_(SHEET_ATTENDANCE), from, to);
  daily = annotateWithPermissions_(daily, from, to);

  var lateByCivil = {};
  var earlyByCivil = {}; // الانصراف المبكر بدون إذن — يُجمع ويُحسم منفصلاً عن التأخير
  var earlyDaysByCivil = {};
  daily.forEach(function (r) {
    var civil = String(r.civil);
    var mins = Number(r.lateMinutes) || 0;
    if (mins > 0) lateByCivil[civil] = (lateByCivil[civil] || 0) + mins; // صافية بعد الاستئذانات
    var em = Number(r.earlyMinutes) || 0;
    if (em > 0) { earlyByCivil[civil] = (earlyByCivil[civil] || 0) + em; earlyDaysByCivil[civil] = (earlyDaysByCivil[civil] || 0) + 1; }
  });

  var doneMap = {};
  sheetToObjects_(getOrCreateSheet_(SHEET_ACTIONS, SHEET_ACTIONS_HEADERS)).forEach(function (r) {
    if (String(r.done) === 'true') doneMap[r.key] = { at: r.doneAt || '', minutes: r.minutes === '' || r.minutes == null ? null : Number(r.minutes) };
  });

  var list = teachers.map(function (t) {
    var civil = String(t.civil);
    var total = lateByCivil[civil] || 0;
    var step7 = LATE_ACTION_TIERS[6];
    var step7Done = doneMap[civil + '||' + step7.action];

    // كل درجات النموذج لهذا المعلّم: مبلوغة؟ منفّذة؟ — السلّم متسلسل، فكل درجة مبلوغة
    // لم تُنفَّذ تبقى مستحقة حتى لو تجاوزها التراكم (نص لفت النظر يفترض سبق التنبيه والتعهد).
    var steps = LATE_ACTION_TIERS.map(function (tier) {
      var key = civil + '||' + tier.action;
      var done = doneMap[key];
      var reached;
      if (tier.repeat) {
        // الإحالة المتكررة: بعد تنفيذ الدرجة ٧ واستمرار التأخر بعدها
        var base = step7Done ? (step7Done.minutes != null ? step7Done.minutes : step7.minMinutes) : null;
        reached = !!done || (base != null && total > base);
      } else {
        reached = total >= tier.minMinutes || !!done;
      }
      return {
        step: tier.step, action: tier.action, actionText: tier.actionText, threshold: tier.threshold,
        minMinutes: tier.minMinutes, deductionDays: tier.deductionDays,
        reached: reached, done: !!done, doneAt: done ? done.at : '', doneMinutes: done ? done.minutes : null,
      };
    });
    var reachedSteps = steps.filter(function (x) { return x.reached; });
    if (!reachedSteps.length) return null; // لم يبلغ ٣٠ دقيقة
    var current = reachedSteps[reachedSteps.length - 1];            // أعلى درجة مبلوغة
    var pending = reachedSteps.filter(function (x) { return !x.done; });
    var next = pending.length ? pending[0] : null;                  // الإجراء التالي بالتسلسل

    var deductionDays = lateDeductionDays_(total);
    var raisedDays = 0; // أيام رُفعت للحسم فعلاً مع الدرجات المنفّذة
    steps.forEach(function (x) {
      if (!x.done || !x.deductionDays) return;
      var d = x.doneMinutes != null ? Math.floor(x.doneMinutes / 420) : x.deductionDays;
      if (d > raisedDays) raisedDays = d;
    });

    return {
      civil: civil, num: t.num, name: t.name,
      totalLateMinutes: total,
      step: current.step,
      action: current.action,
      actionText: current.actionText,
      note: current.threshold,
      steps: steps,
      pendingCount: pending.length,
      nextStep: next ? next.step : null,
      nextAction: next ? next.action : '',
      deductionDays: deductionDays,
      raisedDays: Math.min(raisedDays, deductionDays),
      daysToRaise: Math.max(0, deductionDays - raisedDays),
      actionDone: pending.length === 0,
      actionDoneAt: current.done ? current.doneAt : '',
    };
  }).filter(Boolean);

  list.sort(function (a, b) { return b.totalLateMinutes - a.totalLateMinutes; });

  // الانصراف المبكر بدون إذن: تراكم وحسم منفصلان — يوم عن كل ٤٢٠ دقيقة، والباقي يُرحَّل
  // عدم التواجد أثناء الدوام (خروج بلا بصمة ثم عودة) من سجل المتابعة الميدانية = خروج بدون إذن،
  // فيُضاف لنفس وعاء الانصراف المبكر ويُحسم معه كل ٤٢٠ دقيقة
  var outByCivil = {}, outCountByCivil = {};
  fieldNotesInRange_(from, to).forEach(function (n) {
    if (n.kind !== 'absent_during') return;
    var m = Number(n.minutes) || 0;
    if (m <= 0) return;
    var c = String(n.civil);
    outByCivil[c] = (outByCivil[c] || 0) + m;
    outCountByCivil[c] = (outCountByCivil[c] || 0) + 1;
    earlyByCivil[c] = (earlyByCivil[c] || 0) + m;
  });
  var tnByCivil = {};
  teachers.forEach(function (t) { tnByCivil[String(t.civil)] = t; });
  var earlyList = Object.keys(earlyByCivil).filter(function (c) { return tnByCivil[c]; }).map(function (c) {
    var t = tnByCivil[c];
    var total = earlyByCivil[c];
    var done = doneMap[c + '||' + EARLY_DEDUCT_ACTION];
    var raised = done && done.minutes != null ? Math.floor(done.minutes / 420) : 0;
    var ded = Math.floor(total / 420);
    return {
      civil: c, num: t.num, name: t.name,
      totalEarlyMinutes: total, earlyDays: earlyDaysByCivil[c] || 0,
      outMinutes: outByCivil[c] || 0, outCount: outCountByCivil[c] || 0,
      fingerprintEarlyMinutes: total - (outByCivil[c] || 0),
      deductionDays: ded, raisedDays: Math.min(raised, ded), daysToRaise: Math.max(0, ded - raised),
      toNextDeduction: 420 - (total % 420),
      raisedAt: done ? done.at : '',
    };
  }).sort(function (a, b) { return b.totalEarlyMinutes - a.totalEarlyMinutes; });

  return { ok: true, from: from, to: to, list: list, earlyList: earlyList };
  });
}

function handleGetEmployeeReport_(payload) {
  var from = String(payload.from);
  var to = String(payload.to);
  var civilFilter = payload.civil ? String(payload.civil) : '';

  var daily = getRowsForDateRange_(getSheet_(SHEET_ATTENDANCE), from, to).filter(function (r) {
    return !civilFilter || String(r.civil) === civilFilter;
  });
  daily = annotateWithPermissions_(daily, from, to);
  var absences = getRowsForDateRange_(getSheet_(SHEET_ABSENCE), from, to).filter(function (r) {
    return !civilFilter || String(r.civil) === civilFilter;
  });

  var absenceEvents = absences.map(function (a) {
    return {
      date: normalizeDate_(a.date), civil: String(a.civil), num: a.num, name: a.name,
      classification: a.classification, reason: a.reason, note: a.note,
    };
  }).sort(function (a, b) { return a.date.localeCompare(b.date); });

  var lateEvents = daily.filter(function (r) { return Number(r.lateMinutes) > 0; }).map(function (r) {
    return { date: normalizeDate_(r.date), civil: String(r.civil), num: r.num, name: r.name, checkIn: r.checkIn, lateMinutes: Number(r.lateMinutes) };
  }).sort(function (a, b) { return a.date.localeCompare(b.date); });

  var earlyEvents = daily.filter(function (r) { return Number(r.earlyMinutes) > 0; }).map(function (r) {
    return { date: normalizeDate_(r.date), civil: String(r.civil), num: r.num, name: r.name, checkOut: r.checkOut, earlyMinutes: Number(r.earlyMinutes) };
  }).sort(function (a, b) { return a.date.localeCompare(b.date); });

  var totals = {
    presentDays: daily.length,
    absenceDays: absenceEvents.length,
    lateDays: lateEvents.length,
    lateMinutesTotal: lateEvents.reduce(function (s, e) { return s + e.lateMinutes; }, 0),
    earlyDays: earlyEvents.length,
    earlyMinutesTotal: earlyEvents.reduce(function (s, e) { return s + e.earlyMinutes; }, 0),
  };

  return {
    ok: true, from: from, to: to, civil: civilFilter,
    absenceEvents: absenceEvents, lateEvents: lateEvents, earlyEvents: earlyEvents, totals: totals,
  };
}

function handleUpdateAbsence_(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var id = String(payload.id);
    var sheet = getSheet_(SHEET_ABSENCE);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idCol = headers.indexOf('id');
    var classificationCol = headers.indexOf('classification');
    var reasonCol = headers.indexOf('reason');
    var noteCol = headers.indexOf('note');
    var updatedAtCol = headers.indexOf('updatedAt');

    for (var i = 1; i < data.length; i++) {
      if (normalizeDate_(data[i][idCol]) === id) {
        var row = i + 1;
        // نص عادي إجبارياً قبل الكتابة (حماية من حقن الصيغ) — انظر تعليق FORCE_TEXT_COLUMNS
        sheet.getRange(row, classificationCol + 1).setNumberFormat('@').setValue(payload.classification || '');
        sheet.getRange(row, reasonCol + 1).setNumberFormat('@').setValue(payload.reason || '');
        sheet.getRange(row, noteCol + 1).setNumberFormat('@').setValue(payload.note || '');
        sheet.getRange(row, updatedAtCol + 1).setValue(Date.now());
        return { ok: true, id: id };
      }
    }
    return { ok: false, error: 'لم يُعثر على سجل غياب بهذا المعرّف: ' + id };
  } finally {
    bumpCacheGen_();
    lock.releaseLock();
  }
}

// تصنيف جماعي لعدة سجلات غياب دفعة واحدة، بطلب واحد للخادم بدل طلب منفصل لكل سجل.
// كانت الواجهة تُرسل طلباً مستقلاً لكل سجل محدَّد بالتتابع (حتى ٣٦+ طلباً متتالياً في
// حالات حقيقية)، وكل طلب Apps Script له كلفة شبكة وبدء تشغيل مستقلة — فكان يفشل غالباً
// بمنتصف الطريق بخطأ "Failed to fetch" (انقطاع شبكي أو تقييد من جوجل لتكرار الطلبات
// السريعة) تاركاً بعض السجلات محدَّثة وبعضها لا. هذا الإجراء يقرأ الشيت مرة واحدة
// ويحدّث كل السجلات المطلوبة ضمن تنفيذ واحد فيتجنّب المشكلة كلياً وأسرع بكثير.
function handleBulkUpdateAbsence_(payload) {
  var updates = payload.updates || [];
  if (!updates.length) return { ok: true, updated: 0 };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEET_ABSENCE);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idCol = headers.indexOf('id');
    var classificationCol = headers.indexOf('classification');
    var reasonCol = headers.indexOf('reason');
    var noteCol = headers.indexOf('note');
    var updatedAtCol = headers.indexOf('updatedAt');
    var now = Date.now();

    var byId = {};
    updates.forEach(function (u) { byId[String(u.id)] = u; });
    var remaining = updates.length;
    var updated = 0;

    for (var i = 1; i < data.length && remaining > 0; i++) {
      var rowId = normalizeDate_(data[i][idCol]);
      var u = byId[rowId];
      if (!u) continue;
      var row = i + 1;
      // نص عادي إجبارياً قبل الكتابة (حماية من حقن الصيغ) — انظر تعليق FORCE_TEXT_COLUMNS
      sheet.getRange(row, classificationCol + 1).setNumberFormat('@').setValue(u.classification || '');
      sheet.getRange(row, reasonCol + 1).setNumberFormat('@').setValue(u.reason || '');
      sheet.getRange(row, noteCol + 1).setNumberFormat('@').setValue(u.note || '');
      sheet.getRange(row, updatedAtCol + 1).setValue(now);
      updated++;
      remaining--;
    }
    return { ok: true, updated: updated };
  } finally {
    bumpCacheGen_();
    lock.releaseLock();
  }
}

/* ---------------------------------------------------------- */
/* الإعدادات: فترات الدوام المتعددة + تقويم العطل             */
/* ---------------------------------------------------------- */

function getSettingsObj_() {
  var sheet = getSheet_(SHEET_SETTINGS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === 'app_settings') {
      try {
        var parsed = JSON.parse(data[i][1]);
        return Object.assign({}, DEFAULT_SETTINGS, parsed);
      } catch (e) {
        return DEFAULT_SETTINGS;
      }
    }
  }
  return DEFAULT_SETTINGS;
}

// يكشف أي فترتي دوام متداخلتين بالتاريخ قبل الحفظ — نفس منطق الواجهة، كخط دفاع ثانٍ
// يمنع حالة بيانات فاسدة تُفسد حساب دوام أي يوم ضمن التداخل حتى لو تم تجاوز تحقق الواجهة.
function findOverlappingSchedulePeriods_(periods) {
  var dated = (periods || []).filter(function (p) { return p.dateFrom && p.dateTo; });
  for (var i = 0; i < dated.length; i++) {
    for (var j = i + 1; j < dated.length; j++) {
      var a = dated[i], b = dated[j];
      if (a.dateFrom <= b.dateTo && b.dateFrom <= a.dateTo) return [a, b];
    }
  }
  return null;
}

function handleSaveSettings_(payload) {
  var settings = payload.settings || {};
  var overlap = findOverlappingSchedulePeriods_(settings.schedulePeriods);
  if (overlap) {
    return { ok: false, error: 'الفترتان "' + (overlap[0].label || overlap[0].id) + '" و"' + (overlap[1].label || overlap[1].id) + '" متداخلتان بالتاريخ.' };
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEET_SETTINGS);
    var data = sheet.getDataRange().getValues();
    var targetRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === 'app_settings') { targetRow = i + 1; break; }
    }
    var row = ['app_settings', JSON.stringify(settings), Date.now()];
    if (targetRow === -1) {
      sheet.appendRow(row);
    } else {
      sheet.getRange(targetRow, 1, 1, 3).setValues([row]);
    }
    return { ok: true, settings: settings };
  } finally {
    bumpCacheGen_();
    lock.releaseLock();
  }
}

// ==================== أداة إعادة الحساب: تصحيح سجلات حضور استُوردت بجدول دوام خاطئ ====================
// معاينة فقط — لا تكتب شيئاً. لكل صف حضور بالفترة، نعيد حساب دقائق التأخير/الانصراف المبكر
// باستخدام جدول الدوام "الصحيح" وفق الإعدادات الحالية (لا الجدول المخزَّن على الصف نفسه، الذي
// قد يكون خاطئاً لو استُورد بقيمة افتراضية قديمة قبل تصحيح الإعدادات). نُرجع فقط الصفوف التي
// يختلف فيها الجدول أو الدقائق المحسوبة فعلياً عمّا هو مخزَّن، حتى تُراجَع قبل أي تعديل.
function handlePreviewRecalculate_(payload) {
  var from = String(payload.from);
  var to = String(payload.to);
  var settings = getSettingsObj_();
  var rows = getRowsForDateRange_(getSheet_(SHEET_ATTENDANCE), from, to);
  var changes = [];
  rows.forEach(function (r) {
    var date = normalizeDate_(r.date);
    var correct = resolveScheduleForDate_(date, settings);
    var checkInMin = parseTimeToMinutes_(r.checkIn);
    var checkOutMin = r.checkOut ? parseTimeToMinutes_(r.checkOut) : null;
    var schedStartMin = parseTimeToMinutes_(correct.workStart);
    var schedEndMin = parseTimeToMinutes_(correct.workEnd);
    var newLate = checkInMin != null && schedStartMin != null ? Math.max(0, checkInMin - schedStartMin) : 0;
    var newEarly = checkOutMin != null && schedEndMin != null ? Math.max(0, schedEndMin - checkOutMin) : 0;
    var oldLate = Number(r.lateMinutes) || 0;
    var oldEarly = Number(r.earlyMinutes) || 0;
    if (newLate !== oldLate || newEarly !== oldEarly) {
      // نُبلِّغ فقط عن الصفوف التي يتغيّر فيها الرقم الفعلي المحسوب (تأخير/انصراف مبكر) —
      // لا كل صف تغيّر فيه نص جدول الدوام المخزَّن بلا أي أثر عملي على أي رقم، وإلا امتلأت
      // المعاينة بمئات الصفوف التي لا تستحق مراجعة أصلاً وطغت على الحالات المهمة فعلاً.
      changes.push({
        id: r.id, date: date, civil: String(r.civil), name: r.name,
        oldSchedStart: r.schedStart || '', newSchedStart: correct.workStart || '',
        oldSchedEnd: r.schedEnd || '', newSchedEnd: correct.workEnd || '',
        oldLateMinutes: oldLate, newLateMinutes: newLate,
        oldEarlyMinutes: oldEarly, newEarlyMinutes: newEarly,
      });
    }
  });
  return { ok: true, totalScanned: rows.length, changes: changes };
}

// تطبيق فعلي — يكتب بالضبط ما أرسله العميل (نتيجة المعاينة، كاملة أو جزء مختار منها)، بلا
// إعادة حساب من جديد هنا، حتى يتطابق ما يراه المستخدم بالمعاينة مع ما يُكتب فعلياً بالشيت.
function handleApplyRecalculate_(payload) {
  var changes = payload.changes || [];
  if (!changes.length) return { ok: true, applied: 0 };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEET_ATTENDANCE);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idCol = headers.indexOf('id');
    var schedStartCol = headers.indexOf('schedStart');
    var schedEndCol = headers.indexOf('schedEnd');
    var lateCol = headers.indexOf('lateMinutes');
    var earlyCol = headers.indexOf('earlyMinutes');
    var idToRow = {};
    for (var i = 1; i < data.length; i++) idToRow[String(data[i][idCol])] = i + 1;
    var applied = 0;
    changes.forEach(function (c) {
      var row = idToRow[String(c.id)];
      if (!row) return;
      sheet.getRange(row, schedStartCol + 1).setNumberFormat('@').setValue(c.newSchedStart);
      sheet.getRange(row, schedEndCol + 1).setNumberFormat('@').setValue(c.newSchedEnd);
      sheet.getRange(row, lateCol + 1).setValue(c.newLateMinutes);
      sheet.getRange(row, earlyCol + 1).setValue(c.newEarlyMinutes);
      applied++;
    });
    return { ok: true, applied: applied };
  } finally {
    bumpCacheGen_();
    lock.releaseLock();
  }
}

// يطابق تاريخاً بفترة الدوام السارية عليه (آخر تطابق بمدى تاريخي يتفوّق، يسمح باستثناء)،
// وإلا يستخدم الفترة الافتراضية المحدَّدة في الإعدادات
function resolveScheduleForDate_(dateISO, settings) {
  var periods = settings.schedulePeriods || [];
  var match = null;
  for (var i = 0; i < periods.length; i++) {
    var p = periods[i];
    if (!p.dateFrom || !p.dateTo) continue;
    if (dateISO >= p.dateFrom && dateISO <= p.dateTo) match = p;
  }
  if (!match) {
    var defId = settings.defaultPeriodId || '';
    match = periods.filter(function (p) { return p.id === defId; })[0] || periods[0] || DEFAULT_SCHEDULE_PERIODS[0];
  }
  return { workStart: match.workStart, workEnd: match.workEnd, label: match.label || '', id: match.id || '' };
}

function isNonWorkingDay_(dateISO, settings) {
  var holidays = settings.holidays || [];
  for (var i = 0; i < holidays.length; i++) {
    var h = holidays[i];
    if (h.from && h.to && dateISO >= h.from && dateISO <= h.to) return true;
  }
  var weekend = settings.weekendDays || DEFAULT_SETTINGS.weekendDays;
  var wd = isoToWeekday_(dateISO);
  return weekend.indexOf(wd) !== -1;
}

function isoToWeekday_(dateISO) {
  var parts = dateISO.split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]).getDay();
}

// ==================== تنزيل نسخة كاملة للعمل المحلي (بدون مزامنة) ====================
// تُستخدم عند إيقاف المزامنة في الواجهة: تنزّل الواجهة نسخة من كل الشيتات وتشغّل نفس هذا
// الكود داخل المتصفح عليها. القيم تُنسَّق بنفس formatCellForOutput_ حتى تصل التواريخ والأوقات
// نصوصاً كما يراها التطبيق عادةً، لا كائنات Date.
function handleExportAll_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var names = [SHEET_TEACHERS, SHEET_ATTENDANCE, SHEET_ABSENCE, SHEET_IMPORTS, SHEET_SETTINGS, SHEET_PERMISSIONS, SHEET_ACTIONS, SHEET_USERS, SHEET_FIELD];
  var sheets = {};
  names.forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var values = sheet.getDataRange().getValues();
    var headers = values[0] || [];
    // أعمدة كلمات المرور/الاستعادة لا تُنسخ إلى المتصفح إطلاقاً
    var secretCols = name === SHEET_USERS ? { salt: 1, hash: 1, resetHash: 1, resetExpires: 1, resetTries: 1 } : {};
    sheets[name] = values.map(function (row, i) {
      if (i === 0) return row;
      return row.map(function (v, j) { return secretCols[headers[j]] ? '' : formatCellForOutput_(headers[j], v); });
    });
  });
  return { ok: true, sheets: sheets, exportedAt: Date.now() };
}

// ==================== إدارة المعلمين من التطبيق ====================
// كل المعلمين (النشطون وغير النشطين) — لشاشة إدارة المعلمين بالإعدادات.
function handleGetAllTeachers_() {
  var rows = sheetToObjects_(getSheet_(SHEET_TEACHERS))
    .filter(function (t) { return String(t.civil || '').trim(); })
    .map(function (t) {
      return { civil: String(t.civil).trim(), num: t.num, name: t.name, active: isActive_(t.active),
        role: staffRole_(t), startDate: t.startDate ? normalizeDate_(t.startDate) : '' };
    });
  rows.sort(function (a, b) {
    var na = Number(a.num), nb = Number(b.num);
    if (!isNaN(na) && !isNaN(nb) && a.num !== '' && b.num !== '') return na - nb;
    return String(a.name || '').localeCompare(String(b.name || ''), 'ar');
  });
  return { ok: true, teachers: rows };
}

// إضافة/تعديل معلّم أو أكثر (مطابقة بالسجل المدني). تُدمج القيم مع الصف الموجود حتى لا
// تُمسح أي أعمدة إضافية يضيفها المستخدم يدوياً في شيت المعلمين.
function handleSaveTeachers_(payload) {
  var list = payload.teachers || [];
  var clean = [];
  var skipped = 0;
  list.forEach(function (t) {
    var civil = String(t.civil == null ? '' : t.civil).trim();
    var name = String(t.name == null ? '' : t.name).trim();
    if (!civil || !name) { skipped++; return; }
    var numRaw = t.num == null ? '' : String(t.num).trim();
    var num = numRaw !== '' && !isNaN(Number(numRaw)) ? Number(numRaw) : numRaw;
    var active = t.active === undefined ? true : isActive_(t.active);
    var rec = { civil: civil, name: name, num: num, active: active };
    if (t.role !== undefined) rec.role = String(t.role).trim() === 'إداري' ? 'إداري' : 'معلم';
    if (t.startDate !== undefined) {
      var sd = String(t.startDate || '').trim();
      rec.startDate = /^\d{4}-\d{2}-\d{2}$/.test(sd) ? sd : '';
    }
    clean.push(rec);
  });
  if (!clean.length) return { ok: true, saved: 0, skipped: skipped };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEET_TEACHERS);
    ensureHeaders_(sheet, defaultHeadersFor_(SHEET_TEACHERS));
    var existing = {};
    sheetToObjects_(sheet).forEach(function (r) { existing[String(r.civil).trim()] = r; });
    var records = clean.map(function (t) {
      return Object.assign({}, existing[t.civil] || {}, t);
    });
    upsertRows_(sheet, 'civil', records);
    return { ok: true, saved: records.length, skipped: skipped };
  } finally {
    CacheService.getScriptCache().remove('teachers_v1');
    bumpCacheGen_();
    lock.releaseLock();
  }
}

// حذف معلّم من الروستر. سجلات حضوره وغيابه السابقة تبقى كما هي (تحمل اسمه).
function handleDeleteTeacher_(payload) {
  var civil = String(payload.civil || '').trim();
  if (!civil) return { ok: false, error: 'السجل المدني مطلوب' };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEET_TEACHERS);
    var data = sheet.getDataRange().getValues();
    var idCol = data[0].indexOf('civil');
    var deleted = 0;
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(normalizeDate_(data[i][idCol])).trim() === civil) { sheet.deleteRow(i + 1); deleted++; }
    }
    return { ok: true, deleted: deleted };
  } finally {
    CacheService.getScriptCache().remove('teachers_v1');
    bumpCacheGen_();
    lock.releaseLock();
  }
}

// ==================== الحسابات وتسجيل الدخول ====================
// كلمات المرور لا تُحفظ نصاً أبداً: تُحفظ بصمة SHA-256 مملّحة ومكرّرة. "نسيت كلمة المرور" لا
// يرسل كلمة المرور (لا يمكن استرجاعها أصلاً) بل رمز تحقق من ٦ أرقام صالح ١٥ دقيقة يُرسل للبريد.
// أول حساب يُنشأ يصبح "مديراً" مفعّلاً؛ أي تسجيل بعده يبقى "بانتظار التفعيل" حتى يفعّله مدير —
// حتى لا يستطيع أي شخص يصل للصفحة إنشاء حساب والدخول لبيانات المدرسة.
var SHEET_USERS = 'المستخدمون';
var SHEET_USERS_HEADERS = ['username', 'email', 'salt', 'hash', 'role', 'status', 'createdAt', 'lastLoginAt', 'resetHash', 'resetExpires', 'resetTries'];
var SESSION_DAYS = 7;

function sha256Hex_(str) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { var v = (b < 0 ? b + 256 : b).toString(16); return v.length === 1 ? '0' + v : v; }).join('');
}
function hashPassword_(password, salt) {
  var h = salt + '|' + password;
  for (var i = 0; i < 300; i++) h = sha256Hex_(h + '|' + salt);
  return h;
}
function randomToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}
// سرّ توقيع الجلسات — يُنشأ مرة ويُحفظ في شيت الإعدادات (يعمل كذلك في الوضع المحلي)
function authSecret_() {
  var sheet = getSheet_(SHEET_SETTINGS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) if (data[i][0] === 'auth_secret' && data[i][1]) return String(data[i][1]);
  var secret = randomToken_();
  sheet.appendRow(['auth_secret', secret, Date.now()]);
  return secret;
}
function usersSheet_() {
  var sh = getOrCreateSheet_(SHEET_USERS, SHEET_USERS_HEADERS);
  ensureHeaders_(sh, SHEET_USERS_HEADERS);
  return sh;
}
function normUser_(u) { return String(u || '').trim().toLowerCase(); }
function findUser_(key) {
  var k = normUser_(key);
  if (!k) return null;
  var list = sheetToObjects_(usersSheet_());
  for (var i = 0; i < list.length; i++) {
    if (normUser_(list[i].username) === k || normUser_(list[i].email) === k) return list[i];
  }
  return null;
}
function saveUser_(u) { upsertRows_(usersSheet_(), 'username', [u]); }
function publicUser_(u) { return { username: u.username, email: u.email, role: u.role || 'user', status: u.status || 'active' }; }

function makeSession_(username) {
  var exp = Date.now() + SESSION_DAYS * 86400000;
  var payload = username + '|' + exp;
  return payload + '|' + sha256Hex_(authSecret_() + '|' + payload);
}
function sessionUser_(session) {
  var parts = String(session || '').split('|');
  if (parts.length !== 3) return null;
  var payload = parts[0] + '|' + parts[1];
  if (sha256Hex_(authSecret_() + '|' + payload) !== parts[2]) return null;
  if (Number(parts[1]) < Date.now()) return null;
  var u = findUser_(parts[0]);
  if (!u || u.status !== 'active') return null;
  return u;
}

function handleAuthInfo_() {
  return { ok: true, hasUsers: sheetToObjects_(usersSheet_()).length > 0 };
}

function validEmail_(e) { return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(e || '').trim()); }

function handleRegister_(payload) {
  var username = String(payload.username || '').trim();
  var email = String(payload.email || '').trim();
  var password = String(payload.password || '');
  if (!/^[\u0600-\u06FFA-Za-z0-9_.\- ]{3,40}$/.test(username)) return { ok: false, error: 'اسم المستخدم يجب أن يكون من ٣ إلى ٤٠ حرفاً (حروف أو أرقام).' };
  if (!validEmail_(email)) return { ok: false, error: 'البريد الإلكتروني غير صحيح.' };
  if (password.length < 6) return { ok: false, error: 'كلمة المرور يجب ألا تقل عن ٦ أحرف.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (findUser_(username)) return { ok: false, error: 'اسم المستخدم مستخدم مسبقاً.' };
    if (findUser_(email)) return { ok: false, error: 'هذا البريد مسجّل بحساب آخر.' };
    var first = sheetToObjects_(usersSheet_()).length === 0;
    var salt = randomToken_().slice(0, 24);
    var user = {
      username: username, email: email, salt: salt, hash: hashPassword_(password, salt),
      role: first ? 'admin' : 'user', status: first ? 'active' : 'pending',
      createdAt: new Date().toISOString(), lastLoginAt: '', resetHash: '', resetExpires: '',
    };
    saveUser_(user);
    if (first) return { ok: true, pending: false, session: makeSession_(username), user: publicUser_(user) };
    return { ok: true, pending: true, message: 'تم إنشاء الحساب، وهو بانتظار تفعيله من مدير النظام.' };
  } finally {
    lock.releaseLock();
  }
}

function handleLogin_(payload) {
  var u = findUser_(payload.username);
  if (!u || hashPassword_(String(payload.password || ''), u.salt) !== u.hash) {
    return { ok: false, error: 'اسم المستخدم أو كلمة المرور غير صحيحة.' };
  }
  if (u.status === 'pending') return { ok: false, error: 'الحساب بانتظار التفعيل من مدير النظام.' };
  if (u.status !== 'active') return { ok: false, error: 'هذا الحساب موقوف.' };
  u.lastLoginAt = new Date().toISOString();
  saveUser_(u);
  return { ok: true, session: makeSession_(u.username), user: publicUser_(u) };
}

function handleVerifySession_(payload) {
  var u = sessionUser_(payload.session);
  if (!u) return { ok: false, error: 'انتهت الجلسة — سجّل الدخول من جديد.' };
  return { ok: true, user: publicUser_(u) };
}

function handleForgotPassword_(payload) {
  var u = findUser_(payload.username);
  // الرد نفسه سواء وُجد الحساب أم لا — حتى لا تُستخدم الشاشة لمعرفة الحسابات المسجّلة
  var generic = { ok: true, message: 'إن كان الحساب مسجّلاً فسيصل رمز التحقق إلى بريده خلال دقائق.' };
  if (!u || !u.email) return generic;
  var code = String(Math.floor(100000 + Math.random() * 900000));
  u.resetHash = sha256Hex_(u.salt + '|reset|' + code);
  u.resetExpires = Date.now() + 15 * 60000;
  u.resetTries = 0;
  saveUser_(u);
  try {
  MailApp.sendEmail({
    to: u.email,
    subject: 'رمز استعادة كلمة المرور — متابعة الحضور',
    htmlBody: '<div dir="rtl" style="font-family:Tahoma,Arial;font-size:15px;line-height:1.9">'
      + 'مرحباً ' + u.username + '،<br>طلبتَ استعادة كلمة المرور لنظام متابعة الحضور.<br>'
      + 'رمز التحقق: <b style="font-size:22px;letter-spacing:4px">' + code + '</b><br>'
      + 'الرمز صالح لمدة ١٥ دقيقة. إن لم تطلب ذلك فتجاهل هذه الرسالة.</div>',
  });
  } catch (mailErr) {
    return { ok: false, error: String(mailErr && mailErr.message || mailErr).replace(/^Error:\s*/, '') };
  }
  return generic;
}

function handleResetPassword_(payload) {
  var u = findUser_(payload.username);
  var code = String(payload.code || '').trim();
  if (!u || !u.resetHash || !u.resetExpires || Number(u.resetExpires) < Date.now()) {
    return { ok: false, error: 'رمز التحقق غير صحيح أو منتهي الصلاحية.' };
  }
  if (sha256Hex_(u.salt + '|reset|' + code) !== u.resetHash) {
    // ٥ محاولات خاطئة تُبطل الرمز — يمنع تخمين الرمز المكوّن من ٦ أرقام
    u.resetTries = (Number(u.resetTries) || 0) + 1;
    if (u.resetTries >= 5) { u.resetHash = ''; u.resetExpires = ''; }
    saveUser_(u);
    return { ok: false, error: u.resetTries >= 5 ? 'تجاوزت عدد المحاولات — اطلب رمزاً جديداً.' : 'رمز التحقق غير صحيح أو منتهي الصلاحية.' };
  }
  var pw = String(payload.newPassword || '');
  if (pw.length < 6) return { ok: false, error: 'كلمة المرور يجب ألا تقل عن ٦ أحرف.' };
  u.salt = randomToken_().slice(0, 24);
  u.hash = hashPassword_(pw, u.salt);
  u.resetHash = ''; u.resetExpires = '';
  saveUser_(u);
  return { ok: true, message: 'تم تغيير كلمة المرور، يمكنك الدخول الآن.' };
}

function handleChangePassword_(payload) {
  var u = sessionUser_(payload.session);
  if (!u) return { ok: false, error: 'انتهت الجلسة — سجّل الدخول من جديد.' };
  if (hashPassword_(String(payload.oldPassword || ''), u.salt) !== u.hash) return { ok: false, error: 'كلمة المرور الحالية غير صحيحة.' };
  var pw = String(payload.newPassword || '');
  if (pw.length < 6) return { ok: false, error: 'كلمة المرور الجديدة يجب ألا تقل عن ٦ أحرف.' };
  u.salt = randomToken_().slice(0, 24);
  u.hash = hashPassword_(pw, u.salt);
  if (payload.email !== undefined && validEmail_(payload.email)) u.email = String(payload.email).trim();
  saveUser_(u);
  return { ok: true };
}

function handleListUsers_(payload) {
  var me = sessionUser_(payload.session);
  if (!me) return { ok: false, error: 'انتهت الجلسة — سجّل الدخول من جديد.' };
  var list = sheetToObjects_(usersSheet_()).map(function (u) {
    var p = publicUser_(u); p.createdAt = u.createdAt || ''; p.lastLoginAt = u.lastLoginAt || ''; return p;
  });
  return { ok: true, me: publicUser_(me), users: me.role === 'admin' ? list : [publicUser_(me)] };
}

// المدير فقط: تفعيل/إيقاف/ترقية/حذف حساب (ولا يستطيع إيقاف أو حذف نفسه)
function handleUpdateUser_(payload) {
  var me = sessionUser_(payload.session);
  if (!me) return { ok: false, error: 'انتهت الجلسة — سجّل الدخول من جديد.' };
  if (me.role !== 'admin') return { ok: false, error: 'هذا الإجراء للمدير فقط.' };
  var u = findUser_(payload.username);
  if (!u) return { ok: false, error: 'الحساب غير موجود.' };
  if (normUser_(u.username) === normUser_(me.username)) return { ok: false, error: 'لا يمكنك تعديل حسابك من هنا.' };
  var sheet = usersSheet_();
  if (payload.op === 'delete') {
    var data = sheet.getDataRange().getValues();
    var col = data[0].indexOf('username');
    for (var i = data.length - 1; i >= 1; i--) if (normUser_(data[i][col]) === normUser_(u.username)) sheet.deleteRow(i + 1);
    return { ok: true };
  }
  if (payload.op === 'activate') u.status = 'active';
  else if (payload.op === 'suspend') u.status = 'suspended';
  else if (payload.op === 'makeAdmin') u.role = 'admin';
  else if (payload.op === 'makeUser') u.role = 'user';
  else if (payload.op === 'setPassword') {
    var np = String(payload.newPassword || '');
    if (np.length < 6) return { ok: false, error: 'كلمة المرور يجب ألا تقل عن ٦ أحرف.' };
    u.salt = randomToken_().slice(0, 24);
    u.hash = hashPassword_(np, u.salt);
    u.resetHash = ''; u.resetExpires = '';
  }
  else return { ok: false, error: 'إجراء غير معروف' };
  saveUser_(u);
  return { ok: true };
}

// ==================== محضر الانقطاع عن العمل (نموذج ٩) ====================
// الآليات المنظمة (خامساً/٤): يُعدّ محضر انقطاع إذا بلغ غياب الموظف بدون عذر ١٥ يوماً متصلة أو
// ٣٠ يوماً متفرقة خلال السنة. ومع ملحوظة مساءلة الغياب: غياب قبل عطلة (نهاية أسبوع/إجازة) وبعدها
// تُحتسب العطلة ضمن مدة الغياب. غير المصنّف يُعامل كبدون عذر حتى يُصنَّف (مع التنبيه لذلك).
function absenceStreaks_(dates, settings) {
  // dates: تواريخ غياب مرتبة؛ تُدمج في فترات إن فصلت بينها أيام عطلة فقط
  var streaks = [];
  dates.forEach(function (d) {
    var last = streaks[streaks.length - 1];
    if (last) {
      var gap = [];
      var x = addDaysISO_(last.to, 1);
      var bridged = true;
      while (x < d) {
        if (!isNonWorkingDay_(x, settings)) { bridged = false; break; }
        gap.push(x);
        x = addDaysISO_(x, 1);
      }
      if (bridged) { last.to = d; last.days += gap.length + 1; last.bridged += gap.length; return; }
    }
    streaks.push({ from: d, to: d, days: 1, bridged: 0 });
  });
  return streaks;
}
function computeInterruptionCases_(today, settings, teacherByCivil) {
  var from = fiscalYearStart_(today);
  var byCivil = {};
  getRowsForDateRange_(getSheet_(SHEET_ABSENCE), from, today).forEach(function (a) {
    var c = String(a.civil);
    if (!teacherByCivil[c]) return;
    var cls = String(a.classification || '');
    if (cls === 'بعذر') return;
    (byCivil[c] || (byCivil[c] = { dates: [], unclassified: 0 })).dates.push(normalizeDate_(a.date));
    if (!cls) byCivil[c].unclassified++;
  });
  var out = [];
  Object.keys(byCivil).forEach(function (c) {
    var info = byCivil[c];
    var dates = info.dates.sort();
    var streaks = absenceStreaks_(dates, settings);
    var longest = streaks.reduce(function (b, s) { return s.days > b.days ? s : b; }, { days: 0 });
    var totalWithBridges = streaks.reduce(function (a, s) { return a + s.days; }, 0);
    var reached = longest.days >= 15 ? 'consecutive' : totalWithBridges >= 30 ? 'scattered' : '';
    var near = !reached && (longest.days >= 10 || totalWithBridges >= 20);
    if (!reached && !near) return;
    var t = teacherByCivil[c];
    out.push({
      civil: c, name: t.name, num: t.num,
      longestStreak: longest.days, longestFrom: longest.from || '', longestTo: longest.to || '',
      totalDays: totalWithBridges, absenceDays: dates.length, unclassified: info.unclassified,
      reached: reached, near: near, dates: dates,
    });
  });
  out.sort(function (a, b) { return (b.reached ? 1 : 0) - (a.reached ? 1 : 0) || b.totalDays - a.totalDays; });
  return out;
}

// ==================== سجل المتابعة الميدانية ====================
// حالات لا تكشفها البصمة ويثبتها المدير/الوكيل بالمتابعة: عدم التواجد أثناء الدوام (خروج دون
// بصمة ثم عودة) والحضور/الانصراف من موقع آخر. كل ملاحظة موثّقة بمن سجّلها ووقت التسجيل والشاهد.
var SHEET_FIELD = 'المتابعة الميدانية';
var SHEET_FIELD_HEADERS = ['id', 'date', 'civil', 'name', 'kind', 'fromTime', 'toTime', 'minutes', 'punchKind', 'location',
  'verifiedBy', 'witness', 'note', 'createdBy', 'createdAt', 'updatedAt'];
function fieldSheet_() {
  var sh = getOrCreateSheet_(SHEET_FIELD, SHEET_FIELD_HEADERS);
  ensureHeaders_(sh, SHEET_FIELD_HEADERS);
  return sh;
}
function fieldNotesInRange_(from, to) {
  return getRowsForDateRange_(fieldSheet_(), from, to);
}
function handleGetFieldNotes_(payload) {
  var from = String(payload.from || '0000-01-01'), to = String(payload.to || '9999-12-31');
  var rows = fieldNotesInRange_(from, to).map(function (r) {
    r.date = normalizeDate_(r.date);
    r.minutes = Number(r.minutes) || 0;
    return r;
  });
  rows.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)) || String(b.createdAt).localeCompare(String(a.createdAt)); });
  return { ok: true, rows: rows };
}
function handleSaveFieldNote_(payload) {
  var n = payload.note || {};
  var date = String(n.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: false, error: 'التاريخ غير صحيح.' };
  var civil = String(n.civil || '').trim();
  if (!civil) return { ok: false, error: 'اختر الموظف.' };
  var kind = n.kind === 'other_location' ? 'other_location' : 'absent_during';
  var minutes = 0;
  if (kind === 'absent_during') {
    var a = parseTimeToMinutes_(n.fromTime), b = parseTimeToMinutes_(n.toTime);
    if (a == null || b == null || b <= a) return { ok: false, error: 'حدّد وقت الخروج ووقت العودة (العودة بعد الخروج).' };
    minutes = b - a;
  }
  var teacher = sheetToObjects_(getSheet_(SHEET_TEACHERS)).filter(function (t) { return String(t.civil) === civil; })[0];
  var now = new Date().toISOString();
  var rec = {
    id: n.id || ('fn_' + Date.now().toString(36) + Math.floor(Math.random() * 1e6).toString(36)),
    date: date, civil: civil, name: teacher ? teacher.name : String(n.name || ''), kind: kind,
    fromTime: kind === 'absent_during' ? String(n.fromTime || '') : '',
    toTime: kind === 'absent_during' ? String(n.toTime || '') : '',
    minutes: minutes,
    punchKind: kind === 'other_location' ? (n.punchKind === 'انصراف' ? 'انصراف' : n.punchKind === 'حضور وانصراف' ? 'حضور وانصراف' : 'حضور') : '',
    location: kind === 'other_location' ? String(n.location || '') : '',
    verifiedBy: String(n.verifiedBy || ''), witness: String(n.witness || ''), note: String(n.note || ''),
    createdBy: String(n.createdBy || ''), createdAt: n.createdAt || now, updatedAt: Date.now(),
  };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    upsertRows_(fieldSheet_(), 'id', [rec]);
  } finally {
    bumpCacheGen_();
    lock.releaseLock();
  }
  return { ok: true, note: rec };
}
function handleDeleteFieldNote_(payload) {
  var id = String(payload.id || '');
  if (!id) return { ok: false, error: 'معرّف غير صحيح' };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    deleteRowsByIds_(fieldSheet_(), [id]);
  } finally {
    bumpCacheGen_();
    lock.releaseLock();
  }
  return { ok: true };
}

function pendingFieldNotes_(today, teacherByCivil) {
  var done = {};
  sheetToObjects_(getOrCreateSheet_(SHEET_ACTIONS, SHEET_ACTIONS_HEADERS)).forEach(function (r) {
    if (String(r.done) === 'true') done[r.key] = true;
  });
  return fieldNotesInRange_(fiscalYearStart_(today), today).filter(function (n) {
    var d = normalizeDate_(n.date);
    return teacherByCivil[String(n.civil)] && !done['day::field::' + n.civil + '::' + d + '#' + n.id];
  }).map(function (n) {
    return { id: n.id, civil: String(n.civil), name: n.name, date: normalizeDate_(n.date), kind: n.kind, minutes: Number(n.minutes) || 0 };
  });
}

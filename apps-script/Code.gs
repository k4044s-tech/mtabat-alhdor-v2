/**
 * Apps Script Backend لتطبيق "متابعة الحضور" (V2)
 * يُلصق هذا الكود في: الشيت "متابعة الحضور V29" ← Extensions ← Apps Script
 * بعد اللصق: Project Settings ← Script Properties ← أضف SHARED_SECRET بأي قيمة سرية،
 * ثم Deploy ← New deployment ← Web app ← Execute as: Me ← Who has access: Anyone.
 */

var SHEET_TEACHERS = 'المعلمون';
var SHEET_ATTENDANCE = 'الحضور';
var SHEET_ABSENCE = 'الغياب';
var SHEET_IMPORTS = 'الاستيرادات';

function doPost(e) {
  var payload = JSON.parse(e.postData.contents);
  return handleRequest_(payload);
}

function doGet(e) {
  var payload = {};
  for (var k in e.parameter) payload[k] = e.parameter[k];
  if (payload.rows) payload.rows = JSON.parse(payload.rows);
  return handleRequest_(payload);
}

function handleRequest_(payload) {
  var out;
  try {
    var secret = PropertiesService.getScriptProperties().getProperty('SHARED_SECRET');
    if (secret && payload.token !== secret) {
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
        case 'deleteDay':
          out = handleDeleteDay_(payload);
          break;
        case 'getAbsenceRange':
          out = handleGetAbsenceRange_(payload);
          break;
        case 'updateAbsence':
          out = handleUpdateAbsence_(payload);
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

function getSheet_(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('sheet not found: ' + name);
  return sheet;
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

function normalizeDate_(v) {
  if (isDateValue_(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}

function isActive_(v) {
  return v === true || String(v).toLowerCase() === 'true';
}

// يقرأ "H:MM" أو "HH:MM AM/PM" ويرجّع عدد الدقائق من منتصف الليل، أو null إن تعذّر
function parseTimeToMinutes_(s) {
  if (!s) return null;
  var m = String(s).trim().match(/^(\d{1,2}):(\d{2})\s*(AM|PM|am|pm)?$/);
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
var FORCE_TEXT_COLUMNS = {
  date: 1, checkIn: 1, checkOut: 1, hoursRaw: 1,
  schedStart: 1, schedEnd: 1, schedLabel: 1, schedId: 1,
};

// upsert بمطابقة عمود المعرّف (idColName) — يحدّث الصف الموجود أو يضيف صفاً جديداً
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

// يحذف كل صفوف تاريخ معيّن من الحضور/الغياب/الاستيرادات — لتصحيح استيراد خاطئ
function handleDeleteDay_(payload) {
  var date = String(payload.date);
  [SHEET_ATTENDANCE, SHEET_ABSENCE].forEach(function (name) {
    var sheet = getSheet_(name);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var dateCol = headers.indexOf('date');
    for (var i = data.length - 1; i >= 1; i--) {
      if (normalizeDate_(data[i][dateCol]) === date) sheet.deleteRow(i + 1);
    }
  });
  var importsSheet = getSheet_(SHEET_IMPORTS);
  var idata = importsSheet.getDataRange().getValues();
  var iheaders = idata[0];
  var idateCol = iheaders.indexOf('date');
  for (var i = idata.length - 1; i >= 1; i--) {
    if (normalizeDate_(idata[i][idateCol]) === date) importsSheet.deleteRow(i + 1);
  }
  return { ok: true, date: date };
}

function handleGetTeachers_() {
  var teachers = sheetToObjects_(getSheet_(SHEET_TEACHERS)).filter(function (t) {
    return isActive_(t.active);
  });
  return { ok: true, teachers: teachers };
}

function handleGetDay_(payload) {
  var date = String(payload.date);
  var attendance = sheetToObjects_(getSheet_(SHEET_ATTENDANCE)).filter(function (r) {
    return normalizeDate_(r.date) === date;
  });
  var absence = sheetToObjects_(getSheet_(SHEET_ABSENCE)).filter(function (r) {
    return normalizeDate_(r.date) === date;
  });
  var importInfo = sheetToObjects_(getSheet_(SHEET_IMPORTS)).filter(function (r) {
    return normalizeDate_(r.date) === date;
  })[0] || null;
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

    var schedStartMin = parseTimeToMinutes_(schedStart);
    var schedEndMin = parseTimeToMinutes_(schedEnd);
    var now = Date.now();
    var presentCivils = {};
    var attendanceRecords = [];

    rows.forEach(function (r) {
      var civil = String(r.civil);
      var t = teacherByCivil[civil];
      if (!t) return; // تجاهل من ليس في روستر المعلمين
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

    var absenceRecords = [];
    teachers.forEach(function (t) {
      if (!isActive_(t.active)) return;
      var civil = String(t.civil);
      if (presentCivils[civil]) return;
      absenceRecords.push({
        id: date + '__' + civil,
        date: date,
        civil: civil,
        num: t.num,
        name: t.name,
        classification: '',
        reason: '',
        note: '',
        updatedAt: now,
      });
    });

    upsertRows_(getSheet_(SHEET_ATTENDANCE), 'id', attendanceRecords);
    upsertRows_(getSheet_(SHEET_ABSENCE), 'id', absenceRecords);
    upsertRows_(getSheet_(SHEET_IMPORTS), 'date', [
      {
        date: date,
        importedAt: new Date().toISOString(),
        fileName: payload.fileName || '',
        schedStart: schedStart,
        schedEnd: schedEnd,
        schedLabel: schedLabel,
        updatedAt: now,
      },
    ]);

    return {
      ok: true,
      date: date,
      presentCount: attendanceRecords.length,
      absentCount: absenceRecords.length,
    };
  } finally {
    lock.releaseLock();
  }
}

function handleGetAbsenceRange_(payload) {
  var from = String(payload.from);
  var to = String(payload.to);
  var rows = sheetToObjects_(getSheet_(SHEET_ABSENCE)).filter(function (r) {
    var d = normalizeDate_(r.date);
    return d >= from && d <= to;
  });
  rows.sort(function (a, b) {
    return normalizeDate_(a.date) < normalizeDate_(b.date) ? -1 : 1;
  });
  return { ok: true, from: from, to: to, rows: rows };
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
        sheet.getRange(row, classificationCol + 1).setValue(payload.classification || '');
        sheet.getRange(row, reasonCol + 1).setValue(payload.reason || '');
        sheet.getRange(row, noteCol + 1).setValue(payload.note || '');
        sheet.getRange(row, updatedAtCol + 1).setValue(Date.now());
        return { ok: true, id: id };
      }
    }
    return { ok: false, error: 'لم يُعثر على سجل غياب بهذا المعرّف: ' + id };
  } finally {
    lock.releaseLock();
  }
}

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
};
var WEEKDAY_AR_ = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];

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

// أسرع من sheetToObjects_ للبحث عن تاريخ واحد فقط: يقرأ عمود التاريخ وحده أولاً (رخيص)
// بدل كل الأعمدة، ثم يقرأ فقط النطاق المحصور بين أول وآخر صف مطابق. يفيد كثيراً مع
// نمو الشيت لأن sheetToObjects_ تقرأ كل الصفوف والأعمدة في كل استدعاء.
function getRowsForDate_(sheet, date) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var dateCol = headers.indexOf('date') + 1;
  if (dateCol === 0) return sheetToObjects_(sheet).filter(function (r) { return normalizeDate_(r.date) === date; });

  var dateValues = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues();
  var matchRows = [];
  for (var i = 0; i < dateValues.length; i++) {
    if (normalizeDate_(dateValues[i][0]) === date) matchRows.push(i + 2);
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

    var skipAbsenceMarking = isNonWorkingDay_(date, getSettingsObj_());
    var absenceRecords = [];
    teachers.forEach(function (t) {
      if (!isActive_(t.active)) return;
      var civil = String(t.civil);
      if (presentCivils[civil]) return;
      if (skipAbsenceMarking) return; // يوم عطلة/نهاية أسبوع: لا نُسجّل غياباً لمن لم يظهر بالملف
      var existing = existingAbsenceByCivil[civil];
      absenceRecords.push({
        id: date + '__' + civil,
        date: date,
        civil: civil,
        num: t.num,
        name: t.name,
        classification: existing ? existing.classification : '',
        reason: existing ? existing.reason : '',
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

// ملخص شهري لكل معلّم مرتّب تنازلياً حسب "المخالفات" (غياب + تأخير + انصراف مبكر) — لشاشة الملخص والترتيب
function handleGetMonthlySummary_(payload) {
  var from = String(payload.from);
  var to = String(payload.to);

  var teachers = sheetToObjects_(getSheet_(SHEET_TEACHERS));
  var daily = sheetToObjects_(getSheet_(SHEET_ATTENDANCE)).filter(function (r) {
    var d = normalizeDate_(r.date);
    return d >= from && d <= to;
  });
  var absences = sheetToObjects_(getSheet_(SHEET_ABSENCE)).filter(function (r) {
    var d = normalizeDate_(r.date);
    return d >= from && d <= to;
  });
  var imports = sheetToObjects_(getSheet_(SHEET_IMPORTS)).filter(function (r) {
    var d = normalizeDate_(r.date);
    return d >= from && d <= to;
  });

  function blankEntry(civil, num, name, active) {
    return {
      civil: civil, num: num, name: name, active: active,
      presentDays: 0, absenceDays: 0, excused: 0, unexcused: 0, unclassified: 0,
      lateDays: 0, lateMinutes: 0, earlyDays: 0, earlyMinutes: 0, noCheckoutDays: 0,
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
    if (late > 0) { e.lateDays++; e.lateMinutes += late; }
    if (early > 0) { e.earlyDays++; e.earlyMinutes += early; }
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
  list.forEach(function (e) { e.violations = e.absenceDays + e.lateDays + e.earlyDays; });
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
    return acc;
  }, { absenceDays: 0, excused: 0, unexcused: 0, unclassified: 0, lateDays: 0, lateMinutes: 0, earlyDays: 0, earlyMinutes: 0, noCheckoutDays: 0 });

  return { ok: true, from: from, to: to, list: list, totals: totals, workDays: imports.length };
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

  var daily = sheetToObjects_(getSheet_(SHEET_ATTENDANCE)).filter(function (r) {
    var d = normalizeDate_(r.date);
    return d >= from && d <= to;
  });
  var absences = sheetToObjects_(getSheet_(SHEET_ABSENCE)).filter(function (r) {
    var d = normalizeDate_(r.date);
    return d >= from && d <= to;
  });

  var dateSet = {};
  daily.forEach(function (r) { dateSet[normalizeDate_(r.date)] = true; });
  absences.forEach(function (r) { dateSet[normalizeDate_(r.date)] = true; });
  var dates = Object.keys(dateSet).sort();

  var dailyTrend = dates.map(function (date) {
    var dRows = daily.filter(function (r) { return normalizeDate_(r.date) === date; });
    var aRows = absences.filter(function (r) { return normalizeDate_(r.date) === date; });
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

  return { ok: true, from: from, to: to, dailyTrend: dailyTrend, latenessSeverity: latenessSeverity, absenceReasons: absenceReasons };
}

// سجل كل الاستيرادات مع عدد الحاضرين/الغائبين لكل تاريخ — لشاشة التقارير
function handleListImports_() {
  var imports = sheetToObjects_(getSheet_(SHEET_IMPORTS));
  var attendance = sheetToObjects_(getSheet_(SHEET_ATTENDANCE));
  var absence = sheetToObjects_(getSheet_(SHEET_ABSENCE));

  var presentCountByDate = {};
  attendance.forEach(function (r) {
    var d = normalizeDate_(r.date);
    presentCountByDate[d] = (presentCountByDate[d] || 0) + 1;
  });
  var absentCountByDate = {};
  absence.forEach(function (r) {
    var d = normalizeDate_(r.date);
    absentCountByDate[d] = (absentCountByDate[d] || 0) + 1;
  });

  var list = imports.map(function (r) {
    var d = normalizeDate_(r.date);
    return {
      date: d,
      importedAt: r.importedAt,
      fileName: r.fileName,
      schedStart: r.schedStart,
      schedEnd: r.schedEnd,
      schedLabel: r.schedLabel,
      presentCount: presentCountByDate[d] || 0,
      absentCount: absentCountByDate[d] || 0,
    };
  });
  list.sort(function (a, b) { return b.date.localeCompare(a.date); });
  return { ok: true, list: list };
}

// تقرير الغياب والتأخر والانصراف المبكر — لموظف واحد (civil) أو لجميع الموظفين (civil فارغ)
// جدول التدرج الرسمي (مرفق ٦ — بيان بعدد دقائق التأخير) لوزارة التعليم:
// كل حد تراكمي لدقائق التأخير خلال العام يقابله إجراء تصعيدي، ومن حد ٤٢٠ دقيقة
// فأعلى يعادل عدد أيام غياب تُحسم (بمعدّل يوم واحد لكل ٤٢٠ دقيقة تراكمية).
var LATE_ACTION_TIERS = [
  { minMinutes: 1680, action: 'الرفع لإدارة الموارد البشرية (وحدة متابعة دوام الموظفين)', note: 'مع إرفاق كافة الإجراءات السابقة ورفع الغياب للحسم' },
  { minMinutes: 840, action: 'مساءلة ولفت نظر', note: 'يعادل يومان للحسم' },
  { minMinutes: 420, action: 'لفت نظر', note: 'يعادل يوم للحسم' },
  { minMinutes: 240, action: 'تعهد خطي (٢)', note: 'للمرة الثانية' },
  { minMinutes: 120, action: 'تعهد خطي (١)', note: '' },
  { minMinutes: 60, action: 'تعهد خطي', note: '' },
  { minMinutes: 30, action: 'تنبيه شفوي', note: '' },
];

function lateActionTier_(cumulativeMinutes) {
  for (var i = 0; i < LATE_ACTION_TIERS.length; i++) {
    if (cumulativeMinutes >= LATE_ACTION_TIERS[i].minMinutes) return LATE_ACTION_TIERS[i];
  }
  return null;
}

// كشف الجزاءات السنوية: تراكم دقائق التأخير لكل معلّم خلال فترة، والإجراء الرسمي المقابل
// (المرفق رقم ٦ الصادر عن الإدارة العامة للتعليم) + أيام الحسم المعادلة
function handleGetAnnualPenalties_(payload) {
  var from = String(payload.from);
  var to = String(payload.to);

  var teachers = sheetToObjects_(getSheet_(SHEET_TEACHERS)).filter(function (t) { return isActive_(t.active); });
  var daily = sheetToObjects_(getSheet_(SHEET_ATTENDANCE)).filter(function (r) {
    var d = normalizeDate_(r.date);
    return d >= from && d <= to;
  });

  var lateByCivil = {};
  daily.forEach(function (r) {
    var civil = String(r.civil);
    var mins = Number(r.lateMinutes) || 0;
    if (mins <= 0) return;
    lateByCivil[civil] = (lateByCivil[civil] || 0) + mins;
  });

  var list = teachers.map(function (t) {
    var civil = String(t.civil);
    var totalLateMinutes = lateByCivil[civil] || 0;
    var tier = lateActionTier_(totalLateMinutes);
    return {
      civil: civil, num: t.num, name: t.name,
      totalLateMinutes: totalLateMinutes,
      action: tier ? tier.action : '—',
      note: tier ? tier.note : '',
      deductionDays: totalLateMinutes >= 420 ? Math.round(totalLateMinutes / 420) : 0,
    };
  }).filter(function (e) { return e.totalLateMinutes > 0; });

  list.sort(function (a, b) { return b.totalLateMinutes - a.totalLateMinutes; });

  return { ok: true, from: from, to: to, list: list };
}

function handleGetEmployeeReport_(payload) {
  var from = String(payload.from);
  var to = String(payload.to);
  var civilFilter = payload.civil ? String(payload.civil) : '';

  var daily = sheetToObjects_(getSheet_(SHEET_ATTENDANCE)).filter(function (r) {
    var d = normalizeDate_(r.date);
    if (d < from || d > to) return false;
    if (civilFilter && String(r.civil) !== civilFilter) return false;
    return true;
  });
  var absences = sheetToObjects_(getSheet_(SHEET_ABSENCE)).filter(function (r) {
    var d = normalizeDate_(r.date);
    if (d < from || d > to) return false;
    if (civilFilter && String(r.civil) !== civilFilter) return false;
    return true;
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

function handleSaveSettings_(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var settings = payload.settings || {};
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

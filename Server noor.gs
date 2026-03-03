// ═══════════════════════════════════════════════════════════════
// Server_Noor.gs — واجهة برمجية لربط نور (v3 — كل المراحل + التأخر + التعويضية)
// يستخدم: SPREADSHEET_URL, getSheetName_(), findSheet_(), SHEET_REGISTRY من Config.gs
// ═══════════════════════════════════════════════════════════════

/**
 * جلب السجلات المعلقة التي لم تُرسل لنور بعد
 * @param {string} stage - المرحلة (ابتدائي/متوسط/ثانوي/طفولة مبكرة)
 * @param {string} type  - النوع: violations | tardiness | compensation | excellent | absence | all
 * @return {Object} { success, records[], stats{}, total }
 */
function getNoorPendingRecords(stage, type, filterMode) {
  try {
    filterMode = filterMode || 'today';
    var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    var result = {
      records: [],
      stats: { violations: 0, tardiness: 0, compensation: 0, excellent: 0, absence: 0 }
    };

    // ═══ 1. المخالفات السلوكية ═══
    if (type === 'violations' || type === 'all') {
      var vRecords = getNoorViolationRecords_(ss, stage, filterMode);
      result.records = result.records.concat(vRecords);
      result.stats.violations = vRecords.length;
    }

    // ═══ 2. التأخر الصباحي ═══
    if (type === 'tardiness' || type === 'all') {
      var tRecords = getNoorTardinessRecords_(ss, stage, filterMode);
      result.records = result.records.concat(tRecords);
      result.stats.tardiness = tRecords.length;
    }

    // ═══ 3. السلوك الإيجابي (تعويضية + متمايز) ═══
    if (type === 'compensation' || type === 'excellent' || type === 'positive' || type === 'all') {
      var pRecords = getNoorPositiveRecords_(ss, stage, filterMode);

      if (type === 'compensation') {
        var compRecords = pRecords.filter(function(r) { return r._type === 'compensation'; });
        result.records = result.records.concat(compRecords);
        result.stats.compensation = compRecords.length;
      } else if (type === 'excellent') {
        var excRecords = pRecords.filter(function(r) { return r._type === 'excellent'; });
        result.records = result.records.concat(excRecords);
        result.stats.excellent = excRecords.length;
      } else {
        // all أو positive
        result.records = result.records.concat(pRecords);
        pRecords.forEach(function(r) {
          if (r._type === 'compensation') result.stats.compensation++;
          else result.stats.excellent++;
        });
      }
    }

    // ═══ 4. الغياب اليومي ═══
    if (type === 'absence' || type === 'all') {
      var aRecords = getNoorAbsenceRecords_(ss, stage, filterMode);
      result.records = result.records.concat(aRecords);
      result.stats.absence = aRecords.length;
    }

    // ★ تنظيف البيانات قبل الإرسال للعميل
    // google.script.run لا يستطيع إرسال كائنات Date — تحويلها لنصوص
    result.records = sanitizeRecordsForClient_(result.records);

    result.success = true;
    result.total = result.records.length;
    return result;

  } catch (e) {
    return {
      success: false,
      error: e.message,
      records: [],
      stats: { violations: 0, tardiness: 0, compensation: 0, excellent: 0, absence: 0 },
      total: 0
    };
  }
}

// ═══════════════════════════════════════════════════════════════
// دوال جلب السجلات الفرعية (private)
// ═══════════════════════════════════════════════════════════════

/**
 * جلب المخالفات المعلقة — فلتر: حالة_نور فارغة/معلق
 * filterMode: 'today' = اليوم فقط، 'all' = كل غير الموثق
 */
function getNoorViolationRecords_(ss, stage, filterMode) {
  var sheetName = getSheetName_('المخالفات', stage);
  var sheet = findSheet_(ss, sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var records = getSheetAsRecords_(sheet);
  var cutoff = (filterMode === 'today') ? getDateCutoff_(0) : null;

  return records.filter(function(r) {
    var noorStatus = String(r['حالة_نور'] || '').trim();
    if (noorStatus !== '' && noorStatus !== 'معلق') return false;
    // فلتر التاريخ حسب الوضع
    if (cutoff) {
      var dateStr = String(r['التاريخ_الميلادي'] || r['التاريخ الميلادي'] || '');
      if (dateStr) {
        var recDate = parseDateSafe_(dateStr);
        if (recDate && recDate < cutoff) return false;
      }
    }
    return true;
  }).map(function(r) {
    r._type = 'violation';
    r._noorMode = NOOR_DROPDOWN_MAP.modes.violation;

    // ربط المخالفة بنور عبر رقم المخالفة
    var violId = parseInt(r['رقم المخالفة'] || r['رقم_المخالفة'] || '0');
    var mapping = NOOR_DROPDOWN_MAP.violations[violId];
    if (mapping) {
      r._noorValue = mapping.noorValue;
      r._noorText = mapping.noorText;
    } else {
      // بحث بالنص كخطة بديلة
      var violText = String(r['نص المخالفة'] || r['نص_المخالفة'] || '').trim();
      var textMapping = findNoorViolationByText_(violText);
      if (textMapping) {
        r._noorValue = textMapping.noorValue;
        r._noorText = textMapping.noorText;
      }
    }
    return r;
  });
}

/**
 * جلب سجلات التأخر المعلقة — فلتر: حالة_نور فارغة/معلق
 * filterMode: 'today' = اليوم فقط، 'all' = كل غير الموثق
 * التأخر يُدخل في نور كمخالفة سلوكية: التأخر الصباحي (الدرجة الأولى)
 */
function getNoorTardinessRecords_(ss, stage, filterMode) {
  var sheetName = getSheetName_('التأخر', stage);
  var sheet = findSheet_(ss, sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var records = getSheetAsRecords_(sheet);
  var cutoff = (filterMode === 'today') ? getDateCutoff_(0) : null;

  return records.filter(function(r) {
    var noorStatus = String(r['حالة_نور'] || '').trim();
    if (noorStatus !== '' && noorStatus !== 'معلق') return false;
    // فلتر التاريخ حسب الوضع
    if (cutoff) {
      var dateStr = String(r['وقت_الإدخال'] || r['وقت الإدخال'] || '');
      if (dateStr) {
        var recDate = parseDateSafe_(dateStr);
        if (recDate && recDate < cutoff) return false;
      }
    }
    return true;
  }).map(function(r) {
    r._type = 'tardiness';
    r._noorValue = '1601174,الدرجة الأولى'; // التأخر الصباحي دائماً
    r._noorMode = NOOR_DROPDOWN_MAP.modes.tardiness;
    r._noorText = 'التأخر الصباحي.';
    return r;
  });
}

/**
 * جلب السلوك الإيجابي المعلق — يُقسم إلى: تعويضية / متمايز
 * filterMode: 'today' = اليوم فقط، 'all' = كل غير الموثق
 */
function getNoorPositiveRecords_(ss, stage, filterMode) {
  var sheetName = getSheetName_('السلوك_الإيجابي', stage);
  var sheet = findSheet_(ss, sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var records = getSheetAsRecords_(sheet);
  var cutoff = (filterMode === 'today') ? getDateCutoff_(0) : null;

  return records.filter(function(r) {
    var noorStatus = String(r['حالة_نور'] || '').trim();
    if (noorStatus !== '' && noorStatus !== 'معلق') return false;
    // فلتر التاريخ حسب الوضع
    if (cutoff) {
      var dateStr = String(r['التاريخ_الميلادي'] || r['التاريخ الميلادي'] || r['وقت_الإدخال'] || r['وقت الإدخال'] || '');
      if (dateStr) {
        var recDate = parseDateSafe_(dateStr);
        if (recDate && recDate < cutoff) return false;
      }
    }
    return true;
  }).map(function(r) {
    var behaviorName = String(r['السلوك_المتمايز'] || r['السلوك المتمايز'] || '').trim();
    // التمييز بين التعويضية والمتمايز
    var degree = String(r['الدرجة'] || '');
    r._noorMode = NOOR_DROPDOWN_MAP.modes.compensation; // نفس الوضع للنوعين

    if (behaviorName.indexOf('تعويض') >= 0 || degree.indexOf('تعويض') >= 0) {
      r._type = 'compensation';
      var mapping = findNoorMapping_(NOOR_DROPDOWN_MAP.compensation, behaviorName);
      if (mapping) { r._noorValue = mapping.noorValue; r._noorText = mapping.noorText; }
    } else {
      r._type = 'excellent';
      var mapping = findNoorMapping_(NOOR_DROPDOWN_MAP.excellent, behaviorName);
      if (mapping) { r._noorValue = mapping.noorValue; r._noorText = mapping.noorText; }
    }
    return r;
  });
}

/**
 * جلب الغياب اليومي المعلق — فلتر: حالة_نور فارغة/معلق
 * filterMode: 'today' = اليوم فقط، 'all' = كل غير الموثق
 */
function getNoorAbsenceRecords_(ss, stage, filterMode) {
  var sheetName = getSheetName_('الغياب_اليومي', stage);
  var sheet = findSheet_(ss, sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var records = getSheetAsRecords_(sheet);

  // فلتر التاريخ حسب الوضع
  var today = (filterMode === 'today') ? getTodayHijriDate_() : null;

  return records.filter(function(r) {
    var noorStatus = String(r['حالة_نور'] || '').trim();
    if (noorStatus !== '' && noorStatus !== 'معلق') return false;
    // فلتر اليوم فقط إذا كان الوضع 'today'
    if (today) {
      var dateVal = r['التاريخ_هجري'] || r['التاريخ هجري'] || '';
      // ★ معالجة حالة تحويل Sheets للتاريخ الهجري إلى Date تلقائياً
      if (dateVal instanceof Date) {
        dateVal = readHijriCellValue_(dateVal);
      }
      var dateStr = String(dateVal);
      if (dateStr && normalizeHijriDate_(dateStr) !== normalizeHijriDate_(today)) return false;
    }
    return true;
  }).map(function(r) {
    r._type = 'absence';
    r._noorMode = NOOR_DROPDOWN_MAP.modes.absence;

    // ★ تحويل التاريخ الهجري من Date إلى نص إذا لزم (لعرض صحيح في الواجهة)
    var hijriKeys = ['التاريخ_هجري', 'التاريخ هجري'];
    for (var hi = 0; hi < hijriKeys.length; hi++) {
      if (r[hijriKeys[hi]] instanceof Date) {
        r[hijriKeys[hi]] = readHijriCellValue_(r[hijriKeys[hi]]);
      }
    }

    // ★ نوع_الغياب يخزّن "يوم كامل" أو "حصة" — ليس نوع العذر
    // المطلوب لنور هو نوع العذر (بعذر/بدون عذر) وليس نوع الغياب
    var excuseType = String(r['نوع_العذر'] || r['نوع العذر'] || '').trim();
    var absStatus = String(r['حالة_التأخر'] || r['حالة التأخر'] || '').trim();

    // تحديد قيمة نور بناءً على نوع العذر
    var absKey = '';
    if (excuseType === 'مقبول' || excuseType === 'بعذر' || excuseType === 'معذور') {
      absKey = 'غياب بعذر';
    } else if (excuseType === 'منصة بعذر' || excuseType.indexOf('منصة') >= 0 && excuseType.indexOf('بدون') === -1) {
      absKey = 'غياب منصة بعذر';
    } else if (excuseType === 'منصة بدون عذر' || excuseType.indexOf('منصة') >= 0 && excuseType.indexOf('بدون') >= 0) {
      absKey = 'غياب منصة بدون عذر';
    } else {
      absKey = 'غياب بدون عذر';
    }

    var mapping = NOOR_DROPDOWN_MAP.absence[absKey];
    if (mapping) {
      r._noorValue = mapping.noorValue;
      r._noorText = mapping.noorText;
    } else {
      // قيمة افتراضية: غياب بدون عذر
      r._noorValue = '48,';
      r._noorText = 'الغياب بدون عذر مقبول';
    }
    return r;
  });
}

// ═══════════════════════════════════════════════════════════════
// تحديث حالة السجلات بعد التوثيق في نور
// ═══════════════════════════════════════════════════════════════

/**
 * تحديث حالة نور للسجلات بعد التوثيق
 * @param {string} stage - المرحلة
 * @param {Array} updates - [{ type, rowIndex, status }]
 */
function updateNoorStatus(stage, updates) {
  try {
    var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    var updated = 0, failed = 0;

    // خريطة الأنواع → أسماء الشيتات
    var typeToRegistry = {
      'violation':    'المخالفات',
      'tardiness':    'التأخر',
      'compensation': 'السلوك_الإيجابي',
      'excellent':    'السلوك_الإيجابي',
      'absence':      'الغياب_اليومي'
    };

    // ★ تجميع التحديثات حسب الشيت لتقليل قراءات الترويسة
    var sheetCache = {};

    updates.forEach(function(u) {
      try {
        var registryKey = typeToRegistry[u.type];
        if (!registryKey) { failed++; return; }

        var sheetName = getSheetName_(registryKey, stage);
        var cacheKey = sheetName;

        // كاش الشيت + عمود حالة_نور (مرة واحدة لكل شيت)
        if (!sheetCache[cacheKey]) {
          var sheet = findSheet_(ss, sheetName);
          if (!sheet) { sheetCache[cacheKey] = { valid: false }; failed++; return; }

          var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
          var noorCol = -1;
          for (var i = 0; i < headers.length; i++) {
            if (String(headers[i]).trim() === 'حالة_نور') { noorCol = i; break; }
          }
          if (noorCol === -1) {
            noorCol = headers.length;
            sheet.getRange(1, noorCol + 1).setValue('حالة_نور');
          }

          sheetCache[cacheKey] = { valid: true, sheet: sheet, noorCol: noorCol };
        }

        var cached = sheetCache[cacheKey];
        if (!cached.valid) { failed++; return; }

        cached.sheet.getRange(u.rowIndex, cached.noorCol + 1).setValue(u.status || 'تم');
        updated++;
      } catch (rowErr) {
        failed++;
        Logger.log('⚠️ خطأ في تحديث صف ' + u.rowIndex + ': ' + rowErr.message);
      }
    });

    return { success: true, updated: updated, failed: failed };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
// إحصائيات نور
// ═══════════════════════════════════════════════════════════════

/**
 * جلب إحصائيات السجلات المعلقة لكل نوع
 * يرجع إحصائيات مزدوجة: today (اليوم) + all (كل غير الموثق)
 * @param {string} stage - المرحلة
 * @param {string} filterMode - اختياري: 'today'/'all' — إذا لم يحدد يرجع الاثنين
 */
function getNoorStats(stage, filterMode) {
  try {
    var documentedToday = countDocumentedToday_(stage);

    // إذا طُلب وضع محدد فقط
    if (filterMode === 'today' || filterMode === 'all') {
      var pending = getNoorPendingRecords(stage, 'all', filterMode);
      return {
        success: true,
        pending: {
          violations:   pending.stats.violations   || 0,
          tardiness:    pending.stats.tardiness    || 0,
          compensation: pending.stats.compensation || 0,
          excellent:    pending.stats.excellent    || 0,
          absence:      pending.stats.absence      || 0,
          total:        pending.total              || 0,
          documentedToday: documentedToday
        }
      };
    }

    // الافتراضي: إرجاع إحصائيات اليوم + كل غير الموثق معاً
    var todayPending = getNoorPendingRecords(stage, 'all', 'today');
    var allPending   = getNoorPendingRecords(stage, 'all', 'all');

    return {
      success: true,
      pending: {
        violations:   todayPending.stats.violations   || 0,
        tardiness:    todayPending.stats.tardiness    || 0,
        compensation: todayPending.stats.compensation || 0,
        excellent:    todayPending.stats.excellent    || 0,
        absence:      todayPending.stats.absence      || 0,
        total:        todayPending.total              || 0,
        documentedToday: documentedToday
      },
      allPending: {
        violations:   allPending.stats.violations   || 0,
        tardiness:    allPending.stats.tardiness    || 0,
        compensation: allPending.stats.compensation || 0,
        excellent:    allPending.stats.excellent    || 0,
        absence:      allPending.stats.absence      || 0,
        total:        allPending.total              || 0
      }
    };
  } catch (e) {
    return {
      success: false,
      error: e.message,
      pending: { violations: 0, tardiness: 0, compensation: 0, excellent: 0, absence: 0, total: 0, documentedToday: 0 },
      allPending: { violations: 0, tardiness: 0, compensation: 0, excellent: 0, absence: 0, total: 0 }
    };
  }
}

// ═══════════════════════════════════════════════════════════════
// دوال مساعدة (private)
// ═══════════════════════════════════════════════════════════════

/**
 * ★ تنظيف السجلات قبل إرسالها للعميل عبر google.script.run
 * google.script.run يحوّل الاستجابة لـ null إذا تضمنت كائنات Date
 * هذه الدالة تحوّل Date → نص قابل للقراءة
 */
function sanitizeRecordsForClient_(records) {
  if (!records || !records.length) return records;
  var tz = Session.getScriptTimeZone();
  return records.map(function(r) {
    var clean = {};
    for (var k in r) {
      if (!r.hasOwnProperty(k)) continue;
      var val = r[k];
      if (val instanceof Date) {
        clean[k] = isNaN(val.getTime()) ? '' : Utilities.formatDate(val, tz, 'yyyy-MM-dd HH:mm');
      } else {
        clean[k] = val;
      }
    }
    return clean;
  });
}

/**
 * تحويل ورقة لمصفوفة كائنات مع _rowIndex
 */
function getSheetAsRecords_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var records = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0] && !data[i][1]) continue; // تخطي الصفوف الفارغة
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      if (headers[j]) {
        row[headers[j]] = data[i][j];
        // إضافة نسخة بالشرطات السفلية والمسافات للتوافق
        row[headers[j].replace(/ /g, '_')] = data[i][j];
        row[headers[j].replace(/_/g, ' ')] = data[i][j];
      }
    }
    row._rowIndex = i + 1;
    records.push(row);
  }
  return records;
}

/**
 * حساب تاريخ القطع (قبل N يوم من اليوم)
 */
function getDateCutoff_(days) {
  var now = new Date();
  now.setDate(now.getDate() - days);
  now.setHours(0, 0, 0, 0);
  return now;
}

/**
 * تحليل تاريخ بأمان (يدعم عدة صيغ)
 */
function parseDateSafe_(dateStr) {
  if (!dateStr) return null;

  // ★ إذا كان كائن Date مباشرة (من getValues)
  if (dateStr instanceof Date) {
    return isNaN(dateStr.getTime()) ? null : dateStr;
  }

  var s = String(dateStr).trim();
  if (!s) return null;

  // صيغة: yyyy-mm-dd أو yyyy/mm/dd
  var match = s.match(/(\d{4})[\-\/](\d{1,2})[\-\/](\d{1,2})/);
  if (match) return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));

  // صيغة: dd/mm/yyyy أو dd-mm-yyyy
  match = s.match(/(\d{1,2})[\-\/](\d{1,2})[\-\/](\d{4})/);
  if (match) return new Date(parseInt(match[3]), parseInt(match[2]) - 1, parseInt(match[1]));

  // محاولة التحليل التلقائي
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * جلب تاريخ اليوم بالهجري (صيغة مبسطة)
 * يستخدم getHijriDate_() من Config.gs لضمان تطابق الصيغة
 * مع التاريخ المحفوظ في السجلات (أرقام إنجليزية بدون أصفار بادئة)
 */
function getTodayHijriDate_() {
  try {
    return getHijriDate_(new Date());
  } catch (e) {
    // fallback: إرجاع null ليتم تخطي فلتر التاريخ
    return null;
  }
}

/**
 * توحيد صيغة التاريخ الهجري للمقارنة
 * يحوّل الأرقام العربية → إنجليزية ويزيل الأصفار البادئة
 * مثال: "٠١/٠٨/١٤٤٦" → "1/8/1446"
 */
function normalizeHijriDate_(dateStr) {
  if (!dateStr) return '';
  var s = String(dateStr).trim();
  // إزالة لاحقة "هـ" والأحرف غير المرئية (zero-width / RTL marks)
  s = s.replace(/[\u200F\u200E\u061C\u200B\u200C\u200D\uFEFF]/g, '');
  s = s.replace(/\s*هـ\s*$/, '').trim();
  // تحويل الأرقام العربية إلى إنجليزية
  s = s.replace(/[٠-٩]/g, function(d) {
    return '٠١٢٣٤٥٦٧٨٩'.indexOf(d);
  });
  // إزالة الأصفار البادئة من كل جزء: 01/08/1446 → 1/8/1446
  s = s.replace(/\b0+(\d)/g, '$1');

  // ★ توحيد الترتيب: إذا كان dd/mm/yyyy → تحويل إلى yyyy/mm/dd
  // getHijriDate_ يُرجع yyyy/mm/dd لكن بعض المصادر تحفظ dd/mm/yyyy
  var parts = s.split('/');
  if (parts.length === 3) {
    var first = parseInt(parts[0]);
    var last = parseInt(parts[2]);
    if (last >= 1300 && last <= 1500 && first <= 30) {
      // الصيغة dd/mm/yyyy → قلبها إلى yyyy/mm/dd
      s = parts[2] + '/' + parts[1] + '/' + parts[0];
    }
  }

  return s;
}

/**
 * ★ حساب عدد السجلات الموثقة اليوم (حالة_نور = 'تم')
 * يبحث في جميع الشيتات ويعد السجلات التي وُثقت اليوم
 */
function countDocumentedToday_(stage) {
  try {
    var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    var count = 0;
    var sheetKeys = ['المخالفات', 'التأخر', 'السلوك_الإيجابي', 'الغياب_اليومي'];

    sheetKeys.forEach(function(key) {
      try {
        var sheetName = getSheetName_(key, stage);
        var sheet = findSheet_(ss, sheetName);
        if (!sheet || sheet.getLastRow() < 2) return;

        var data = sheet.getDataRange().getValues();
        var headers = data[0].map(function(h) { return String(h).trim(); });
        var noorCol = -1, timeCol = -1;
        for (var i = 0; i < headers.length; i++) {
          if (headers[i] === 'حالة_نور') noorCol = i;
          if (headers[i] === 'وقت_الإدخال' || headers[i] === 'وقت الإدخال') timeCol = i;
          if (timeCol === -1 && (headers[i] === 'التاريخ_الميلادي' || headers[i] === 'التاريخ الميلادي')) timeCol = i;
        }
        if (noorCol === -1) return;

        for (var r = 1; r < data.length; r++) {
          if (String(data[r][noorCol] || '').trim() !== 'تم') continue;
          // فحص تاريخ اليوم إذا توفر عمود وقت
          if (timeCol >= 0) {
            var entryDate = parseDateSafe_(data[r][timeCol]);
            if (entryDate && (entryDate < today || entryDate >= tomorrow)) continue;
          }
          count++;
        }
      } catch (e) { /* تجاهل أخطاء الشيتات */ }
    });
    return count;
  } catch (e) {
    return 0;
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// بحث عن مخالفة في خريطة نور بالنص (خطة بديلة عندما يكون رقم المخالفة غير متوفر)
// ═══════════════════════════════════════════════════════════════════════════
function findNoorViolationByText_(text) {
  if (!text) return null;
  var norm = normalizeArabicForMatch_(text);
  var vMap = NOOR_DROPDOWN_MAP.violations;
  for (var id in vMap) {
    if (vMap.hasOwnProperty(id)) {
      var noorNorm = normalizeArabicForMatch_(vMap[id].noorText);
      if (noorNorm.indexOf(norm) >= 0 || norm.indexOf(noorNorm) >= 0) {
        return vMap[id];
      }
    }
  }
  return null;
}

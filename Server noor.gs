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
function getNoorPendingRecords(stage, type) {
  try {
    var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    var result = {
      records: [],
      stats: { violations: 0, tardiness: 0, compensation: 0, excellent: 0, absence: 0 }
    };

    // ═══ 1. المخالفات السلوكية ═══
    if (type === 'violations' || type === 'all') {
      var vRecords = getNoorViolationRecords_(ss, stage);
      result.records = result.records.concat(vRecords);
      result.stats.violations = vRecords.length;
    }

    // ═══ 2. التأخر الصباحي ═══
    if (type === 'tardiness' || type === 'all') {
      var tRecords = getNoorTardinessRecords_(ss, stage);
      result.records = result.records.concat(tRecords);
      result.stats.tardiness = tRecords.length;
    }

    // ═══ 3. السلوك الإيجابي (تعويضية + متمايز) ═══
    if (type === 'compensation' || type === 'excellent' || type === 'positive' || type === 'all') {
      var pRecords = getNoorPositiveRecords_(ss, stage);

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
      var aRecords = getNoorAbsenceRecords_(ss, stage);
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
 * جلب المخالفات المعلقة — فلتر: حالة_نور فارغة/معلق + آخر 14 يوم
 */
function getNoorViolationRecords_(ss, stage) {
  var sheetName = getSheetName_('المخالفات', stage);
  var sheet = findSheet_(ss, sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var records = getSheetAsRecords_(sheet);
  var cutoff = getDateCutoff_(14);

  return records.filter(function(r) {
    var noorStatus = String(r['حالة_نور'] || '').trim();
    if (noorStatus !== '' && noorStatus !== 'معلق') return false;
    // فلتر 14 يوم
    var dateStr = String(r['التاريخ_الميلادي'] || r['التاريخ الميلادي'] || '');
    if (dateStr && cutoff) {
      var recDate = parseDateSafe_(dateStr);
      if (recDate && recDate < cutoff) return false;
    }
    return true;
  }).map(function(r) {
    r._type = 'violation';
    // استبعاد مخالفة 101 (التأخر الصباحي) لأنها تُعرض في تبويب التأخر
    // ملاحظة: لا نستبعدها هنا لأن المخالفات المسجلة يدوياً كمخالفة 101 يجب أن تظهر
    return r;
  });
}

/**
 * جلب سجلات التأخر المعلقة — فلتر: حالة_نور فارغة/معلق + آخر 14 يوم
 * التأخر يُدخل في نور كمخالفة سلوكية: التأخر الصباحي (الدرجة الأولى)
 */
function getNoorTardinessRecords_(ss, stage) {
  var sheetName = getSheetName_('التأخر', stage);
  var sheet = findSheet_(ss, sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var records = getSheetAsRecords_(sheet);
  var cutoff = getDateCutoff_(14);

  return records.filter(function(r) {
    var noorStatus = String(r['حالة_نور'] || '').trim();
    if (noorStatus !== '' && noorStatus !== 'معلق') return false;
    // فلتر 14 يوم (التأخر يستخدم التاريخ_هجري فقط، لذا نتحقق من وقت_الإدخال)
    var dateStr = String(r['وقت_الإدخال'] || r['وقت الإدخال'] || '');
    if (dateStr && cutoff) {
      var recDate = parseDateSafe_(dateStr);
      if (recDate && recDate < cutoff) return false;
    }
    return true;
  }).map(function(r) {
    r._type = 'tardiness';
    r._noorValue = '1601174,الدرجة الأولى'; // التأخر الصباحي دائماً
    return r;
  });
}

/**
 * جلب السلوك الإيجابي المعلق — يُقسم إلى: تعويضية / متمايز
 * فلتر: حالة_نور فارغة/معلق + آخر 14 يوم
 */
function getNoorPositiveRecords_(ss, stage) {
  var sheetName = getSheetName_('السلوك_الإيجابي', stage);
  var sheet = findSheet_(ss, sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var records = getSheetAsRecords_(sheet);
  var cutoff = getDateCutoff_(14);

  return records.filter(function(r) {
    var noorStatus = String(r['حالة_نور'] || '').trim();
    if (noorStatus !== '' && noorStatus !== 'معلق') return false;
    // فلتر 14 يوم
    var dateStr = String(r['التاريخ_الميلادي'] || r['التاريخ الميلادي'] || r['وقت_الإدخال'] || r['وقت الإدخال'] || '');
    if (dateStr && cutoff) {
      var recDate = parseDateSafe_(dateStr);
      if (recDate && recDate < cutoff) return false;
    }
    return true;
  }).map(function(r) {
    var behaviorName = String(r['السلوك_المتمايز'] || r['السلوك المتمايز'] || '');
    // التمييز بين التعويضية والمتمايز
    // التعويضية: السلوك يحتوي على كلمة "تعويض" أو "تعويضي" أو الدرجة تحتوي على "تعويض"
    var degree = String(r['الدرجة'] || '');
    if (behaviorName.indexOf('تعويض') >= 0 || degree.indexOf('تعويض') >= 0) {
      r._type = 'compensation';
    } else {
      r._type = 'excellent';
    }
    return r;
  });
}

/**
 * جلب الغياب اليومي المعلق — فلتر: حالة_نور فارغة/معلق + اليوم فقط
 */
function getNoorAbsenceRecords_(ss, stage) {
  var sheetName = getSheetName_('الغياب_اليومي', stage);
  var sheet = findSheet_(ss, sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var records = getSheetAsRecords_(sheet);
  var today = getTodayHijriDate_();

  return records.filter(function(r) {
    var noorStatus = String(r['حالة_نور'] || '').trim();
    if (noorStatus !== '' && noorStatus !== 'معلق') return false;
    // فلتر اليوم فقط (الغياب يُدخل في نفس اليوم قبل 11:59)
    var dateStr = String(r['التاريخ_هجري'] || r['التاريخ هجري'] || '');
    if (today && dateStr) {
      if (normalizeHijriDate_(dateStr) !== normalizeHijriDate_(today)) return false;
    }
    return true;
  }).map(function(r) {
    r._type = 'absence';
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
 */
function getNoorStats(stage) {
  try {
    var pending = getNoorPendingRecords(stage, 'all');
    var documentedToday = countDocumentedToday_(stage);
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
  } catch (e) {
    return {
      success: false,
      error: e.message,
      pending: { violations: 0, tardiness: 0, compensation: 0, excellent: 0, absence: 0, total: 0, documentedToday: 0 }
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
  // تحويل الأرقام العربية إلى إنجليزية
  s = s.replace(/[٠-٩]/g, function(d) {
    return '٠١٢٣٤٥٦٧٨٩'.indexOf(d);
  });
  // إزالة الأصفار البادئة من كل جزء: 01/08/1446 → 1/8/1446
  s = s.replace(/\b0+(\d)/g, '$1');
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

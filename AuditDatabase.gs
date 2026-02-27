// =================================================================
// AuditDatabase.gs — أداة تدقيق قاعدة البيانات ومطابقتها مع التطبيق
// الاستخدام: شغّل الدالة auditDatabase() من المحرر أو أضفها في القائمة
// =================================================================

// ★ رابط الجدول — غيّره إذا كان مختلفاً عندك
var AUDIT_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1jgtMv8kj7Qrb4WVCuajjKnSGw-5d8mK0CUe1bRRv1q4/edit";

/**
 * ★ الدالة الرئيسية: تدقيق شامل لقاعدة البيانات
 * تعرض النتائج في سجل التنفيذ (Logger) وتنشئ شيت تقرير
 */
function auditDatabase() {
  var ss;
  // محاولة استخدام الدالة المشتركة أولاً، ثم الفتح مباشرة
  try { ss = getSpreadsheet_(); } catch(e) {
    ss = SpreadsheetApp.openByUrl(AUDIT_SPREADSHEET_URL);
  }
  var allSheets = ss.getSheets();

  // ═══════════════════════════════════════════════════
  // 1. تحديد المراحل المفعّلة
  // ═══════════════════════════════════════════════════
  var enabledStages = getEnabledStagesForAudit_(ss);

  // ═══════════════════════════════════════════════════
  // 2. تعريف الشيتات المتوقعة بالكامل (من كود التطبيق)
  // ═══════════════════════════════════════════════════
  var EXPECTED = buildExpectedSheets_(enabledStages);

  // ═══════════════════════════════════════════════════
  // 3. جمع بيانات الشيتات الفعلية
  // ═══════════════════════════════════════════════════
  var actualSheets = {};
  allSheets.forEach(function(sheet) {
    var name = sheet.getName();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var headers = [];
    if (lastCol > 0 && lastRow > 0) {
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
        return String(h || '').trim();
      });
    }
    actualSheets[name] = {
      name: name,
      rows: Math.max(0, lastRow - 1), // عدد الصفوف بدون الترويسة
      cols: lastCol,
      headers: headers
    };
  });

  // ═══════════════════════════════════════════════════
  // 4. المقارنة والتدقيق
  // ═══════════════════════════════════════════════════
  var report = [];
  var summary = { total: 0, found: 0, missing: 0, extraCols: 0, missingCols: 0, warnings: [] };

  // --- (أ) فحص الشيتات المتوقعة ---
  var expectedNames = Object.keys(EXPECTED);
  summary.total = expectedNames.length;

  expectedNames.forEach(function(expName) {
    var exp = EXPECTED[expName];
    var actual = actualSheets[expName];
    var entry = {
      sheet: expName,
      category: exp.category,
      status: '',
      rows: 0,
      cols: 0,
      expectedCols: exp.headers.length,
      actualHeaders: [],
      expectedHeaders: exp.headers,
      missingHeaders: [],
      extraHeaders: [],
      notes: []
    };

    if (!actual) {
      entry.status = '❌ غير موجود';
      entry.notes.push('الشيت مطلوب من التطبيق ولكنه غير موجود في قاعدة البيانات');
      summary.missing++;
    } else {
      entry.status = '✅ موجود';
      entry.rows = actual.rows;
      entry.cols = actual.cols;
      entry.actualHeaders = actual.headers;
      summary.found++;

      // مقارنة الأعمدة
      var colResult = compareHeaders_(exp.headers, actual.headers);
      entry.missingHeaders = colResult.missing;
      entry.extraHeaders = colResult.extra;
      summary.missingCols += colResult.missing.length;
      summary.extraCols += colResult.extra.length;

      if (colResult.missing.length > 0) {
        entry.notes.push('⚠️ أعمدة ناقصة: ' + colResult.missing.join(' | '));
      }
      if (colResult.extra.length > 0) {
        entry.notes.push('ℹ️ أعمدة إضافية: ' + colResult.extra.join(' | '));
      }
      if (actual.rows === 0) {
        entry.notes.push('📭 الشيت فارغ (بدون بيانات)');
      }

      // تحقق ترتيب الأعمدة
      var orderIssues = checkHeaderOrder_(exp.headers, actual.headers);
      if (orderIssues.length > 0) {
        entry.notes.push('🔄 ترتيب مختلف: ' + orderIssues.join(' | '));
      }
    }

    report.push(entry);
  });

  // --- (ب) شيتات إضافية غير متوقعة في قاعدة البيانات ---
  var extraSheets = [];
  Object.keys(actualSheets).forEach(function(name) {
    if (!EXPECTED[name]) {
      extraSheets.push(actualSheets[name]);
    }
  });

  // ═══════════════════════════════════════════════════
  // 5. طباعة التقرير في Logger
  // ═══════════════════════════════════════════════════
  printReport_(report, extraSheets, summary, enabledStages, allSheets.length);

  // ═══════════════════════════════════════════════════
  // 6. إنشاء شيت تقرير مفصّل
  // ═══════════════════════════════════════════════════
  createReportSheet_(ss, report, extraSheets, summary, enabledStages);

  return { success: true, message: 'تم إنشاء التقرير — افتح شيت "📊 تقرير_التدقيق" أو سجل التنفيذ' };
}


// =================================================================
// ★ بناء قائمة الشيتات المتوقعة من كود التطبيق
// =================================================================
function buildExpectedSheets_(stages) {
  var expected = {};

  // ─── (1) شيتات النظام (ثابتة) ───
  var systemSheets = {
    'إعدادات_المدرسة': {
      category: 'نظام',
      headers: ['المفتاح', 'القيمة', 'الوصف', 'تاريخ التحديث']
    },
    'هيكل_المدرسة': {
      category: 'نظام',
      headers: ['المفتاح', 'القيمة', 'تاريخ التحديث']
    },
    'المستخدمين': {
      category: 'نظام',
      headers: ['المعرف', 'الاسم', 'الدور', 'الجوال', 'البريد الإلكتروني', 'الصلاحيات', 'نوع النطاق', 'قيمة النطاق', 'الحالة', 'تاريخ الإنشاء', 'تاريخ التحديث', 'رمز_الرابط', 'الرابط'],
      notes: 'العمودان 12-13 يُضافان تلقائياً عند تفعيل روابط المعلمين'
    },
    'المعلمين': {
      category: 'نظام',
      headers: ['المعرف', 'السجل المدني', 'الاسم', 'الجوال', 'المواد', 'الفصول المسندة', 'الصلاحيات', 'الحالة', 'تاريخ الإنشاء', 'تاريخ التحديث', 'رمز_الرابط', 'الرابط', 'تاريخ_التفعيل'],
      notes: 'الأعمدة 11-13 تُضاف تلقائياً عند تفعيل روابط المعلمين'
    },
    'المواد': {
      category: 'نظام',
      headers: ['المعرف', 'اسم المادة', 'تاريخ الإضافة']
    },
    'اللجان': {
      category: 'نظام',
      headers: ['المعرف', 'الاسم', 'الأعضاء', 'الحالة']
    },
    'روابط_المعلمين': {
      category: 'نظام',
      headers: ['الجوال', 'الاسم', 'النوع', 'المرحلة', 'الفصول', 'تم الربط بواسطة', 'تاريخ الربط']
    },
    'إعدادات_واتساب': {
      category: 'نظام',
      headers: ['الإعداد', 'القيمة', 'الوصف']
    },
    'جلسات_واتساب': {
      category: 'نظام',
      headers: ['رقم_الواتساب', 'المرحلة', 'نوع_المستخدم', 'حالة_الاتصال', 'تاريخ_الربط', 'آخر_استخدام', 'عدد_الرسائل', 'الرقم_الرئيسي']
    },
    'سجل_النشاطات': {
      category: 'نظام',
      headers: ['التاريخ', 'الوقت', 'المستخدم', 'النوع', 'التفاصيل', 'العدد', 'المرحلة']
    },
    'أنواع_الملاحظات_التربوية': {
      category: 'نظام',
      headers: ['المرحلة', 'نوع الملاحظة', 'تاريخ الإضافة']
    },
    'اعذار_اولياء_الامور': {
      category: 'نظام',
      headers: ['رقم_الطالب', 'اسم_الطالب', 'الصف', 'الفصل', 'المرحلة', 'نص_العذر', 'مرفقات', 'تاريخ_الغياب', 'تاريخ_التقديم', 'وقت_التقديم', 'الحالة', 'ملاحظات_المدرسة', 'الرمز']
    },
    'رموز_اولياء_الامور': {
      category: 'نظام',
      headers: ['الرمز', 'رقم_الطالب', 'المرحلة', 'تاريخ_الإنشاء', 'تاريخ_الانتهاء', 'مستخدم']
    }
  };

  for (var sName in systemSheets) {
    expected[sName] = systemSheets[sName];
  }

  // ─── (2) شيتات الطلاب (حسب المراحل المفعّلة) ───
  var STAGE_SHEET_MAP = {
    'طفولة مبكرة': 'طلاب_طفولة_مبكرة',
    'ابتدائي': 'طلاب_ابتدائي',
    'متوسط': 'طلاب_متوسط',
    'ثانوي': 'طلاب_ثانوي'
  };

  stages.forEach(function(stage) {
    var sheetName = STAGE_SHEET_MAP[stage] || ('طلاب_' + stage);
    expected[sheetName] = {
      category: 'طلاب — ' + stage,
      headers: ['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'رقم الجوال', 'تاريخ الإضافة']
    };
  });

  // ─── (3) شيتات السجلات لكل مرحلة (من SHEET_REGISTRY) ───
  var perStageRecords = {
    'سجل_المخالفات_': {
      label: 'مخالفات',
      headers: ['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'رقم المخالفة', 'نص المخالفة', 'نوع المخالفة', 'الدرجة', 'التاريخ الهجري', 'التاريخ الميلادي', 'مستوى التكرار', 'الإجراءات', 'النقاط', 'اليوم', 'النماذج المحفوظة', 'المستخدم', 'وقت الإدخال', 'تم الإرسال']
    },
    'سجل_التأخر_': {
      label: 'تأخر',
      headers: ['رقم_الطالب', 'اسم_الطالب', 'الصف', 'الفصل', 'رقم_الجوال', 'نوع_التأخر', 'الحصة', 'التاريخ_هجري', 'المسجل', 'وقت_الإدخال', 'تم_الإرسال']
    },
    'سجل_الاستئذان_': {
      label: 'استئذان',
      headers: ['رقم_الطالب', 'اسم_الطالب', 'الصف', 'الفصل', 'رقم_الجوال', 'وقت_الخروج', 'السبب', 'المستلم', 'المسؤول', 'التاريخ_هجري', 'المسجل', 'وقت_الإدخال', 'وقت_التأكيد', 'تم_الإرسال']
    },
    'سجل_الغياب_': {
      label: 'غياب تراكمي',
      headers: ['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'غياب بعذر', 'غياب بدون عذر', 'تأخير', 'آخر تحديث']
    },
    'سجل_الغياب_اليومي_': {
      label: 'غياب يومي',
      headers: ['رقم_الطالب', 'اسم_الطالب', 'الصف', 'الفصل', 'رقم_الجوال', 'نوع_الغياب', 'الحصة', 'التاريخ_هجري', 'اليوم', 'المسجل', 'وقت_الإدخال', 'حالة_الاعتماد', 'نوع_العذر', 'تم_الإرسال', 'حالة_التأخر', 'وقت_الحضور', 'ملاحظات']
    },
    'سجل_الملاحظات_التربوية_': {
      label: 'ملاحظات تربوية',
      headers: ['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'رقم الجوال', 'نوع الملاحظة', 'التفاصيل', 'المعلم/المسجل', 'التاريخ', 'وقت الإدخال', 'تم الإرسال']
    },
    'سجل_التواصل_': {
      label: 'تواصل',
      headers: ['م', 'التاريخ الهجري', 'التاريخ الميلادي', 'الوقت', 'رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'رقم الجوال', 'نوع الرسالة', 'عنوان الرسالة', 'نص الرسالة', 'حالة الإرسال', 'المرسل', 'ملاحظات']
    },
    'سجل_التحصيل_': {
      label: 'تحصيل',
      headers: ['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'المادة', 'نوع التقييم', 'الدرجة', 'من', 'التاريخ', 'المعلم', 'ملاحظات', 'وقت الإدخال']
    },
    'سجل_السلوك_الإيجابي_': {
      label: 'سلوك إيجابي',
      headers: ['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'رقم الجوال', 'السلوك المتمايز', 'الدرجة', 'المعلم', 'اليوم', 'التاريخ الهجري', 'التاريخ الميلادي', 'وقت الإدخال', 'تم الإرسال']
    }
  };

  stages.forEach(function(stage) {
    for (var prefix in perStageRecords) {
      var rec = perStageRecords[prefix];
      expected[prefix + stage] = {
        category: rec.label + ' — ' + stage,
        headers: rec.headers
      };
    }

    // شيتات التحصيل الأكاديمي
    expected['تحصيل_ملخص_' + stage] = {
      category: 'تحصيل ملخص — ' + stage,
      headers: ['رقم_الهوية', 'اسم_الطالب', 'الصف', 'الفصل', 'الفصل_الدراسي', 'الفترة', 'المعدل', 'التقدير_العام', 'ترتيب_الصف', 'ترتيب_الفصل', 'الغياب', 'التأخر', 'السلوك_متميز', 'السلوك_إيجابي']
    };
    expected['تحصيل_درجات_' + stage] = {
      category: 'تحصيل درجات — ' + stage,
      headers: ['رقم_الهوية', 'اسم_الطالب', 'الصف', 'الفصل', 'الفصل_الدراسي', 'الفترة', 'المادة', 'المجموع', 'اختبار_نهائي', 'أدوات_تقييم', 'اختبارات_قصيرة', 'التقدير']
    };
  });

  return expected;
}


// =================================================================
// ★ اكتشاف المراحل المفعّلة — مستقل عن Config.gs
// يبحث عن شيتات الطلاب الموجودة فعلياً (طلاب_متوسط، طلاب_ثانوي...)
// =================================================================
function getEnabledStagesForAudit_(ss) {
  // ★ المراحل الممكنة وأسماء شيتات الطلاب المقابلة
  var STAGE_MAP = {
    'طلاب_طفولة_مبكرة': 'طفولة مبكرة',
    'طلاب_ابتدائي': 'ابتدائي',
    'طلاب_متوسط': 'متوسط',
    'طلاب_ثانوي': 'ثانوي'
  };
  var stages = [];

  // ★ الطريقة الأضمن: البحث عن شيتات الطلاب الموجودة فعلياً
  var allSheetNames = ss.getSheets().map(function(s) { return s.getName(); });
  for (var sheetName in STAGE_MAP) {
    if (allSheetNames.indexOf(sheetName) > -1) {
      stages.push(STAGE_MAP[sheetName]);
    }
  }

  // fallback
  if (stages.length === 0) {
    stages = ['متوسط', 'ثانوي'];
  }

  return stages;
}


// =================================================================
// ★ مقارنة الأعمدة المتوقعة مع الفعلية (مع مراعاة _ و مسافة)
// =================================================================
function compareHeaders_(expected, actual) {
  // تطبيع الاسم: إزالة الفروق بين _ والمسافة
  function normalize(h) {
    return String(h || '').trim().replace(/[\s_]+/g, '_');
  }

  var actualNorm = actual.map(normalize);
  var expectedNorm = expected.map(normalize);

  var missing = [];
  var extra = [];

  // أعمدة ناقصة (متوقعة لكن غير موجودة)
  for (var i = 0; i < expectedNorm.length; i++) {
    if (actualNorm.indexOf(expectedNorm[i]) === -1) {
      missing.push(expected[i]);
    }
  }

  // أعمدة إضافية (موجودة لكن غير متوقعة)
  for (var j = 0; j < actualNorm.length; j++) {
    if (actualNorm[j] && expectedNorm.indexOf(actualNorm[j]) === -1) {
      extra.push(actual[j]);
    }
  }

  return { missing: missing, extra: extra };
}


// =================================================================
// ★ فحص ترتيب الأعمدة
// =================================================================
function checkHeaderOrder_(expected, actual) {
  function normalize(h) {
    return String(h || '').trim().replace(/[\s_]+/g, '_');
  }

  var issues = [];
  var actualNorm = actual.map(normalize);

  for (var i = 0; i < expected.length; i++) {
    var expNorm = normalize(expected[i]);
    var actualIdx = actualNorm.indexOf(expNorm);
    if (actualIdx !== -1 && actualIdx !== i) {
      issues.push(expected[i] + ': متوقع في العمود ' + (i+1) + ' لكنه في ' + (actualIdx+1));
    }
  }

  return issues;
}


// =================================================================
// ★ طباعة التقرير في سجل التنفيذ (Logger)
// =================================================================
function printReport_(report, extraSheets, summary, stages, totalSheets) {
  var L = Logger;

  L.log('╔══════════════════════════════════════════════════════════════╗');
  L.log('║     📊 تقرير تدقيق قاعدة البيانات — نظام المخالفات السلوكية     ║');
  L.log('╠══════════════════════════════════════════════════════════════╣');
  L.log('║ التاريخ: ' + new Date().toLocaleString('ar-SA'));
  L.log('║ المراحل المفعّلة: ' + stages.join(' | '));
  L.log('║ إجمالي الشيتات في قاعدة البيانات: ' + totalSheets);
  L.log('║ الشيتات المتوقعة من التطبيق: ' + summary.total);
  L.log('║ ✅ موجودة: ' + summary.found + '  |  ❌ ناقصة: ' + summary.missing);
  L.log('║ أعمدة ناقصة: ' + summary.missingCols + '  |  أعمدة إضافية: ' + summary.extraCols);
  L.log('╚══════════════════════════════════════════════════════════════╝');
  L.log('');

  // --- الشيتات المتوقعة ---
  var prevCategory = '';
  report.forEach(function(entry) {
    // عنوان الفئة
    var cat = entry.category.split(' — ')[0];
    if (cat !== prevCategory) {
      L.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      L.log('📁 ' + cat);
      L.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      prevCategory = cat;
    }

    L.log('  ' + entry.status + ' ' + entry.sheet);
    if (entry.status.indexOf('✅') > -1) {
      L.log('      📊 صفوف: ' + entry.rows + ' | أعمدة فعلية: ' + entry.cols + ' | أعمدة متوقعة: ' + entry.expectedCols);
      L.log('      📋 الأعمدة الفعلية: ' + entry.actualHeaders.join(' | '));
    }
    L.log('      📐 الأعمدة المتوقعة: ' + entry.expectedHeaders.join(' | '));

    if (entry.notes.length > 0) {
      entry.notes.forEach(function(n) { L.log('      ' + n); });
    }
    L.log('');
  });

  // --- شيتات إضافية ---
  if (extraSheets.length > 0) {
    L.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    L.log('📦 شيتات إضافية (غير مُعرّفة في كود التطبيق): ' + extraSheets.length);
    L.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    extraSheets.forEach(function(s) {
      L.log('  🔹 ' + s.name + '  (صفوف: ' + s.rows + ' | أعمدة: ' + s.cols + ')');
      if (s.headers.length > 0) {
        L.log('      📋 ' + s.headers.join(' | '));
      }
    });
  }

  // --- ملخص المشاكل ---
  L.log('');
  L.log('═══════════════════════════════════════════');
  L.log('🔍 ملخص المشاكل المكتشفة:');
  L.log('═══════════════════════════════════════════');

  var problems = report.filter(function(e) {
    return e.status.indexOf('❌') > -1 || e.missingHeaders.length > 0;
  });

  if (problems.length === 0) {
    L.log('  ✅ لا توجد مشاكل — قاعدة البيانات مطابقة للتطبيق!');
  } else {
    problems.forEach(function(p) {
      if (p.status.indexOf('❌') > -1) {
        L.log('  ❌ شيت مفقود: ' + p.sheet + ' ← مطلوب من: ' + p.category);
      }
      if (p.missingHeaders.length > 0) {
        L.log('  ⚠️ أعمدة ناقصة في [' + p.sheet + ']: ' + p.missingHeaders.join(', '));
      }
    });
  }
}


// =================================================================
// ★ إنشاء شيت تقرير مرئي مع تنسيق ملون
// =================================================================
function createReportSheet_(ss, report, extraSheets, summary, stages) {
  var REPORT_NAME = '📊 تقرير_التدقيق';

  // حذف الشيت السابق إن وجد
  var existing = ss.getSheetByName(REPORT_NAME);
  if (existing) ss.deleteSheet(existing);

  var rs = ss.insertSheet(REPORT_NAME, 0);
  rs.setRightToLeft(true);

  var row = 1;
  var BLUE = '#1e3a5f';
  var WHITE = '#ffffff';
  var GREEN_BG = '#d4edda';
  var RED_BG = '#f8d7da';
  var YELLOW_BG = '#fff3cd';
  var GRAY_BG = '#e9ecef';

  // ─── العنوان الرئيسي ───
  rs.getRange(row, 1, 1, 10).merge()
    .setValue('📊 تقرير تدقيق قاعدة البيانات — ' + new Date().toLocaleDateString('ar-SA'))
    .setBackground(BLUE).setFontColor(WHITE).setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center');
  row++;

  // ─── الملخص ───
  var summaryData = [
    ['المراحل المفعّلة', stages.join(' ، ')],
    ['الشيتات المتوقعة', summary.total],
    ['✅ موجودة', summary.found],
    ['❌ ناقصة', summary.missing],
    ['⚠️ أعمدة ناقصة', summary.missingCols],
    ['ℹ️ أعمدة إضافية', summary.extraCols]
  ];
  summaryData.forEach(function(r) {
    rs.getRange(row, 1).setValue(r[0]).setFontWeight('bold').setBackground(GRAY_BG);
    rs.getRange(row, 2, 1, 2).merge().setValue(r[1]);
    row++;
  });
  row++;

  // ─── جدول التفاصيل ───
  var headerRow = ['#', 'اسم الشيت', 'الفئة', 'الحالة', 'الصفوف', 'أعمدة فعلية', 'أعمدة متوقعة', 'أعمدة ناقصة', 'أعمدة إضافية', 'ملاحظات'];
  rs.getRange(row, 1, 1, headerRow.length).setValues([headerRow])
    .setBackground(BLUE).setFontColor(WHITE).setFontWeight('bold')
    .setHorizontalAlignment('center');
  row++;

  report.forEach(function(entry, idx) {
    var isMissing = entry.status.indexOf('❌') > -1;
    var hasIssues = entry.missingHeaders.length > 0;
    var bg = isMissing ? RED_BG : (hasIssues ? YELLOW_BG : GREEN_BG);

    var rowData = [
      idx + 1,
      entry.sheet,
      entry.category,
      isMissing ? '❌ غير موجود' : '✅ موجود',
      isMissing ? '-' : entry.rows,
      isMissing ? '-' : entry.cols,
      entry.expectedCols,
      entry.missingHeaders.length > 0 ? entry.missingHeaders.join('\n') : '—',
      entry.extraHeaders.length > 0 ? entry.extraHeaders.join('\n') : '—',
      entry.notes.join('\n') || '—'
    ];

    rs.getRange(row, 1, 1, rowData.length).setValues([rowData]).setBackground(bg)
      .setVerticalAlignment('top').setWrap(true);
    row++;
  });

  row++;

  // ─── شيتات إضافية ───
  if (extraSheets.length > 0) {
    rs.getRange(row, 1, 1, 10).merge()
      .setValue('📦 شيتات إضافية (غير مُعرّفة في كود التطبيق): ' + extraSheets.length)
      .setBackground('#6c757d').setFontColor(WHITE).setFontWeight('bold')
      .setHorizontalAlignment('center');
    row++;

    rs.getRange(row, 1, 1, 5).setValues([['#', 'اسم الشيت', 'الصفوف', 'الأعمدة', 'أسماء الأعمدة']])
      .setBackground(GRAY_BG).setFontWeight('bold');
    row++;

    extraSheets.forEach(function(s, idx) {
      rs.getRange(row, 1, 1, 5).setValues([[
        idx + 1,
        s.name,
        s.rows,
        s.cols,
        s.headers.join(' | ')
      ]]).setWrap(true);
      row++;
    });
  }

  row += 2;

  // ─── تفصيل الأعمدة لكل شيت ───
  rs.getRange(row, 1, 1, 10).merge()
    .setValue('📋 تفصيل أعمدة كل شيت (الفعلية مقابل المتوقعة)')
    .setBackground(BLUE).setFontColor(WHITE).setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center');
  row++;

  report.forEach(function(entry) {
    if (entry.status.indexOf('❌') > -1) return; // تخطي الشيتات المفقودة

    rs.getRange(row, 1, 1, 10).merge()
      .setValue('🔹 ' + entry.sheet + ' (' + entry.category + ')')
      .setBackground('#dee2e6').setFontWeight('bold');
    row++;

    // صف ترويسة
    rs.getRange(row, 1).setValue('').setBackground(GRAY_BG);
    var maxCols = Math.max(entry.actualHeaders.length, entry.expectedHeaders.length);

    // الأعمدة الفعلية
    rs.getRange(row, 1).setValue('الفعلي:').setFontWeight('bold').setBackground('#d1ecf1');
    for (var c = 0; c < entry.actualHeaders.length; c++) {
      var cellBg = '#d1ecf1';
      // تلوين إذا كانت غير متوقعة
      var norm = String(entry.actualHeaders[c]).replace(/[\s_]+/g, '_');
      var isExtra = entry.extraHeaders.some(function(eh) {
        return String(eh).replace(/[\s_]+/g, '_') === norm;
      });
      if (isExtra) cellBg = YELLOW_BG;
      rs.getRange(row, c + 2).setValue(entry.actualHeaders[c]).setBackground(cellBg);
    }
    row++;

    rs.getRange(row, 1).setValue('المتوقع:').setFontWeight('bold').setBackground('#e2e3e5');
    for (var c2 = 0; c2 < entry.expectedHeaders.length; c2++) {
      var cellBg2 = '#e2e3e5';
      var isMissing2 = entry.missingHeaders.indexOf(entry.expectedHeaders[c2]) > -1;
      if (isMissing2) cellBg2 = RED_BG;
      rs.getRange(row, c2 + 2).setValue(entry.expectedHeaders[c2]).setBackground(cellBg2);
    }
    row++;
    row++;
  });

  // تعديل عرض الأعمدة
  rs.setColumnWidth(1, 40);   // #
  rs.setColumnWidth(2, 250);  // اسم الشيت
  rs.setColumnWidth(3, 150);  // الفئة
  rs.setColumnWidth(4, 120);  // الحالة
  rs.setColumnWidth(5, 70);   // الصفوف
  rs.setColumnWidth(6, 90);   // أعمدة فعلية
  rs.setColumnWidth(7, 100);  // أعمدة متوقعة
  rs.setColumnWidth(8, 180);  // أعمدة ناقصة
  rs.setColumnWidth(9, 180);  // أعمدة إضافية
  rs.setColumnWidth(10, 300); // ملاحظات

  // تجميد الصفوف العليا
  rs.setFrozenRows(1);

  SpreadsheetApp.flush();
}

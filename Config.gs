// =================================================================
// CONFIGURATION - إعدادات المشروع
// =================================================================
var SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1jgtMv8kj7Qrb4WVCuajjKnSGw-5d8mK0CUe1bRRv1q4/edit";

// ★ كاش مرجع الجدول — يُفتح مرة واحدة فقط لكل تنفيذ بدلاً من 4+ مرات
var _cachedSpreadsheet = null;
function getSpreadsheet_() {
  if (!_cachedSpreadsheet) {
    _cachedSpreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  }
  return _cachedSpreadsheet;
}

// =================================================================
// ★ شيتات الطلاب — تُبنى ديناميكياً من هيكل_المدرسة
// لا تُعدّل يدوياً! يُملأ تلقائياً عند أول استخدام
// =================================================================
var STUDENTS_SHEETS = null; // يُبنى من هيكل_المدرسة (lazy)

// خرائط ثابتة: ربط معرّف المرحلة → الاسم العربي → اسم شيت الطلاب
var STAGE_ID_TO_ARABIC = {
  'kindergarten': 'طفولة مبكرة',
  'primary': 'ابتدائي',
  'intermediate': 'متوسط',
  'secondary': 'ثانوي'
};
var STAGE_ARABIC_TO_ID = {
  'طفولة مبكرة': 'kindergarten',
  'ابتدائي': 'primary',
  'متوسط': 'intermediate',
  'ثانوي': 'secondary'
};
var STAGE_ARABIC_TO_SHEET = {
  'طفولة مبكرة': 'طلاب_طفولة_مبكرة',
  'ابتدائي': 'طلاب_ابتدائي',
  'متوسط': 'طلاب_متوسط',
  'ثانوي': 'طلاب_ثانوي'
};

// =================================================================
// ★ بناء STUDENTS_SHEETS من هيكل_المدرسة
// تقرأ الشيت → تبحث عن المراحل المفعّلة → تبني الخريطة
// =================================================================
function buildStudentsSheetsFromStructure_() {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('هيكل_المدرسة');

    if (!sheet || sheet.getLastRow() < 2) {
      return null; // الهيكل غير مُعدّ
    }

    var data = sheet.getDataRange().getValues();
    var result = {};
    var secondarySystem = 'فصلي';

    // قراءة النظام الثانوي أولاً
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === 'secondary_system') {
        secondarySystem = data[i][1] || 'فصلي';
        break;
      }
    }

    // قراءة المراحل المفعّلة
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== 'stage_config') continue;

      try {
        var stageData = JSON.parse(data[i][1]);
        if (!stageData.stageId || !stageData.grades) continue;

        var arabicName = STAGE_ID_TO_ARABIC[stageData.stageId];
        if (!arabicName) continue;

        // فحص: هل فيه صف مفعّل واحد على الأقل بفصل واحد على الأقل؟
        var hasEnabled = false;
        var grades = stageData.grades;
        for (var g in grades) {
          if (grades[g] && grades[g].enabled && grades[g].classCount > 0) {
            hasEnabled = true;
            break;
          }
        }

        if (hasEnabled) {
          result[arabicName] = STAGE_ARABIC_TO_SHEET[arabicName];
        }
      } catch(e) {}
    }

    return Object.keys(result).length > 0 ? result : null;
  } catch(e) {
    Logger.log('⚠️ خطأ في قراءة هيكل_المدرسة: ' + e.toString());
    return null;
  }
}

// ★ تأكد من تحميل STUDENTS_SHEETS (lazy initialization)
function ensureStudentsSheetsLoaded_() {
  if (STUDENTS_SHEETS !== null) return;
  STUDENTS_SHEETS = buildStudentsSheetsFromStructure_();
  if (!STUDENTS_SHEETS) {
    STUDENTS_SHEETS = {}; // فارغ — الهيكل غير مُعدّ
  }
}

// ★ فحص: هل تم إعداد هيكل المدرسة؟
function isStructureConfigured_() {
  ensureStudentsSheetsLoaded_();
  return Object.keys(STUDENTS_SHEETS).length > 0;
}

// ★ مسح كاش STUDENTS_SHEETS (يُستدعى بعد حفظ الهيكل)
function resetStudentsSheetsCache_() {
  STUDENTS_SHEETS = null;
}

// ★ جلب المراحل المفعّلة مع تفاصيل الصفوف (للحقن في الفورمات)
function getEnabledStagesConfig_() {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('هيكل_المدرسة');
    if (!sheet || sheet.getLastRow() < 2) return { stages: {}, secondarySystem: 'فصلي' };

    var data = sheet.getDataRange().getValues();
    var config = { stages: {}, secondarySystem: 'فصلي' };

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === 'secondary_system') {
        config.secondarySystem = data[i][1] || 'فصلي';
      } else if (data[i][0] === 'stage_config') {
        try {
          var stageData = JSON.parse(data[i][1]);
          var arabicName = STAGE_ID_TO_ARABIC[stageData.stageId];
          if (!arabicName || !stageData.grades) continue;

          var enabledGrades = [];
          for (var g in stageData.grades) {
            if (stageData.grades[g] && stageData.grades[g].enabled && stageData.grades[g].classCount > 0) {
              enabledGrades.push(g + ' ' + arabicName);
            }
          }
          if (enabledGrades.length > 0) {
            config.stages[arabicName] = enabledGrades;
          }
        } catch(e) {}
      }
    }
    return config;
  } catch(e) {
    return { stages: {}, secondarySystem: 'فصلي' };
  }
}

// دالة جلب شيت طلاب مرحلة معينة
function getStudentsSheet_(stage) {
  ensureStudentsSheetsLoaded_();
  var ss = getSpreadsheet_();
  var name = STUDENTS_SHEETS[stage];
  if (!name) return null;
  return ss.getSheetByName(name);
}

// دالة جلب جميع شيتات الطلاب الموجودة
function getAllStudentsSheets_() {
  ensureStudentsSheetsLoaded_();
  var ss = getSpreadsheet_();
  var result = [];
  var stages = Object.keys(STUDENTS_SHEETS);
  for (var i = 0; i < stages.length; i++) {
    var sheet = ss.getSheetByName(STUDENTS_SHEETS[stages[i]]);
    if (sheet) {
      result.push({ sheet: sheet, stage: stages[i] });
    }
  }
  return result;
}

// دالة تنظيف اسم الصف
function cleanGradeName_(grade) {
  var s = String(grade || '')
    .replace(/_/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  
  // ١. حذف لواحق نور (قسم عام، السنة المشتركة، المسار العام، نظام عام)
  s = s.replace(/\s*(قسم عام|السنة المشتركة|نظام عام|المسار العام)/g, '').trim();
  
  // ٢. توحيد "المتوسط" → "متوسط" (كلمة كاملة فقط)
  s = s.replace(/(^|\s)المتوسط(\s|$)/g, '$1متوسط$2');

  // ٣. توحيد "الثانوي" → "ثانوي" (كلمة كاملة فقط)
  s = s.replace(/(^|\s)الثانوي(\s|$)/g, '$1ثانوي$2');

  // ٤. توحيد "الابتدائي" → "ابتدائي" (كلمة كاملة فقط)
  s = s.replace(/(^|\s)الابتدائي(\s|$)/g, '$1ابتدائي$2');

  return s.replace(/\s+/g, ' ').trim();
}

// دالة اكتشاف المرحلة من اسم الصف
function detectStageFromGrade_(grade) {
  grade = String(grade || '');
  if (grade.includes('ثانو')) return 'ثانوي';
  if (grade.includes('متوسط')) return 'متوسط';
  if (grade.includes('ابتدا')) return 'ابتدائي';
  if (grade.includes('طفولة') || grade.includes('روضة')) return 'طفولة مبكرة';
  return '';
}
// =================================================================
// ★ ثوابت أسماء الشيتات - التسمية المعتمدة النهائية
// القاعدة: سجل_[النوع]_[المرحلة] (شرطات سفلية دائماً)
// =================================================================
var LOG_SHEET_INTERMEDIATE = "سجل_المخالفات_متوسط";
var LOG_SHEET_SECONDARY = "سجل_المخالفات_ثانوي";

var ABSENCE_SHEET_INT_NAME = "سجل_الغياب_متوسط";
var ABSENCE_SHEET_SEC_NAME = "سجل_الغياب_ثانوي";

var SETTINGS_SHEET_NAME = "إعدادات_المدرسة";
var ACADEMIC_SHEET_NAME = "التحصيل الدراسي";

var SCHOOL_SETTINGS_SHEET = "إعدادات_المدرسة";
var SCHOOL_STRUCTURE_SHEET = "هيكل_المدرسة";
var USERS_SHEET = "المستخدمين";
var COMMITTEES_SHEET = "اللجان";
var TEACHERS_SHEET = "المعلمين";

// =================================================================
// ★ السجل المركزي للشيتات (SHEET_REGISTRY)
// كل نوع شيت معرّف هنا مرة واحدة - القاعدة لكل شيء
// prefix + '_' + stage = الاسم النهائي
// =================================================================
var SHEET_REGISTRY = {
  'المخالفات':           { prefix: 'سجل_المخالفات',           color: '#e74c3c', perStage: true },
  'التأخر':              { prefix: 'سجل_التأخر',              color: '#dc2626', perStage: true },
  'الاستئذان':           { prefix: 'سجل_الاستئذان',           color: '#7c3aed', perStage: true },
  'الغياب':              { prefix: 'سجل_الغياب',              color: '#f59e0b', perStage: true },
  'الغياب_اليومي':       { prefix: 'سجل_الغياب_اليومي',       color: '#ea580c', perStage: true },
  'الملاحظات_التربوية':  { prefix: 'سجل_الملاحظات_التربوية',  color: '#1e3a5f', perStage: true },
  'التواصل':             { prefix: 'سجل_التواصل',             color: '#4a5568', perStage: true },
  'التحصيل':             { prefix: 'سجل_التحصيل',             color: '#059669', perStage: true },
  'السلوك_الإيجابي':     { prefix: 'سجل_السلوك_الإيجابي',     color: '#10b981', perStage: true },
  'تحصيل_ملخص':          { prefix: 'تحصيل_ملخص',              color: '#059669', perStage: true },
  'تحصيل_درجات':         { prefix: 'تحصيل_درجات',             color: '#0d9488', perStage: true }
};

// =================================================================
// ★ تعريف شيتات النظام الأساسية (SYSTEM_SHEETS_DEFINITIONS)
// تُستخدم في setupNewSchool() و ensureAllSheets_()
// الترويسات مطابقة لما في Server_Settings.js, Server_Users.js, etc.
// =================================================================
var SYSTEM_SHEETS_DEFINITIONS = {
  'إعدادات_المدرسة': {
    headers: ['المفتاح', 'القيمة', 'الوصف', 'تاريخ التحديث'],
    tabColor: '#4285f4'
  },
  'هيكل_المدرسة': {
    headers: ['المفتاح', 'القيمة', 'تاريخ التحديث'],
    tabColor: '#0f9d58'
  },
  'المستخدمين': {
    headers: ['المعرف', 'الاسم', 'الدور', 'الجوال', 'البريد الإلكتروني', 'الصلاحيات', 'نوع النطاق', 'قيمة النطاق', 'الحالة', 'تاريخ الإنشاء', 'تاريخ التحديث', 'رمز_الرابط', 'الرابط'],
    tabColor: '#e91e63'
  },
  'المعلمين': {
    headers: ['المعرف', 'السجل المدني', 'الاسم', 'الجوال', 'المواد', 'الفصول المسندة', 'الصلاحيات', 'الحالة', 'تاريخ الإنشاء', 'تاريخ التحديث', 'رمز_الرابط', 'الرابط', 'تاريخ_التفعيل'],
    tabColor: '#9c27b0'
  },
  'اللجان': {
    headers: ['المعرف', 'الاسم', 'الأعضاء', 'الحالة'],
    tabColor: '#ff9800'
  },
  'المواد': {
    headers: ['المعرف', 'اسم المادة', 'تاريخ الإضافة'],
    tabColor: '#607d8b'
  },
  'روابط_المعلمين': {
    headers: ['الجوال', 'الاسم', 'النوع', 'المرحلة', 'الفصول', 'تم الربط بواسطة', 'تاريخ الربط'],
    tabColor: '#795548'
  },
  'إعدادات_واتساب': {
    headers: ['الإعداد', 'القيمة', 'الوصف'],
    tabColor: '#4caf50'
  },
  'جلسات_واتساب': {
    headers: ['رقم_الواتساب', 'المرحلة', 'نوع_المستخدم', 'حالة_الاتصال', 'تاريخ_الربط', 'آخر_استخدام', 'عدد_الرسائل', 'الرقم_الرئيسي'],
    tabColor: '#4caf50'
  },
  'سجل_النشاطات': {
    headers: ['التاريخ', 'الوقت', 'المستخدم', 'النوع', 'التفاصيل', 'العدد', 'المرحلة'],
    tabColor: '#546e7a'
  },
  'أنواع_الملاحظات_التربوية': {
    headers: ['المرحلة', 'نوع الملاحظة', 'تاريخ الإضافة'],
    tabColor: '#1e3a5f'
  },
  'اعذار_اولياء_الامور': {
    headers: ['رقم_الطالب', 'اسم_الطالب', 'الصف', 'الفصل', 'المرحلة', 'نص_العذر', 'مرفقات', 'تاريخ_الغياب', 'تاريخ_التقديم', 'وقت_التقديم', 'الحالة', 'ملاحظات_المدرسة', 'الرمز'],
    tabColor: '#7c3aed'
  },
  'رموز_اولياء_الامور': {
    headers: ['الرمز', 'رقم_الطالب', 'المرحلة', 'تاريخ_الإنشاء', 'تاريخ_الانتهاء', 'مستخدم'],
    tabColor: '#7c3aed'
  }
};

// =================================================================
// ★ دالة بناء اسم الشيت من السجل المركزي
// getSheetName_('التأخر', 'ثانوي') → 'سجل_التأخر_ثانوي'
// =================================================================
function getSheetName_(type, stage) {
  var entry = SHEET_REGISTRY[type];
  if (!entry) {
    Logger.log('⚠️ نوع شيت غير معروف: ' + type);
    return 'سجل_' + type + '_' + stage;
  }
  return entry.prefix + '_' + stage;
}

// =================================================================
// ★ خريطة الأسماء البديلة الشاملة (للتوافق مع أي تسمية قديمة)
// المفتاح = الاسم المعتمد (الجديد)
// القيمة = كل الأسماء القديمة المحتملة
// =================================================================
var SHEET_ALIASES = {
  // ── المخالفات ──
  'سجل_المخالفات_متوسط': [
    'سجل المخالفات متوسط',           // الاسم الفعلي الحالي في الجدول
    'سجل_مخالفات_(المتوسط)',          // الاسم القديم في الكود
    'سجل مخالفات (المتوسط)',
    'سجل_مخالفات_المتوسط',
    'سجل مخالفات المتوسط'
  ],
  'سجل_المخالفات_ثانوي': [
    'سجل المخالفات ثانوي',            // الاسم الفعلي الحالي في الجدول
    'سجل_مخالفات_(الثانوي)',           // الاسم القديم في الكود
    'سجل مخالفات (الثانوي)',
    'سجل_مخالفات_الثانوي',
    'سجل مخالفات الثانوي'
  ],
  
  // ── الغياب العام ──
  'سجل_الغياب_متوسط': ['سجل الغياب متوسط', 'سجل_غياب_متوسط', 'سجل غياب متوسط'],
  'سجل_الغياب_ثانوي': ['سجل الغياب ثانوي', 'سجل_غياب_ثانوي', 'سجل غياب ثانوي'],
  
  // ── الغياب اليومي ──
  'سجل_الغياب_اليومي_متوسط': ['سجل الغياب اليومي متوسط', 'سجل_غياب_يومي_متوسط'],
  'سجل_الغياب_اليومي_ثانوي': ['سجل الغياب اليومي ثانوي', 'سجل_غياب_يومي_ثانوي'],
  
  // ── التأخر ──
  'سجل_التأخر_متوسط': ['سجل التأخر متوسط', 'سجل_تأخر_متوسط', 'سجل تأخر متوسط', 'متأخرين متوسط'],
  'سجل_التأخر_ثانوي': ['سجل التأخر ثانوي', 'سجل_تأخر_ثانوي', 'سجل تأخر ثانوي', 'متأخرين ثانوي'],
  
  // ── الاستئذان ──
  'سجل_الاستئذان_متوسط': ['سجل الاستئذان متوسط', 'سجل_استئذان_متوسط', 'سجل استئذان متوسط', 'استئذان متوسط'],
  'سجل_الاستئذان_ثانوي': ['سجل الاستئذان ثانوي', 'سجل_استئذان_ثانوي', 'سجل استئذان ثانوي', 'استئذان ثانوي'],
  
  // ── الملاحظات التربوية ──
  'سجل_الملاحظات_التربوية_متوسط': [
    'سجل الملاحظات التربوية متوسط',   // الاسم الفعلي الحالي في الجدول
    'سجل_ملاحظات_تربوية_متوسط',
    'سجل ملاحظات تربوية متوسط'
  ],
  'سجل_الملاحظات_التربوية_ثانوي': [
    'سجل الملاحظات التربوية ثانوي',    // الاسم الفعلي الحالي في الجدول
    'سجل_ملاحظات_تربوية_ثانوي',
    'سجل ملاحظات تربوية ثانوي'
  ],
  
  // ── التواصل ──
  'سجل_التواصل_متوسط': ['سجل التواصل متوسط', 'سجل_التواصل', 'سجل التواصل'],
  'سجل_التواصل_ثانوي': ['سجل التواصل ثانوي'],
  
  // ── التحصيل ──
  'سجل_التحصيل_متوسط': ['سجل التحصيل متوسط', 'التحصيل الدراسي متوسط'],
  'سجل_التحصيل_ثانوي': ['سجل التحصيل ثانوي', 'التحصيل الدراسي ثانوي'],
  
  // ── السلوك الإيجابي ──
  'سجل_السلوك_الإيجابي_متوسط': ['سجل السلوك الإيجابي متوسط', 'سجل_سلوك_إيجابي_متوسط', 'سجل سلوك إيجابي متوسط'],
  'سجل_السلوك_الإيجابي_ثانوي': ['سجل السلوك الإيجابي ثانوي', 'سجل_سلوك_إيجابي_ثانوي', 'سجل سلوك إيجابي ثانوي'],
  'سجل_السلوك_الإيجابي_ابتدائي': ['سجل السلوك الإيجابي ابتدائي', 'سجل_سلوك_إيجابي_ابتدائي', 'سجل سلوك إيجابي ابتدائي'],
  
  // ── الإعدادات (ثوابت) ──
  'إعدادات_المدرسة': ['الإعدادات', 'إعدادات المدرسة', 'اعدادات_المدرسة']
};

// =================================================================
// ★ دالة تشخيص الشيتات - شغّلها للتحقق من صحة كل شيء
// =================================================================
function diagnoseSheets_() {
  ensureStudentsSheetsLoaded_();
  var ss = getSpreadsheet_();
  var allSheets = ss.getSheets().map(function(s) { return s.getName(); });
  
  Logger.log('═══════════════════════════════════════');
  Logger.log('★ تشخيص الشيتات - ' + new Date().toLocaleString('ar-SA'));
  Logger.log('═══════════════════════════════════════');
  Logger.log('إجمالي الشيتات في الجدول: ' + allSheets.length);
  Logger.log('');
  
  // 1. فحص شيتات السجلات (لكل مرحلة)
  var stages = Object.keys(STUDENTS_SHEETS);
  var found = 0, missing = 0, viaAlias = 0;
  
  var types = Object.keys(SHEET_REGISTRY);
  for (var t = 0; t < types.length; t++) {
    var type = types[t];
    var entry = SHEET_REGISTRY[type];
    if (!entry.perStage) continue;
    
    for (var s = 0; s < stages.length; s++) {
      var stage = stages[s];
      var canonical = entry.prefix + '_' + stage;
      
      // بحث مباشر
      if (allSheets.indexOf(canonical) >= 0) {
        Logger.log('✅ ' + canonical);
        found++;
        continue;
      }
      
      // بحث بالأسماء البديلة
      var aliases = SHEET_ALIASES[canonical] || [];
      var foundAlias = false;
      for (var a = 0; a < aliases.length; a++) {
        if (allSheets.indexOf(aliases[a]) >= 0) {
          Logger.log('🔄 ' + canonical + ' ← وُجد كـ "' + aliases[a] + '" (يحتاج إعادة تسمية)');
          viaAlias++;
          foundAlias = true;
          break;
        }
      }
      
      if (!foundAlias) {
        Logger.log('❌ ' + canonical + ' ← غير موجود (سيُنشأ تلقائياً عند الحاجة)');
        missing++;
      }
    }
  }
  
  // 2. فحص شيتات النظام
  Logger.log('');
  Logger.log('── شيتات النظام ──');
  var systemSheets = [SCHOOL_SETTINGS_SHEET, SCHOOL_STRUCTURE_SHEET, USERS_SHEET, TEACHERS_SHEET, 'روابط_المعلمين', 'المواد', 'إعدادات_واتساب', 'جلسات_واتساب', 'سجل_النشاطات', 'أنواع_الملاحظات_التربوية', 'اعذار_اولياء_الامور', 'رموز_اولياء_الامور'];
  for (var i = 0; i < systemSheets.length; i++) {
    var name = systemSheets[i];
    if (allSheets.indexOf(name) >= 0) {
      Logger.log('✅ ' + name);
      found++;
    } else {
      Logger.log('❌ ' + name + ' ← غير موجود!');
      missing++;
    }
  }
  
  // 3. فحص شيتات الطلاب
  Logger.log('');
  Logger.log('── شيتات الطلاب ──');
  for (var st in STUDENTS_SHEETS) {
    var sn = STUDENTS_SHEETS[st];
    if (allSheets.indexOf(sn) >= 0) {
      Logger.log('✅ ' + sn + ' (' + st + ')');
      found++;
    } else {
      Logger.log('⚪ ' + sn + ' (' + st + ') ← غير موجود (اختياري)');
    }
  }
  
  // 4. شيتات غير معروفة (ليست في السجل)
  Logger.log('');
  Logger.log('── شيتات غير مسجلة ──');
  var knownNames = [];
  // جمع كل الأسماء المعروفة
  for (var key in SHEET_ALIASES) {
    knownNames.push(key);
    var al = SHEET_ALIASES[key];
    for (var j = 0; j < al.length; j++) knownNames.push(al[j]);
  }
  for (var st2 in STUDENTS_SHEETS) knownNames.push(STUDENTS_SHEETS[st2]);
  systemSheets.forEach(function(n) { knownNames.push(n); });
  
  var unknown = allSheets.filter(function(n) { return knownNames.indexOf(n) < 0; });
  if (unknown.length > 0) {
    unknown.forEach(function(n) {
      Logger.log('⚠️ ' + n + ' ← شيت غير مسجل (قديم؟ يدوي؟)');
    });
  } else {
    Logger.log('لا يوجد شيتات غير معروفة ✅');
  }
  
  // ملخص
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
  Logger.log('✅ موجود: ' + found + ' | 🔄 بديل: ' + viaAlias + ' | ❌ مفقود: ' + missing);
  Logger.log('═══════════════════════════════════════');
}

// =================================================================
// دالة البحث الذكي عن الشيت
// تبحث بالاسم الأساسي أولاً، ثم الأسماء البديلة
// =================================================================
function findSheet_(ss, sheetName) {
  // 1. البحث بالاسم الأساسي (الجديد)
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) return sheet;
  
  // 2. البحث في الأسماء البديلة (القديمة)
  var aliases = SHEET_ALIASES[sheetName];
  if (aliases) {
    for (var i = 0; i < aliases.length; i++) {
      sheet = ss.getSheetByName(aliases[i]);
      if (sheet) return sheet;
    }
  }
  
  // 3. لم يُعثر على الشيت
  return null;
}

// =================================================================
// دالة البحث المرن عن عمود في الأعمدة
// تبحث بالاسم بالشرطة السفلية والمسافة
// =================================================================
function findHeaderIndex_(headers, colName) {
  var idx = headers.indexOf(colName);
  if (idx >= 0) return idx;
  
  // محاولة بالصيغة البديلة (شرطة ↔ مسافة)
  var alt = colName.indexOf('_') >= 0 ? colName.replace(/_/g, ' ') : colName.replace(/\s+/g, '_');
  idx = headers.indexOf(alt);
  if (idx >= 0) return idx;
  
  return -1;
}

// =================================================================
// دالة استخراج قيمة التاريخ من صف (تبحث في كلا العمودين المكررين)
// =================================================================
function getRowDateValue_(row, headers, colName) {
  // البحث بالاسم الأصلي
  var idx = findHeaderIndex_(headers, colName);
  if (idx >= 0 && row[idx]) return row[idx];
  
  // البحث بالصيغة البديلة (لو العمود الأول فارغ والثاني فيه قيمة)
  var alt = colName.indexOf('_') >= 0 ? colName.replace(/_/g, ' ') : colName.replace(/\s+/g, '_');
  var altIdx = headers.indexOf(alt);
  if (altIdx >= 0 && altIdx !== idx && row[altIdx]) return row[altIdx];
  
  return null;
}

// =================================================================
// دالة فحص: هل القيمة تمثل تاريخ اليوم؟
// تتعامل مع: Date object، نص تاريخ، فارغ
// فارغ = ليس اليوم (يُستبعد)
// =================================================================
function isTodayDate_(value) {
  if (!value && value !== 0) return false; // فارغ → ليس اليوم
  
  var today = new Date();
  var tz = Session.getScriptTimeZone();
  var todayStr = Utilities.formatDate(today, tz, 'yyyy-MM-dd');
  
  // كائن تاريخ حقيقي
  if (value instanceof Date) {
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd') === todayStr;
  }
  
  // نص تاريخ
  if (typeof value === 'string' && value.length > 0) {
    // صيغة yyyy/MM/dd أو yyyy-MM-dd
    var parts = value.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (parts) {
      var dateStr = parts[1] + '-' + ('0' + parts[2]).slice(-2) + '-' + ('0' + parts[3]).slice(-2);
      return dateStr === todayStr;
    }
    // محاولة تحويل عام
    try {
      var d = new Date(value);
      if (!isNaN(d.getTime())) {
        return Utilities.formatDate(d, tz, 'yyyy-MM-dd') === todayStr;
      }
    } catch(e) {}
  }
  
  return false; // لا يمكن تحليله → ليس اليوم
}

// =================================================================
// دالة فلترة سجلات اليوم (مركزية لكل الوحدات)
// تستقبل: data (مصفوفة كاملة)، headers (صف العناوين)
// ترجع: مصفوفة سجلات اليوم فقط مع توحيد الأعمدة
// =================================================================
function filterTodayRecords_(data, headers, dateColName) {
  var records = [];
  
  // بناء خريطة الأعمدة الموحدة (حذف المكررات)
  var headerMap = {};
  var seen = {};
  for (var h = 0; h < headers.length; h++) {
    var standardized = String(headers[h] || '').trim().replace(/\s+/g, '_');
    if (!standardized || seen[standardized]) continue;
    seen[standardized] = true;
    headerMap[h] = standardized;
  }
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] && !row[1]) continue; // صف فارغ
    
    // استخراج التاريخ (من أي عمود متاح)
    var dateValue = getRowDateValue_(row, headers, dateColName);
    
    // الفلتر: فقط سجلات اليوم (الفارغ يُستبعد)
    if (!isTodayDate_(dateValue)) continue;
    
    // بناء السجل بأعمدة موحدة
    var record = { rowIndex: i };
    for (var j in headerMap) {
      var value = row[j];
      if (value instanceof Date) {
        if (headerMap[j] === 'وقت_الخروج' || headerMap[j] === 'وقت_التأكيد') {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
        } else {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
        }
      }
      record[headerMap[j]] = String(value || '');
    }
    records.push(record);
  }
  
  return records;
}

// =================================================================
// ★ تحويل الأرقام العربية إلى غربية وتنظيف صيغة التاريخ الهجري
// =================================================================
function arabicToWesternNumerals_(str) {
  if (!str) return str;
  var arabicMap = {'٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9'};
  var result = String(str);
  for (var k in arabicMap) {
    result = result.split(k).join(arabicMap[k]);
  }
  // إزالة "هـ" والمسافات الزائدة وحروف Unicode غير المرئية
  result = result.replace(/\s*هـ\s*/g, '').replace(/[\u200e\u200f\u200b\u200c\u200d\u2066\u2067\u2068\u2069\u061c]/g, '').trim();
  return result;
}

// =================================================================
// ★ دالة التاريخ الهجري المركزية (المصدر الوحيد)
// المصدر: API تقويم أم القرى (aladhan.com) مع كاش 24 ساعة
// Fallback: حساب محلي مع تعديل HIJRI_OFFSET
// =================================================================
function getHijriDate_(date) {
  try {
    var d = date || new Date();
    var dateKey = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MM-yyyy');
    
    // ١. محاولة الكاش أولاً (سريع)
    var cache = CacheService.getScriptCache();
    var cached = cache.get('hijri_' + dateKey);
    if (cached) return cached;
    
    // ٢. محاولة API أم القرى
    try {
      var url = 'https://api.aladhan.com/v1/gToH/' + dateKey + '?calendarMethod=umm-al-qura';
      var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      if (response.getResponseCode() === 200) {
        var json = JSON.parse(response.getContentText());
        if (json.code === 200 && json.data && json.data.hijri) {
          var h = json.data.hijri;
          var hijriStr = h.day + '/' + h.month.number + '/' + h.year;
          // حفظ في الكاش لـ 6 ساعات (21600 ثانية)
          cache.put('hijri_' + dateKey, hijriStr, 21600);
          return hijriStr;
        }
      }
    } catch(apiErr) {
      Logger.log('Hijri API error (using fallback): ' + apiErr);
    }
    
    // ٣. Fallback: الحساب المحلي مع التعديل
    var offset = getHijriOffset_();
    var fallbackDate = new Date(d.getTime());
    if (offset !== 0) fallbackDate.setDate(fallbackDate.getDate() + offset);
    var fallbackStr = fallbackDate.toLocaleDateString('ar-SA-u-ca-islamic', {
      day: 'numeric', month: 'numeric', year: 'numeric'
    });
    return arabicToWesternNumerals_(fallbackStr);
  } catch (e) {
    return Utilities.formatDate(date || new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  }
}

// ★ جلب التاريخ الهجري الكامل (مع اسم الشهر) — للداشبورد
function getHijriDateFull_(date) {
  try {
    var d = date || new Date();
    var dateKey = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MM-yyyy');
    
    var cache = CacheService.getScriptCache();
    var cachedFull = cache.get('hijri_full_' + dateKey);
    if (cachedFull) return JSON.parse(cachedFull);
    
    try {
      var url = 'https://api.aladhan.com/v1/gToH/' + dateKey + '?calendarMethod=umm-al-qura';
      var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      if (response.getResponseCode() === 200) {
        var json = JSON.parse(response.getContentText());
        if (json.code === 200 && json.data && json.data.hijri) {
          var h = json.data.hijri;
          var g = json.data.gregorian;
          var result = {
            hijriDay: h.day,
            hijriMonth: h.month.ar || h.month.en,
            hijriMonthNum: h.month.number,
            hijriYear: h.year,
            hijriStr: h.day + ' ' + (h.month.ar || h.month.en) + ' ' + h.year + ' هـ',
            gregorianStr: (g ? g.day + ' ' + (g.month.en || '') + ' ' + g.year : ''),
            weekdayAr: h.weekday ? h.weekday.ar : ''
          };
          cache.put('hijri_full_' + dateKey, JSON.stringify(result), 21600);
          return result;
        }
      }
    } catch(e) {}
    
    // Fallback
    var offset = getHijriOffset_();
    var fd = new Date(d.getTime());
    if (offset !== 0) fd.setDate(fd.getDate() + offset);
    return {
      hijriStr: fd.toLocaleDateString('ar-SA-u-ca-islamic', {day:'numeric',month:'long',year:'numeric'}),
      gregorianStr: fd.toLocaleDateString('ar-EG', {day:'numeric',month:'long',year:'numeric'}),
      weekdayAr: ''
    };
  } catch(e) { return {hijriStr: '', gregorianStr: '', weekdayAr: ''}; }
}

// ★ جلب قيمة التعديل الهجري من الإعدادات
function getHijriOffset_() {
  try {
    var val = PropertiesService.getScriptProperties().getProperty('HIJRI_OFFSET');
    return val ? parseInt(val) : 0;
  } catch(e) { return 0; }
}

// ★ تعيين التعديل الهجري (يُستخدم فقط كـ fallback)
function setHijriOffset(offset) {
  PropertiesService.getScriptProperties().setProperty('HIJRI_OFFSET', String(offset || 0));
  return { success: true, offset: offset, message: 'تم تعيين تعديل التاريخ الهجري: ' + offset };
}

// ★ مسح كاش التاريخ الهجري (لإجبار إعادة الجلب من API)
function clearHijriCache() {
  var cache = CacheService.getScriptCache();
  var now = new Date();
  var dateKey = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  cache.remove('hijri_' + dateKey);
  cache.remove('hijri_full_' + dateKey);
  return {success: true, message: 'تم مسح كاش التاريخ الهجري'};
}

// =================================================================
// ★ اسم اليوم بالعربي (مركزية)
// =================================================================
function getDayNameAr_(date) {
  var days = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
  return days[(date || new Date()).getDay()];
}

// =================================================================
// ★★★ نظام إدارة الشيتات التلقائي ★★★
// =================================================================

// =================================================================
// تعريفات الشيتات المركزية (الترويسات + الألوان + العرض)
// هذا المصدر الوحيد للحقيقة - كل شيء يُبنى من هنا
// =================================================================
var SHEET_DEFINITIONS = {
  'المخالفات': {
    // ★ 18 عمود موحد (اليوم + تم الإرسال، بدون ملاحظات)
    headers: ['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'رقم المخالفة', 'نص المخالفة', 'نوع المخالفة', 'الدرجة', 'التاريخ الهجري', 'التاريخ الميلادي', 'مستوى التكرار', 'الإجراءات', 'النقاط', 'اليوم', 'النماذج المحفوظة', 'المستخدم', 'وقت الإدخال', 'تم الإرسال'],
    headerBg: '#e74c3c',
    headerColor: '#ffffff',
    tabColors: { 'متوسط': '#4285f4', 'ثانوي': '#0f9d58', 'ابتدائي': '#f4b400', 'طفولة مبكرة': '#e91e63' },
    widths: [100, 150, 120, 60, 80, 200, 80, 50, 100, 100, 60, 200, 50, 80, 100, 100, 120, 80]
  },
  'التأخر': {
    // ★ 11 عمود - مطابق لـ getLateSheet() في Server_Attendance.gs
    headers: ['رقم_الطالب', 'اسم_الطالب', 'الصف', 'الفصل', 'رقم_الجوال', 'نوع_التأخر', 'الحصة', 'التاريخ_هجري', 'المسجل', 'وقت_الإدخال', 'تم_الإرسال'],
    headerBg: '#dc2626',
    headerColor: '#ffffff',
    tabColors: { 'متوسط': '#4285f4', 'ثانوي': '#0f9d58', 'ابتدائي': '#f4b400', 'طفولة مبكرة': '#e91e63' },
    widths: [100, 150, 80, 60, 110, 100, 60, 100, 100, 120, 80]
  },
  'الاستئذان': {
    // ★ 14 عمود - مطابق لـ getPermissionSheet() في Server_Attendance.gs
    headers: ['رقم_الطالب', 'اسم_الطالب', 'الصف', 'الفصل', 'رقم_الجوال', 'وقت_الخروج', 'السبب', 'المستلم', 'المسؤول', 'التاريخ_هجري', 'المسجل', 'وقت_الإدخال', 'وقت_التأكيد', 'تم_الإرسال'],
    headerBg: '#7c3aed',
    headerColor: '#ffffff',
    tabColors: { 'متوسط': '#4285f4', 'ثانوي': '#0f9d58', 'ابتدائي': '#f4b400', 'طفولة مبكرة': '#e91e63' },
    widths: [100, 150, 80, 60, 110, 80, 150, 120, 100, 100, 100, 120, 80, 80]
  },
  'الغياب': {
    headers: ['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'غياب بعذر', 'غياب بدون عذر', 'تأخير', 'آخر تحديث'],
    headerBg: '#f59e0b',
    headerColor: '#ffffff',
    tabColors: { 'متوسط': '#4285f4', 'ثانوي': '#0f9d58', 'ابتدائي': '#f4b400', 'طفولة مبكرة': '#e91e63' },
    widths: [100, 150, 80, 60, 80, 80, 60, 120]
  },
  'الغياب_اليومي': {
    headers: ['رقم_الطالب', 'اسم_الطالب', 'الصف', 'الفصل', 'رقم_الجوال', 'نوع_الغياب', 'الحصة', 'التاريخ_هجري', 'اليوم', 'المسجل', 'وقت_الإدخال', 'حالة_الاعتماد', 'نوع_العذر', 'تم_الإرسال', 'حالة_التأخر', 'وقت_الحضور', 'ملاحظات'],
    headerBg: '#ea580c',
    headerColor: '#ffffff',
    tabColors: { 'متوسط': '#4285f4', 'ثانوي': '#0f9d58', 'ابتدائي': '#f4b400', 'طفولة مبكرة': '#e91e63' },
    widths: [100, 150, 80, 60, 110, 100, 60, 100, 80, 100, 120, 80, 80, 80, 80, 80, 150]
  },
  'الملاحظات_التربوية': {
    headers: ['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'رقم الجوال', 'نوع الملاحظة', 'التفاصيل', 'المعلم/المسجل', 'التاريخ', 'وقت الإدخال', 'تم الإرسال'],
    headerBg: '#1e3a5f',
    headerColor: '#ffffff',
    tabColors: { 'متوسط': '#4285f4', 'ثانوي': '#0f9d58', 'ابتدائي': '#f4b400', 'طفولة مبكرة': '#e91e63' },
    widths: [100, 150, 80, 60, 110, 150, 200, 100, 100, 100, 80]
  },
  'التواصل': {
    headers: ['م', 'التاريخ الهجري', 'التاريخ الميلادي', 'الوقت', 'رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'رقم الجوال', 'نوع الرسالة', 'عنوان الرسالة', 'نص الرسالة', 'حالة الإرسال', 'المرسل', 'ملاحظات'],
    headerBg: '#4a5568',
    headerColor: '#ffffff',
    tabColors: { 'متوسط': '#4285f4', 'ثانوي': '#0f9d58', 'ابتدائي': '#f4b400', 'طفولة مبكرة': '#e91e63' },
    widths: [40, 100, 100, 70, 100, 150, 80, 60, 110, 80, 150, 300, 80, 100, 150]
  },
  'التحصيل': {
    headers: ['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'المادة', 'نوع التقييم', 'الدرجة', 'من', 'التاريخ', 'المعلم', 'ملاحظات', 'وقت الإدخال'],
    headerBg: '#059669',
    headerColor: '#ffffff',
    tabColors: { 'متوسط': '#4285f4', 'ثانوي': '#0f9d58', 'ابتدائي': '#f4b400', 'طفولة مبكرة': '#e91e63' },
    widths: [100, 150, 80, 60, 100, 100, 60, 60, 100, 100, 150, 120]
  },
  'السلوك_الإيجابي': {
    // ★ 13 عمود - مطابق لـ getSheetHeaders_(positive) في Server_TeacherInput.gs
    headers: ['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'رقم الجوال', 'السلوك المتمايز', 'الدرجة', 'المعلم', 'اليوم', 'التاريخ الهجري', 'التاريخ الميلادي', 'وقت الإدخال', 'تم الإرسال'],
    headerBg: '#10b981',
    headerColor: '#ffffff',
    tabColors: { 'متوسط': '#4285f4', 'ثانوي': '#0f9d58', 'ابتدائي': '#f4b400', 'طفولة مبكرة': '#e91e63' },
    widths: [100, 150, 80, 60, 110, 200, 60, 120, 80, 100, 100, 120, 80]
  },
  'تحصيل_ملخص': {
    // ★ 14 عمود - مطابق لـ SUMMARY_HEADERS في Server_Academic.gs
    headers: ['رقم_الهوية', 'اسم_الطالب', 'الصف', 'الفصل', 'الفصل_الدراسي', 'الفترة', 'المعدل', 'التقدير_العام', 'ترتيب_الصف', 'ترتيب_الفصل', 'الغياب', 'التأخر', 'السلوك_متميز', 'السلوك_إيجابي'],
    headerBg: '#059669',
    headerColor: '#ffffff',
    tabColors: { 'متوسط': '#4285f4', 'ثانوي': '#0f9d58', 'ابتدائي': '#f4b400', 'طفولة مبكرة': '#e91e63' },
    widths: [110, 180, 100, 60, 80, 100, 70, 80, 70, 70, 60, 60, 80, 80]
  },
  'تحصيل_درجات': {
    // ★ 12 عمود - مطابق لـ GRADES_HEADERS في Server_Academic.gs
    headers: ['رقم_الهوية', 'اسم_الطالب', 'الصف', 'الفصل', 'الفصل_الدراسي', 'الفترة', 'المادة', 'المجموع', 'اختبار_نهائي', 'أدوات_تقييم', 'اختبارات_قصيرة', 'التقدير'],
    headerBg: '#0d9488',
    headerColor: '#ffffff',
    tabColors: { 'متوسط': '#4285f4', 'ثانوي': '#0f9d58', 'ابتدائي': '#f4b400', 'طفولة مبكرة': '#e91e63' },
    widths: [110, 180, 100, 60, 80, 100, 120, 70, 80, 80, 80, 70]
  }
};

// =================================================================
// ★ ensureAllSheets_ - تعمل عند كل فتح للتطبيق
// سريعة: تفحص الموجود فقط → تصلح أو تنشئ المفقود
// =================================================================
function ensureAllSheets_() {
  ensureStudentsSheetsLoaded_();
  var ss = getSpreadsheet_();

  // جمع أسماء الشيتات الموجودة
  var allSheets = ss.getSheets();
  var existingMap = {};
  for (var i = 0; i < allSheets.length; i++) {
    existingMap[allSheets[i].getName()] = allSheets[i];
  }

  // ★ إنشاء شيتات النظام الأساسية إذا كانت مفقودة
  var sysNames = Object.keys(SYSTEM_SHEETS_DEFINITIONS);
  for (var sn = 0; sn < sysNames.length; sn++) {
    var sysName = sysNames[sn];
    if (!existingMap[sysName]) {
      var sysDef = SYSTEM_SHEETS_DEFINITIONS[sysName];
      var sysSheet = ss.insertSheet(sysName);
      sysSheet.setRightToLeft(true);
      sysSheet.appendRow(sysDef.headers);
      sysSheet.getRange(1, 1, 1, sysDef.headers.length)
        .setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
      sysSheet.setFrozenRows(1);
      if (sysDef.tabColor) sysSheet.setTabColor(sysDef.tabColor);
      existingMap[sysName] = sysSheet;
      Logger.log('✅ إنشاء شيت نظام: ' + sysName);
    }
  }

  // تحديد المراحل النشطة من هيكل_المدرسة
  var activeStages = [];
  for (var stage in STUDENTS_SHEETS) {
    activeStages.push(stage);
    // إنشاء شيت الطلاب إذا لم يكن موجوداً
    if (!existingMap[STUDENTS_SHEETS[stage]]) {
      var studentSheet = ss.insertSheet(STUDENTS_SHEETS[stage]);
      studentSheet.setRightToLeft(true);
      studentSheet.appendRow(['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'رقم الجوال', 'تاريخ الإضافة']);
      studentSheet.getRange(1, 1, 1, 6)
        .setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
      studentSheet.setFrozenRows(1);
      existingMap[STUDENTS_SHEETS[stage]] = studentSheet;
      Logger.log('✅ إنشاء شيت طلاب: ' + STUDENTS_SHEETS[stage]);
    }
  }

  if (activeStages.length === 0) {
    Logger.log('⚠️ لا توجد مراحل نشطة (هيكل_المدرسة غير مُعدّ)');
    return { created: 0, renamed: 0, existed: 0 };
  }
  
  var created = 0, renamed = 0, existed = 0;
  
  // لكل نوع سجل × لكل مرحلة نشطة
  var types = Object.keys(SHEET_REGISTRY);
  for (var t = 0; t < types.length; t++) {
    var type = types[t];
    var reg = SHEET_REGISTRY[type];
    if (!reg.perStage) continue;
    
    for (var s = 0; s < activeStages.length; s++) {
      var stage = activeStages[s];
      var canonicalName = reg.prefix + '_' + stage;
      
      // 1. الشيت موجود بالاسم الصحيح؟
      if (existingMap[canonicalName]) {
        existed++;
        continue;
      }
      
      // 2. موجود باسم قديم؟ → أعد تسميته
      var foundViaAlias = false;
      var aliases = SHEET_ALIASES[canonicalName] || [];
      for (var a = 0; a < aliases.length; a++) {
        if (existingMap[aliases[a]]) {
          existingMap[aliases[a]].setName(canonicalName);
          Logger.log('🔄 إعادة تسمية: "' + aliases[a] + '" → "' + canonicalName + '"');
          renamed++;
          foundViaAlias = true;
          break;
        }
      }
      if (foundViaAlias) continue;
      
      // 3. غير موجود نهائياً → أنشئه
      var def = SHEET_DEFINITIONS[type];
      if (def) {
        createSheetFromDefinition_(ss, canonicalName, def, stage);
        created++;
        Logger.log('✅ إنشاء: ' + canonicalName);
      }
    }
  }
  
  Logger.log('');
  Logger.log('📊 النتيجة: ✅ موجود=' + existed + ' | 🔄 أُعيد تسميته=' + renamed + ' | ➕ أُنشئ=' + created);
  return { created: created, renamed: renamed, existed: existed };
}

// =================================================================
// ★ إنشاء شيت من التعريف المركزي
// =================================================================
function createSheetFromDefinition_(ss, sheetName, def, stage) {
  var sheet = ss.insertSheet(sheetName);
  sheet.setRightToLeft(true);
  
  // الترويسة
  sheet.appendRow(def.headers);
  sheet.getRange(1, 1, 1, def.headers.length)
    .setBackground(def.headerBg)
    .setFontColor(def.headerColor)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  
  // عرض الأعمدة
  if (def.widths) {
    for (var w = 0; w < Math.min(def.widths.length, def.headers.length); w++) {
      sheet.setColumnWidth(w + 1, def.widths[w]);
    }
  }
  
  // لون التبويب حسب المرحلة
  if (def.tabColors && def.tabColors[stage]) {
    sheet.setTabColor(def.tabColors[stage]);
  }
  
  return sheet;
}

// =================================================================
// ★ repairSheetHeaders_ — إصلاح أعمدة الشيتات الموجودة
// تضيف الأعمدة الناقصة في نهاية الشيت (آمنة — لا تحذف بيانات)
// شغّلها يدوياً بعد التحديث إذا كانت شيتات قديمة تحتاج أعمدة جديدة
// =================================================================
function repairSheetHeaders_() {
  var ss = getSpreadsheet_();
  var repaired = 0;
  var L = Logger;

  L.log('═══════════════════════════════════════');
  L.log('★ إصلاح أعمدة الشيتات — ' + new Date().toLocaleString('ar-SA'));
  L.log('═══════════════════════════════════════');

  function normalize(h) {
    return String(h || '').trim().replace(/[\s_]+/g, '_');
  }

  // (1) إصلاح شيتات النظام
  var sysNames = Object.keys(SYSTEM_SHEETS_DEFINITIONS);
  for (var sn = 0; sn < sysNames.length; sn++) {
    var sysName = sysNames[sn];
    var sheet = ss.getSheetByName(sysName);
    if (!sheet || sheet.getLastColumn() === 0) continue;

    var sysDef = SYSTEM_SHEETS_DEFINITIONS[sysName];
    var actualHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(function(h) { return String(h || '').trim(); });
    var actualNorm = actualHeaders.map(normalize);

    var missing = [];
    for (var i = 0; i < sysDef.headers.length; i++) {
      if (actualNorm.indexOf(normalize(sysDef.headers[i])) === -1) {
        missing.push(sysDef.headers[i]);
      }
    }

    if (missing.length > 0) {
      // إضافة الأعمدة الناقصة في نهاية الشيت
      var startCol = sheet.getLastColumn() + 1;
      for (var m = 0; m < missing.length; m++) {
        sheet.getRange(1, startCol + m).setValue(missing[m])
          .setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
      }
      L.log('🔧 ' + sysName + ': أُضيفت أعمدة: ' + missing.join(' | '));
      repaired++;
    }
  }

  // (2) إصلاح شيتات السجلات (حسب المراحل المفعّلة)
  ensureStudentsSheetsLoaded_();
  var activeStages = Object.keys(STUDENTS_SHEETS);
  var types = Object.keys(SHEET_DEFINITIONS);

  for (var t = 0; t < types.length; t++) {
    var type = types[t];
    var reg = SHEET_REGISTRY[type];
    if (!reg || !reg.perStage) continue;
    var def = SHEET_DEFINITIONS[type];
    if (!def) continue;

    for (var s = 0; s < activeStages.length; s++) {
      var stage = activeStages[s];
      var sheetName = reg.prefix + '_' + stage;
      var sheet = findSheet_(ss, sheetName);
      if (!sheet || sheet.getLastColumn() === 0) continue;

      var actualHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
        .map(function(h) { return String(h || '').trim(); });
      var actualNorm = actualHeaders.map(normalize);

      var missing = [];
      for (var i = 0; i < def.headers.length; i++) {
        if (actualNorm.indexOf(normalize(def.headers[i])) === -1) {
          missing.push(def.headers[i]);
        }
      }

      if (missing.length > 0) {
        var startCol = sheet.getLastColumn() + 1;
        for (var m = 0; m < missing.length; m++) {
          sheet.getRange(1, startCol + m).setValue(missing[m])
            .setBackground(def.headerBg || '#1e3a5f')
            .setFontColor(def.headerColor || '#ffffff')
            .setFontWeight('bold');
        }
        L.log('🔧 ' + sheetName + ': أُضيفت أعمدة: ' + missing.join(' | '));
        repaired++;
      }
    }
  }

  L.log('');
  L.log('═══════════════════════════════════════');
  L.log('✅ اكتمل الإصلاح — شيتات أُصلحت: ' + repaired);
  L.log('═══════════════════════════════════════');

  return { success: true, repaired: repaired };
}

// =================================================================
// ★ setupNewSchool - إعداد مدرسة جديدة من الصفر
// شغّلها مرة واحدة من محرر السكربت بعد إنشاء جوجل شيت فاضي
// تنشئ كل شيتات النظام الأساسية بالترويسات والتنسيق الصحيح
// =================================================================
function setupNewSchool() {
  var ss = getSpreadsheet_();
  var existingSheets = {};
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    existingSheets[sheets[i].getName()] = sheets[i];
  }

  var created = 0;
  var names = Object.keys(SYSTEM_SHEETS_DEFINITIONS);
  for (var n = 0; n < names.length; n++) {
    var name = names[n];
    if (existingSheets[name]) {
      Logger.log('⏭️ موجود: ' + name);
      continue;
    }
    var def = SYSTEM_SHEETS_DEFINITIONS[name];
    var sheet = ss.insertSheet(name);
    sheet.setRightToLeft(true);
    sheet.appendRow(def.headers);
    sheet.getRange(1, 1, 1, def.headers.length)
      .setBackground('#1e3a5f')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    if (def.tabColor) sheet.setTabColor(def.tabColor);
    created++;
    Logger.log('✅ إنشاء: ' + name);
  }

  // حذف الشيت الافتراضي (Sheet1 / ورقة1)
  var defaultNames = ['Sheet1', 'ورقة1'];
  for (var d = 0; d < defaultNames.length; d++) {
    var defSheet = existingSheets[defaultNames[d]];
    if (defSheet && ss.getSheets().length > 1) {
      ss.deleteSheet(defSheet);
      Logger.log('🗑️ حذف الشيت الافتراضي: ' + defaultNames[d]);
    }
  }

  Logger.log('');
  Logger.log('══════════════════════════════════════');
  Logger.log('✅ تم إعداد المدرسة الجديدة!');
  Logger.log('📊 شيتات أُنشئت: ' + created + ' من ' + names.length);
  Logger.log('');
  Logger.log('📋 الخطوات التالية:');
  Logger.log('   1. انشر كـ Web App (نشر ← عمليات نشر جديدة)');
  Logger.log('   2. افتح رابط التطبيق');
  Logger.log('   3. اختر المراحل الدراسية من شاشة الإعداد');
  Logger.log('   4. النظام سينشئ باقي الشيتات تلقائياً');
  Logger.log('══════════════════════════════════════');
}

// =================================================================
// ★ resetRecordSheets_ - حذف كل شيتات السجلات وإعادة إنشائها نظيفة
// ⚠️ تحذف البيانات! شغّلها يدوياً مرة واحدة فقط
// =================================================================
function resetRecordSheets_() {
  ensureStudentsSheetsLoaded_();
  var ss = getSpreadsheet_();
  var allSheets = ss.getSheets();
  
  // شيتات النظام المحمية (لا تُحذف أبداً)
  var protectedSheets = {};
  for (var stage in STUDENTS_SHEETS) {
    protectedSheets[STUDENTS_SHEETS[stage]] = true;
  }
  var systemNames = [
    SCHOOL_SETTINGS_SHEET, SCHOOL_STRUCTURE_SHEET, USERS_SHEET,
    TEACHERS_SHEET, 'روابط_المعلمين', 'المواد',
    'إعدادات_واتساب', 'جلسات_واتساب', COMMITTEES_SHEET,
    'سجل_النشاطات', 'أنواع_الملاحظات_التربوية',
    'اعذار_اولياء_الامور', 'رموز_اولياء_الامور'
  ];
  for (var i = 0; i < systemNames.length; i++) {
    protectedSheets[systemNames[i]] = true;
  }
  
  // حذف كل شيتات السجلات
  var deleted = 0;
  Logger.log('═══════════════════════════════════════');
  Logger.log('★ بدء إعادة تهيئة شيتات السجلات');
  Logger.log('═══════════════════════════════════════');
  
  // نحتاج نحافظ على شيت واحد على الأقل (شرط Google)
  // نبدأ بعد الشيتات المحمية
  for (var i = allSheets.length - 1; i >= 0; i--) {
    var name = allSheets[i].getName();
    
    // لا تحذف المحمية
    if (protectedSheets[name]) {
      Logger.log('🔒 محمي: ' + name);
      continue;
    }
    
    // احذف شيتات السجلات + أي شيت غير معروف
    // لكن أبقِ شيت واحد على الأقل
    if (ss.getSheets().length <= 1) {
      Logger.log('⚠️ لا يمكن حذف آخر شيت: ' + name);
      continue;
    }
    
    Logger.log('🗑️ حذف: ' + name);
    ss.deleteSheet(allSheets[i]);
    deleted++;
  }
  
  Logger.log('');
  Logger.log('🗑️ تم حذف ' + deleted + ' شيت');
  Logger.log('');
  
  // إعادة إنشاء الكل نظيف
  Logger.log('★ إعادة إنشاء الشيتات...');
  var result = ensureAllSheets_();
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
  Logger.log('✅ اكتملت التهيئة!');
  Logger.log('   حُذف: ' + deleted + ' | أُنشئ: ' + result.created);
  Logger.log('═══════════════════════════════════════');
  
  return { deleted: deleted, created: result.created };
}

// =================================================================
// ★ دوال عامة للتشغيل من الـ dropdown في محرر السكربت
// (الدوال الأصلية تنتهي بـ _ فلا تظهر في القائمة)
// =================================================================

/** إنشاء الشيتات المفقودة — شغّلها من dropdown المحرر */
function runEnsureAllSheets() {
  var result = ensureAllSheets_();
  Logger.log('✅ النتيجة: أُنشئ ' + result.created + ' شيت');
  return result;
}

/** إصلاح ترويسات الشيتات الموجودة — شغّلها من dropdown المحرر */
function runRepairSheetHeaders() {
  var result = repairSheetHeaders_();
  Logger.log('✅ النتيجة: أُصلح ' + result.repaired + ' شيت');
  return result;
}

// ═══════════════════════════════════════════════════════════════════════════
// ★ NOOR_DROPDOWN_MAP — خريطة ربط التطبيق بمنسدلات نور
// مرجع واحد لجميع النصوص والقيم — مستخرج من صفحة نور الفعلية
// ═══════════════════════════════════════════════════════════════════════════

var NOOR_DROPDOWN_MAP = {

  // ═══ أوضاع نور (mowadaba + deductType) ═══
  modes: {
    absence:      { mowadaba: '2', deductType: null },
    violation:    { mowadaba: '1', deductType: '1' },
    tardiness:    { mowadaba: '1', deductType: '1' },
    compensation: { mowadaba: '1', deductType: '2' },
    excellent:    { mowadaba: '1', deductType: '2' }
  },

  // ═══════════════════════════════════════════════════════════════════════
  //  الغياب — ddlDegreeDeductAmount عند مواظبة=2
  //  المفتاح = نص التطبيق (نوع_الغياب)، القيمة = نص ورقم نور
  // ═══════════════════════════════════════════════════════════════════════
  absence: {
    'غائب':                              { noorText: 'الغياب بدون عذر مقبول',                         noorValue: '48,' },
    'غياب بدون عذر':                     { noorText: 'الغياب بدون عذر مقبول',                         noorValue: '48,' },
    'الغياب بدون عذر مقبول':             { noorText: 'الغياب بدون عذر مقبول',                         noorValue: '48,' },
    'غياب بعذر':                         { noorText: 'الغياب بعذر',                                   noorValue: '141,' },
    'الغياب بعذر':                       { noorText: 'الغياب بعذر',                                   noorValue: '141,' },
    'غياب منصة بعذر':                    { noorText: 'الغياب بعذر مقبول عبر منصة مدرستي',              noorValue: '800667,' },
    'الغياب بعذر مقبول عبر منصة مدرستي': { noorText: 'الغياب بعذر مقبول عبر منصة مدرستي',              noorValue: '800667,' },
    'غياب منصة بدون عذر':                { noorText: 'الغياب بدون عذر مقبول عبر منصة مدرستي',          noorValue: '1201153,' },
    'الغياب بدون عذر مقبول عبر منصة مدرستي': { noorText: 'الغياب بدون عذر مقبول عبر منصة مدرستي',      noorValue: '1201153,' }
  },

  // ═══════════════════════════════════════════════════════════════════════
  //  المخالفات — ddlDegreeDeductAmount عند سلوك=1, مخالفة=1
  //  المفتاح = رقم المخالفة في التطبيق (من getRulesData_)
  //  ★ الدرجات هنا لمرحلة ابتدائي (degree_ابتدائي عند وجوده)
  // ═══════════════════════════════════════════════════════════════════════
  violations: {
    // ── الدرجة الأولى (ابتدائي) ──
    101: { noorText: 'التأخر الصباحي.',                                                                                           noorValue: '1601174,الدرجة الأولى' },
    102: { noorText: 'عدم حضور الاصطفاف الصباحي ( في حال كان الطالب متواجدا داخل المدرسة ).',                                      noorValue: '1201074,الدرجة الأولى' },
    103: { noorText: 'التأخر عن الاصطفاف الصباحي ( في حال كان الطالب متواجدا داخل المدرسة) أو العبث أثناءه.',                      noorValue: '1201099,الدرجة الأولى' },
    104: { noorText: 'التأخر في الدخول إلى الحصص.',                                                                               noorValue: '1601175,الدرجة الأولى' },
    106: { noorText: 'النوم داخل الفصل.',                                                                                         noorValue: '1201075,الدرجة الأولى' },
    107: { noorText: 'تكرار خروج ودخول الطلبة من البوابة قبل وقت الحضور والانصراف.',                                               noorValue: '1201101,الدرجة الأولى' },
    108: { noorText: 'التجمهر أمام بوابة المدرسة.',                                                                               noorValue: '1201077,الدرجة الأولى' },
    109: { noorText: 'تناول الأطعمة أو المشروبات أثناء الدرس بدون استئذان.',                                                      noorValue: '1601176,الدرجة الأولى' },
    301: { noorText: 'عدم التقيد بالزي المدرسي.',                                                                                 noorValue: '1601186,الدرجة الأولى' },

    // ── الدرجة الثانية (ابتدائي) ──
    201: { noorText: 'عدم حضور الحصة الدراسية أو الهروب منها.',                                                                   noorValue: '1201081,الدرجة الثانية' },
    202: { noorText: 'الدخول أو الخـروج مـن الفصـل دون اسـتئذان',                                                                 noorValue: '1601190,الدرجة الثانية' },
    203: { noorText: 'دخول فصل آخر دون استئذان.',                                                                                 noorValue: '1601177,الدرجة الثانية' },
    204: { noorText: 'إثارة الفوضى داخل الفصل أو المدرسة، أو في وسائل النقل المدرسي.',                                            noorValue: '1201102,الدرجة الثانية' },
    302: { noorText: 'الشجار أو الاشتراك في مضاربة جماعية.',                                                                      noorValue: '1201106,الدرجة الثانية' },
    303: { noorText: 'الإشارة بحركات مخلة بالأدب تجاه الطلبة.',                                                                   noorValue: '1201105,الدرجة الثانية' },
    304: { noorText: 'التلفظ بكلمات نابية على الطلبة، أو تهديدهم، أو السخرية منهم.',                                               noorValue: '1201080,الدرجة الثانية' },
    305: { noorText: 'إلحاق الضرر المتعمد بممتلكات الطلبة.',                                                                      noorValue: '1201079,الدرجة الثانية' },
    306: { noorText: 'العبث بتجهيزات المدرسة أو مبانيها (كأجهزة الحاسوب، أدوات ومعدات الأمن والسلامة المدرسية ، الكهرباء ، المعامل ،حافلة المدرسة، والكتابة على الجدار وغيره).', noorValue: '1201082,الدرجة الثانية' },
    311: { noorText: 'امتهان الكتب الدراسية.',                                                                                    noorValue: '1201108,الدرجة الثانية' },

    // ── الدرجة الثالثة (ابتدائي) ──
    310: { noorText: 'التوقيع عن ولي الأمر من غير علمه على المكاتبات المتبادلة بين المدرسة وولي الأمر.',                            noorValue: '1201089,الدرجة الثالثة' },
    401: { noorText: 'التعرض لأحــد الطلبة بالضــرب.',                                                                            noorValue: '1601213,الدرجة الثالثة' },
    402: { noorText: 'سرقة شيء من ممتلكات الطلبة أو المدرسة.',                                                                    noorValue: '1601188,الدرجة الثالثة' },
    403: { noorText: 'التصوير أو التسجيل الصوتي للطلبة.',                                                                         noorValue: '1201119,الدرجة الثالثة' },
    404: { noorText: 'إلحاق الضرر المتعمد بتجهيزات المدرسة أو مبانيها (كأجهزة الحاسوب ، أدوات ومعدات الأمن والسلامة المدرسية، الكهرباء، المعامل، الحافلة المدرسية).', noorValue: '1201084,الدرجة الثالثة' },
    406: { noorText: 'الهروب من المدرسة.',                                                                                        noorValue: '1201092,الدرجة الثالثة' },
    407: { noorText: 'إحضار أو استخدام المواد أو الألعاب الخطرة إلى المدرسة، مثل (الألعاب النارية، البخاخات الغازية الملونة، المواد الكيميائية).', noorValue: '1601184,الدرجة الثالثة' },

    // ── الدرجة الرابعة (ابتدائي) ──
    308: { noorText: 'حيازة السجائر بأنواعها.',                                                                                   noorValue: '1201097,الدرجة الرابعة' },
    309: { noorText: 'حيازة أو عرض المواد الإعلامية الممنوعة المقروءة، أو المسموعة ، أو المرئية.',                                 noorValue: '1601189,الدرجة الرابعة' },
    405: { noorText: 'التدخين بأنواعه داخل المدرسة.',                                                                             noorValue: '1201114,الدرجة الرابعة' },
    501: { noorText: 'الإساءة أو الاستهزاء بشيء من شعائر الإسلام.',                                                               noorValue: '1201121,الدرجة الرابعة' },
    502: { noorText: 'الإساءة للدولة أو رموزها.',                                                                                 noorValue: '1201122,الدرجة الرابعة' },
    506: { noorText: 'التحرش الجنسي.',                                                                                            noorValue: '1201095,الدرجة الرابعة' },
    507: { noorText: 'المظاهر أو الصور أو الشعارات التي تدل على الشذوذ الجنسي أو الترويج لها.',                                    noorValue: '1601179,الدرجة الرابعة' },
    508: { noorText: 'إشعال النار داخل المدرسة.',                                                                                 noorValue: '1201096,الدرجة الرابعة' },
    509: { noorText: 'حيازة آلة حادة ( مثل السكاكين).',                                                                           noorValue: '1601178,الدرجة الرابعة' },
    511: { noorText: 'الجرائم المعلوماتية بكافة أنواعها.',                                                                         noorValue: '1201126,الدرجة الرابعة' },
    513: { noorText: 'التنمر بجميع أنواعه وأشكاله.',                                                                              noorValue: '1601185,الدرجة الرابعة' },
    702: { noorText: 'التلفظ بألفاظ غير لائقة تجاه المعلمين، أو الإداريين، أو من في حكمهم من منسوبي المدرسة.',                     noorValue: '1201129,الدرجة الرابعة' },
    706: { noorText: 'السخرية من المعلمين أو الإداريين أو من في حكمهم من منسوبي المدرسة، قولًا أو فعلًا.',                         noorValue: '1601214,الدرجة الرابعة' },
    707: { noorText: 'التوقيع عن أحد منسوبي المدرسة على المكاتبات المتبادلة بين المدرسة وأولياء الأمور.',                          noorValue: '1201131,الدرجة الرابعة' },
    708: { noorText: 'تصوير المعلمين أو الإداريين، أو من في حكمهم من منسوبي المدرسة، أو التسجيل الصوتي لهم ( مالم يؤخذ إذن خطي بالموافقة الصريحة على ذلك).', noorValue: '1201132,الدرجة الرابعة' },

    // ── الدرجة الخامسة (ابتدائي) ──
    703: { noorText: 'الاعتداء بالضرب على المعلمين أو الإداريين أو من في حكمهم من منسوبي المدرسة.',                                 noorValue: '1201135,الدرجة الخامسة' },
    704: { noorText: 'ابتزاز المعلمين، أو الإداريين ، أو من في حكمهم من منسوبي المدرسة.',                                          noorValue: '1201136,الدرجة الخامسة' },
    705: { noorText: 'الجرائم المعلوماتية تجاه المعلمين أو الإداريين أو من في حكمهم من منسوبي المدرسة.',                            noorValue: '1601183,الدرجة الخامسة' },
    709: { noorText: 'إلحاق الضرر بممتلكات المعلمين أو الإداريين، أو من في حكمهم من منسوبي المدرسة، أو سرقتها.',                   noorValue: '1201133,الدرجة الخامسة' },
    710: { noorText: 'الإشارة بحركات مخلة بالأدب تجاه المعلمين أو الإداريين، أو من في حكمهم من منسوبي المدرسة.',                   noorValue: '1201134,الدرجة الخامسة' }
    // ★ ملاحظة: المخالفات 105, 307, 503-505, 510, 512 (متوسط وثانوي فقط) غير موجودة في منسدلة ابتدائي
    // ★ ملاحظة: المخالفات الرقمية 601-620 غير موجودة في هذه المنسدلة
    // ★ ملاحظة: المخالفة 701 (تهديد المعلمين) مدمجة مع 702 في نور
  },

  // ═══════════════════════════════════════════════════════════════════════
  //  سلوك متميز — ddlDegreeDeductAmount عند سلوك=1, ايجابية=2
  //  المفتاح = نص السلوك المختصر (بدون لاحقة التصنيف)
  // ═══════════════════════════════════════════════════════════════════════
  excellent: {
    'انضباط الطالب وعدم غيابه بدون عذر خلال الفصل الدراسي':  { noorText: 'انضباط الطالب وعدم غيابه بدون عذر خلال الفصل الدراسي. (سلوك متميز)',                          noorValue: '1601248,' },
    'التعاون مع الزملاء والمعلمين وإدارة المدرسة':            { noorText: 'التعاون مع الزملاء والمعلمين وإدارة المدرسة. (سلوك متميز)',                                    noorValue: '1601240,' },
    'المشاركة في الإذاعة':                                    { noorText: 'المشاركة في الإذاعة. (سلوك متميز)',                                                            noorValue: '1601241,' },
    'المشاركة في الخدمة المجتمعية خارج المدرسة':              { noorText: 'المشاركة في الخدمة المجتمعية خارج المدرسة. (سلوك متميز)',                                       noorValue: '1601242,' },
    'المشاركة في أنشطة المهارات الرقمية':                     { noorText: 'المشاركة في أنشطة المهارات الرقمية( إعداد العروض، تصميم المحتوى الإلكتروني ). (سلوك متميز)',     noorValue: '1601243,' },
    'المشاركة في أنشطة مهارات الاتصال':                       { noorText: 'المشاركة في أنشطة مهارات الاتصال ( العمل الجماعي ، التعلم بالأقران،..). (سلوك متميز)',           noorValue: '1601244,' },
    'المشاركة في أنشطة مهارات القيادة والمسؤولية':            { noorText: 'المشاركة في أنشطة مهارات القيادة والمسؤولية (التخطيط، التحفيز،..). (سلوك متميز)',               noorValue: '1601245,' },
    'المشاركة في أنشطة مهارة إدارة الوقت':                    { noorText: 'المشاركة في أنشطة مهارة إدارة الوقت. (سلوك متميز)',                                             noorValue: '1601246,' },
    'المشاركة في حملة توعوية':                                { noorText: 'المشاركة في حملة توعوية. (سلوك متميز)',                                                         noorValue: '1601247,' },
    'تقديم فعالية حوارية':                                    { noorText: 'تقديم فعالية حوارية. (سلوك متميز)',                                                             noorValue: '1601249,' },
    'تقديم مقترح لصالح المجتمع المدرسي':                      { noorText: 'تقديم مقترح لصالح المجتمع المدرسي. (سلوك متميز)',                                               noorValue: '1601250,' },
    'عرض تجارب شخصية ناجحة':                                  { noorText: 'عرض تجارب شخصية ناجحة. (سلوك متميز)',                                                           noorValue: '1601251,' },
    'كتابة رسالة شكر':                                        { noorText: 'كتابة رسالة شكر(مثلا رسالة للوطن، للقيادة الرشيدة. للأسرة، للمعلم...إلخ). (سلوك متميز)',       noorValue: '1601252,' },
    'الالتحاق ببرنامج أو دورة':                               { noorText: 'الالتحاق ببرنامج أو دورة . (سلوك متميز)',                                                      noorValue: '1601239,' },
    // أخرى — 6 متغيرات بفروقات مسافات وترقيم
    'أخرى (بناءً على توصية لجنة التوجيه الطلابي) متميز':      { noorText: 'أخرى ( بناءً على توصية لجنة التوجيه الطلابي) (سلوك متميز)',                                     noorValue: '1601234,' }
  },

  // ═══════════════════════════════════════════════════════════════════════
  //  فرص تعويض — ddlDegreeDeductAmount عند سلوك=1, ايجابية=2
  //  المفتاح = نص السلوك المختصر (بدون لاحقة التصنيف)
  // ═══════════════════════════════════════════════════════════════════════
  compensation: {
    'انضباط الطالب وعدم غيابه بدون عذر خلال الفصل الدراسي':  { noorText: 'انضباط الطالب وعدم غيابه بدون عذر خلال الفصل الدراسي. (فرص تعويض)',                            noorValue: '189,' },
    'التعاون مع الزملاء والمعلمين وإدارة المدرسة':            { noorText: 'التعاون مع الزملاء والمعلمين وإدارة المدرسة. (فرص تعويض)',                                      noorValue: '1201017,' },
    'المشاركة في الإذاعة':                                    { noorText: 'المشاركة في الإذاعة. (فرص تعويض)',                                                              noorValue: '1601204,' },
    'المشاركة في الخدمة المجتمعية خارج المدرسة':              { noorText: 'المشاركة في الخدمة المجتمعية خارج المدرسة. (فرص تعويض)',                                         noorValue: '1601194,' },
    'المشاركة في أنشطة المهارات الرقمية':                     { noorText: 'المشاركة في أنشطة المهارات الرقمية( إعداد العروض، تصميم المحتوى الإلكتروني ). (فرص تعويض)',       noorValue: '1601201,' },
    'المشاركة في أنشطة مهارات الاتصال':                       { noorText: 'المشاركة في أنشطة مهارات الاتصال ( العمل الجماعي ، التعلم بالأقران،..). (فرص تعويض)',             noorValue: '1601199,' },
    'المشاركة في أنشطة مهارات القيادة والمسؤولية':            { noorText: 'المشاركة في أنشطة مهارات القيادة والمسؤولية (التخطيط، التحفيز،..). (فرص تعويض)',                 noorValue: '1601200,' },
    'المشاركة في أنشطة مهارة إدارة الوقت':                    { noorText: 'المشاركة في أنشطة مهارة إدارة الوقت. (فرص تعويض)',                                               noorValue: '1601202,' },
    'المشاركة في حملة توعوية':                                { noorText: 'المشاركة في حملة توعوية. (فرص تعويض)',                                                           noorValue: '1601196,' },
    'تقديم فعالية حوارية':                                    { noorText: 'تقديم فعالية حوارية. (فرص تعويض)',                                                               noorValue: '1601195,' },
    'تقديم مقترح لصالح المجتمع المدرسي':                      { noorText: 'تقديم مقترح لصالح المجتمع المدرسي. (فرص تعويض)',                                                 noorValue: '1601205,' },
    'عرض تجارب شخصية ناجحة':                                  { noorText: 'عرض تجارب شخصية ناجحة. (فرص تعويض)',                                                             noorValue: '1601197,' },
    'كتابة رسالة شكر':                                        { noorText: 'كتابة رسالة شكر(مثلا رسالة للوطن، للقيادة الرشيدة. للأسرة، للمعلم...إلخ). (فرص تعويض)',         noorValue: '1601203,' },
    'الالتحاق ببرنامج أو دورة':                               { noorText: 'الالتحاق ببرنامج أو دورة . (فرص تعويض)',                                                        noorValue: '1601198,' },
    // أخرى — الافتراضي
    'أخرى (بناءً على توصية لجنة التوجيه الطلابي) تعويض':      { noorText: 'اخرى ( بناءً على توصية لجنة التوجيه الطلابي) (فرص تعويض)',                                       noorValue: '1601207,' }
  }
};

// ═══════════════════════════════════════════════════════════════════════════
// دوال مساعدة لمطابقة النصوص العربية مع خريطة نور
// ═══════════════════════════════════════════════════════════════════════════

/**
 * تطبيع النص العربي — إزالة الترقيم وتوحيد الحروف
 */
function normalizeArabicForMatch_(text) {
  return String(text || '').trim()
    .replace(/[.،,؛:!؟\u200c\u200d]/g, '')
    .replace(/\s+/g, ' ')
    .replace(/[أإآ]/g, 'ا')
    .replace(/ة/g, 'ه')
    .replace(/ى/g, 'ي');
}

/**
 * بحث في خريطة نور عن أقرب تطابق نصي
 * @param {Object} map - كائن الخريطة { key: { noorText, noorValue } }
 * @param {string} text - النص المراد البحث عنه
 * @return {Object|null} { noorText, noorValue } أو null
 */
function findNoorMapping_(map, text) {
  if (!map || !text) return null;
  var norm = normalizeArabicForMatch_(text);

  // 1. بحث مباشر بالمفتاح
  if (map[text]) return map[text];

  // 2. بحث بالمفتاح بعد التطبيع
  for (var key in map) {
    if (map.hasOwnProperty(key)) {
      if (normalizeArabicForMatch_(key) === norm) return map[key];
    }
  }

  // 3. بحث جزئي (النص يحتوي على المفتاح أو العكس)
  for (var key2 in map) {
    if (map.hasOwnProperty(key2)) {
      var normKey = normalizeArabicForMatch_(key2);
      if (norm.indexOf(normKey) >= 0 || normKey.indexOf(norm) >= 0) return map[key2];
    }
  }

  return null;
}
// =================================================================
// ⚠️ ملف تطويري فقط — يجب حذفه من بيئة الإنتاج
// TestDebug.gs - تشخيص شامل (شغّل diagnoseAllSections من المحرر)
// =================================================================

function diagnoseAllSections() {
  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var tz = Session.getScriptTimeZone();
  var todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  
  Logger.log('═══════════════════════════════════════════');
  Logger.log('  تشخيص شامل - ' + todayStr);
  Logger.log('  TimeZone: ' + tz);
  Logger.log('═══════════════════════════════════════════');
  
  var sheets = [
    { name: 'سجل_الملاحظات_التربوية_متوسط', dateCol: 'وقت الإدخال', type: 'ملاحظات متوسط' },
    { name: 'سجل_الملاحظات_التربوية_ثانوي', dateCol: 'وقت الإدخال', type: 'ملاحظات ثانوي' },
    { name: 'سجل_التأخر_متوسط', dateCol: 'وقت_الإدخال', type: 'تأخر متوسط' },
    { name: 'سجل_التأخر_ثانوي', dateCol: 'وقت_الإدخال', type: 'تأخر ثانوي' },
    { name: 'سجل_الاستئذان_متوسط', dateCol: 'وقت_الإدخال', type: 'استئذان متوسط' },
    { name: 'سجل_الاستئذان_ثانوي', dateCol: 'وقت_الإدخال', type: 'استئذان ثانوي' },
    { name: 'سجل_الغياب_اليومي_متوسط', dateCol: 'وقت_الإدخال', type: 'غياب يومي متوسط' },
    { name: 'سجل_الغياب_اليومي_ثانوي', dateCol: 'وقت_الإدخال', type: 'غياب يومي ثانوي' }
  ];
  
  for (var s = 0; s < sheets.length; s++) {
    var info = sheets[s];
    var sheet = findSheet_(ss, info.name);
    
    if (!sheet) {
      Logger.log('\n❌ ' + info.type + ' → الشيت غير موجود: ' + info.name);
      continue;
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var totalRows = data.length - 1;
    
    var dateIdx = findHeaderIndex_(headers, info.dateCol);
    if (dateIdx < 0) dateIdx = findHeaderIndex_(headers, info.dateCol.replace(/ /g, '_'));
    if (dateIdx < 0) dateIdx = findHeaderIndex_(headers, info.dateCol.replace(/_/g, ' '));
    
    var todayCount = 0;
    var totalValid = 0;
    var samples = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[1]) continue;
      totalValid++;
      
      if (dateIdx >= 0) {
        var dateVal = row[dateIdx];
        var isToday = isTodayDate_(dateVal);
        if (isToday) todayCount++;
        if (samples.length < 3) {
          samples.push({ 
            row: i, 
            value: String(dateVal).substring(0, 35), 
            type: dateVal instanceof Date ? 'Date' : typeof dateVal,
            isToday: isToday
          });
        }
      }
    }
    
    var filtered = filterTodayRecords_(data, headers, info.dateCol);
    
    Logger.log('\n📋 ' + info.type + ' [' + sheet.getName() + ']');
    Logger.log('   عمود التاريخ: "' + info.dateCol + '" → index=' + dateIdx + ' (من ' + headers.length + ' عمود)');
    Logger.log('   صفوف: ' + totalRows + ' صالحة | اليوم: ' + todayCount + ' | filterTodayRecords_: ' + filtered.length);
    
    if (todayCount !== filtered.length) {
      Logger.log('   ⚠️ فرق بين العد اليدوي والفلتر!');
    }
    
    for (var j = 0; j < samples.length; j++) {
      var sm = samples[j];
      Logger.log('   صف' + sm.row + ': [' + sm.type + '] ' + sm.value + (sm.isToday ? ' ← اليوم ✅' : ' ✗'));
    }
    
    if (todayCount === 0) Logger.log('   ⚠️ لا يوجد أي سجل لهذا اليوم!');
  }
  
  Logger.log('\n═══════════════════════════════════════════');
  Logger.log('✅ اكتمل التشخيص - انسخ النتيجة كاملة');
  Logger.log('═══════════════════════════════════════════');
}

// =================================================================
// ★★★ اختبار المسار الكامل: غياب يومي → ظهور في نور
// يستخدم طلاب حقيقيين من بيانات النظام
// شغّل TEST_AbsenceToNoor() من المحرر
// =================================================================

// ★ متغير عام — يحفظ أرقام الطلاب المُدخلين للتنظيف لاحقاً
var _testInsertedIds = [];

/**
 * ★ الدالة الرئيسية — شغّلها من المحرر
 * تقرأ طالبين حقيقيين → تدخلهم كغياب يومي → تتحقق من ظهورهم في نور
 */
function TEST_AbsenceToNoor() {
  var stage = _getTestStage_();
  Logger.log('═══════════════════════════════════════════');
  Logger.log('★ اختبار المسار الكامل: غياب → نور');
  Logger.log('  المرحلة: ' + stage);
  Logger.log('═══════════════════════════════════════════');

  // ── 1. جلب طالبين حقيقيين من شيت الطلاب ──
  Logger.log('\n── الخطوة 1: جلب طالبين حقيقيين ──');
  var allStudents = getStudents_();
  var stageStudents = allStudents.filter(function(s) { return s['المرحلة'] === stage; });

  if (stageStudents.length < 2) {
    Logger.log('❌ لا يوجد طلاب كافيين في مرحلة ' + stage + ' (وُجد: ' + stageStudents.length + ')');
    Logger.log('  المراحل المتاحة:');
    var stageCount = {};
    allStudents.forEach(function(s) { var st = s['المرحلة'] || '?'; stageCount[st] = (stageCount[st] || 0) + 1; });
    for (var sk in stageCount) Logger.log('    ' + sk + ': ' + stageCount[sk] + ' طالب');
    return;
  }

  // اختيار طالبين من فصول مختلفة إن أمكن
  var student1 = stageStudents[0];
  var student2 = stageStudents.length > 5 ? stageStudents[5] : stageStudents[1];

  Logger.log('  الطالب 1: ' + student1['اسم الطالب'] + ' | ' + student1['رقم الطالب'] + ' | ' + student1['الصف'] + '/' + student1['الفصل']);
  Logger.log('  الطالب 2: ' + student2['اسم الطالب'] + ' | ' + student2['رقم الطالب'] + ' | ' + student2['الصف'] + '/' + student2['الفصل']);

  // حفظ الأرقام للتنظيف
  _testInsertedIds = [String(student1['رقم الطالب']), String(student2['رقم الطالب'])];

  // ── 2. إدخالهم كغياب يومي ──
  Logger.log('\n── الخطوة 2: إدخال الغياب اليومي ──');
  var now = new Date();
  var hijriDate = getHijriDate_(now);
  var dayName = getDayNameAr_(now);
  Logger.log('  التاريخ الهجري: ' + hijriDate);
  Logger.log('  اليوم: ' + dayName);

  var sheet = ensureDailyAbsenceSheet_(stage);
  var rows = [
    // الطالب 1: بدون عذر
    [
      student1['رقم الطالب'], student1['اسم الطالب'],
      student1['الصف'], student1['الفصل'],
      student1['رقم الجوال'] || '',
      'يوم كامل', '', hijriDate, dayName,
      'اختبار_نور', now, 'معلق', 'بدون عذر', 'لا', 'غائب', '', 'اختبار_مسار_نور'
    ],
    // الطالب 2: بعذر
    [
      student2['رقم الطالب'], student2['اسم الطالب'],
      student2['الصف'], student2['الفصل'],
      student2['رقم الجوال'] || '',
      'يوم كامل', '', hijriDate, dayName,
      'اختبار_نور', now, 'معلق', 'بعذر', 'لا', 'غائب', '', 'اختبار_مسار_نور'
    ]
  ];

  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 8, rows.length, 1).setNumberFormat('@');
  sheet.getRange(startRow, 1, rows.length, 17).setValues(rows);

  Logger.log('  ✅ تم إدخال سجلين في: سجل_الغياب_اليومي_' + stage);
  Logger.log('    صف ' + startRow + ': ' + student1['اسم الطالب'] + ' → بدون عذر (كود نور: 48)');
  Logger.log('    صف ' + (startRow + 1) + ': ' + student2['اسم الطالب'] + ' → بعذر (كود نور: 141)');

  // ── 3. التحقق من ظهورهم في getNoorPendingRecords ──
  Logger.log('\n── الخطوة 3: فحص ظهورهم في تبويب نور ──');
  var noorResult = getNoorPendingRecords(stage, 'absence');
  Logger.log('  getNoorPendingRecords: success=' + noorResult.success + ' | total=' + noorResult.total);

  if (!noorResult.success) {
    Logger.log('  ❌ خطأ: ' + (noorResult.error || 'غير معروف'));
    return;
  }

  // البحث عن الطالبين في النتائج
  var found1 = false, found2 = false;
  for (var i = 0; i < (noorResult.records || []).length; i++) {
    var rec = noorResult.records[i];
    var recId = String(rec['رقم_الطالب'] || rec['رقم الطالب'] || '');
    var recName = rec['اسم_الطالب'] || rec['اسم الطالب'] || '';
    var recExcuse = rec['نوع_العذر'] || rec['نوع العذر'] || '';
    var recHijri = rec['التاريخ_هجري'] || rec['التاريخ هجري'] || '';

    if (recId === _testInsertedIds[0]) {
      found1 = true;
      Logger.log('  ✅ ظهر الطالب 1: ' + recName + ' | عذر: ' + recExcuse + ' | هجري: ' + recHijri + ' | كود نور: 48,');
    }
    if (recId === _testInsertedIds[1]) {
      found2 = true;
      Logger.log('  ✅ ظهر الطالب 2: ' + recName + ' | عذر: ' + recExcuse + ' | هجري: ' + recHijri + ' | كود نور: 141,');
    }
  }

  if (!found1 || !found2) {
    Logger.log('');
    if (!found1) Logger.log('  ❌ الطالب 1 لم يظهر! (' + student1['اسم الطالب'] + ')');
    if (!found2) Logger.log('  ❌ الطالب 2 لم يظهر! (' + student2['اسم الطالب'] + ')');

    // تشخيص
    var todayHijri = getTodayHijriDate_();
    Logger.log('\n  ── تشخيص ──');
    Logger.log('  تاريخ اليوم الهجري: [' + todayHijri + '] مطبّع: [' + normalizeHijriDate_(todayHijri) + ']');
    Logger.log('  التاريخ المُدخل:    [' + hijriDate + '] مطبّع: [' + normalizeHijriDate_(hijriDate) + ']');
    Logger.log('  تطابق: ' + (normalizeHijriDate_(hijriDate) === normalizeHijriDate_(todayHijri)));
  }

  // ── 4. ملخص ──
  Logger.log('\n═══════════════════════════════════════════');
  if (found1 && found2) {
    Logger.log('🎉 نجح الاختبار! الطالبان يظهران في تبويب نور');
    Logger.log('');
    Logger.log('→ الخطوة التالية:');
    Logger.log('  1. افتح التطبيق');
    Logger.log('  2. اختر المرحلة: ' + stage);
    Logger.log('  3. اذهب لـ "التوثيق في نور"');
    Logger.log('  4. اختر تبويب "غياب يومي"');
    Logger.log('  5. يجب أن ترى الطالبين مع علامة ✓ مطابق');
    Logger.log('  6. اضغط "بدء التوثيق في نور" لإكمال العملية');
    Logger.log('');
    Logger.log('→ بعد الانتهاء شغّل TEST_CleanupAbsence() للتنظيف');
  } else {
    Logger.log('⚠️ الاختبار لم يكتمل — راجع التشخيص أعلاه');
  }
  Logger.log('═══════════════════════════════════════════');
}

/**
 * ★ تنظيف سجلات الاختبار
 * يحذف السجلات التي أدخلتها TEST_AbsenceToNoor (بالبحث عن "اختبار_مسار_نور" في الملاحظات)
 */
function TEST_CleanupAbsence() {
  var stage = _getTestStage_();
  Logger.log('═══════════════════════════════════════════');
  Logger.log('★ تنظيف سجلات اختبار الغياب');
  Logger.log('  المرحلة: ' + stage);
  Logger.log('═══════════════════════════════════════════');

  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var sheetName = 'سجل_الغياب_اليومي_' + stage;
  var sheet = findSheet_(ss, sheetName);

  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('❌ الشيت فارغ أو غير موجود');
    return;
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var notesCol = headers.indexOf('ملاحظات');
  var recorderCol = headers.indexOf('المسجل');
  if (notesCol < 0) notesCol = headers.indexOf('ملاحظات');
  var deleted = 0;
  var deletedNames = [];

  // حذف من الأسفل للأعلى
  for (var i = data.length - 1; i >= 1; i--) {
    var notes = String(data[i][notesCol] || '');
    var recorder = recorderCol >= 0 ? String(data[i][recorderCol] || '') : '';
    if (notes === 'اختبار_مسار_نور' || recorder === 'اختبار_نور') {
      deletedNames.push(String(data[i][1] || ''));
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }

  if (deleted > 0) {
    Logger.log('✅ تم حذف ' + deleted + ' سجل:');
    for (var d = 0; d < deletedNames.length; d++) {
      Logger.log('   - ' + deletedNames[d]);
    }
  } else {
    Logger.log('ℹ️ لا توجد سجلات اختبار لحذفها');
  }
  Logger.log('═══════════════════════════════════════════');
}

// =================================================================
// ★★★ تشخيص مشكلة التوثيق في نور
// السجلات لا تظهر رغم وجود الأرقام في الإحصائيات
// شغّل TEST_DiagNoorRecords() من المحرر
// =================================================================
function TEST_DiagNoorRecords() {
  var stage = _getTestStage_();
  Logger.log('═══════════════════════════════════════════');
  Logger.log('★ تشخيص التوثيق في نور');
  Logger.log('  المرحلة: ' + stage);
  Logger.log('═══════════════════════════════════════════');

  // ── 1. اختبار getNoorStats (ما يظهر في الأرقام فوق) ──
  Logger.log('\n── 1. getNoorStats (الإحصائيات) ──');
  try {
    var stats = getNoorStats(stage);
    Logger.log('  success: ' + stats.success);
    if (stats.success) {
      Logger.log('  المخالفات: ' + stats.pending.violations);
      Logger.log('  التأخر: ' + stats.pending.tardiness);
      Logger.log('  التعويضية: ' + stats.pending.compensation);
      Logger.log('  المتمايز: ' + stats.pending.excellent);
      Logger.log('  الغياب: ' + stats.pending.absence);
      Logger.log('  المجموع: ' + stats.pending.total);
    } else {
      Logger.log('  ❌ خطأ: ' + stats.error);
    }
  } catch (e) {
    Logger.log('  ❌ استثناء: ' + e.message);
  }

  // ── 2. اختبار كل نوع على حدة ──
  var types = ['violations', 'tardiness', 'compensation', 'excellent', 'absence'];
  for (var t = 0; t < types.length; t++) {
    var type = types[t];
    Logger.log('\n── 2.' + (t + 1) + ' getNoorPendingRecords("' + type + '") ──');
    try {
      var result = getNoorPendingRecords(stage, type);
      Logger.log('  success: ' + result.success);
      Logger.log('  total: ' + result.total);
      Logger.log('  records.length: ' + (result.records ? result.records.length : 'NULL'));
      if (!result.success) {
        Logger.log('  ❌ خطأ: ' + (result.error || 'غير معروف'));
      }
      if (result.records && result.records.length > 0) {
        var rec = result.records[0];
        Logger.log('  أول سجل:');
        Logger.log('    اسم: ' + (rec['اسم_الطالب'] || rec['اسم الطالب'] || '?'));
        Logger.log('    _rowIndex: ' + rec._rowIndex);
        Logger.log('    _type: ' + rec._type);
        // فحص حجم البيانات
        var jsonSize = JSON.stringify(result).length;
        Logger.log('  حجم JSON: ' + jsonSize + ' حرف (' + Math.round(jsonSize / 1024) + ' KB)');
      }
    } catch (e) {
      Logger.log('  ❌ استثناء: ' + e.message);
      Logger.log('  Stack: ' + e.stack);
    }
  }

  // ── 3. اختبار 'all' ──
  Logger.log('\n── 3. getNoorPendingRecords("all") ──');
  try {
    var allResult = getNoorPendingRecords(stage, 'all');
    Logger.log('  success: ' + allResult.success);
    Logger.log('  total: ' + allResult.total);
    Logger.log('  records.length: ' + (allResult.records ? allResult.records.length : 'NULL'));
    if (!allResult.success) {
      Logger.log('  ❌ خطأ: ' + (allResult.error || 'غير معروف'));
    }
    if (allResult.records && allResult.records.length > 0) {
      var jsonSize = JSON.stringify(allResult).length;
      Logger.log('  حجم JSON الكلي: ' + jsonSize + ' حرف (' + Math.round(jsonSize / 1024) + ' KB)');
    }
    Logger.log('  stats: ' + JSON.stringify(allResult.stats));
  } catch (e) {
    Logger.log('  ❌ استثناء: ' + e.message);
  }

  // ── 4. فحص مباشر للشيتات ──
  Logger.log('\n── 4. فحص الشيتات مباشرة ──');
  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var sheetTypes = {
    'المخالفات': 'violations',
    'التأخر': 'tardiness',
    'السلوك_الإيجابي': 'positive',
    'الغياب_اليومي': 'absence'
  };
  for (var sType in sheetTypes) {
    var sheetName = getSheetName_(sType, stage);
    var sheet = findSheet_(ss, sheetName);
    if (!sheet) {
      Logger.log('  ❌ ' + sType + ': الشيت غير موجود (' + sheetName + ')');
    } else {
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      Logger.log('  ✅ ' + sType + ': ' + sheet.getName() + ' | صفوف: ' + lastRow + ' | أعمدة: ' + lastCol);
      if (lastRow >= 1) {
        var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        var headerNames = headers.map(function(h) { return String(h).trim(); }).filter(function(h) { return h; });
        Logger.log('     الترويسات: ' + headerNames.join(' | '));
        var hasNoorCol = headerNames.indexOf('حالة_نور') >= 0 || headerNames.indexOf('حالة نور') >= 0;
        Logger.log('     عمود حالة_نور: ' + (hasNoorCol ? 'موجود ✅' : 'غير موجود ⚠️'));
      }
    }
  }

  // ── 5. فحص التاريخ الهجري ──
  Logger.log('\n── 5. فحص التاريخ الهجري ──');
  var todayHijri = getTodayHijriDate_();
  Logger.log('  getTodayHijriDate_(): [' + todayHijri + ']');
  Logger.log('  normalizeHijriDate_(): [' + normalizeHijriDate_(todayHijri) + ']');
  Logger.log('  getHijriDate_(new Date()): [' + getHijriDate_(new Date()) + ']');

  Logger.log('\n═══════════════════════════════════════════');
  Logger.log('✅ اكتمل التشخيص');
  Logger.log('═══════════════════════════════════════════');
}

// ── دوال مساعدة للاختبار ──

function _getTestStage_() {
  ensureStudentsSheetsLoaded_();
  var stages = STUDENTS_SHEETS ? Object.keys(STUDENTS_SHEETS) : [];
  if (stages.indexOf('متوسط') >= 0) return 'متوسط';
  if (stages.indexOf('ثانوي') >= 0) return 'ثانوي';
  if (stages.length > 0) return stages[0];
  return 'متوسط';
}
// =================================================================
// Server_Absence_Daily.gs - دوال الغياب اليومي
// =================================================================

/**
 * التحقق من صحة rowIndex — يمنع الوصول لصف الترويسة أو صفوف خارج النطاق
 * rowIndex هنا 0-indexed (الصف الأول من البيانات = 0)، يُضاف إليه 1 عند استخدام getRange
 */
function validateRowIndex_(rowIndex, sheet) {
  if (typeof rowIndex !== 'number' || !isFinite(rowIndex) || rowIndex < 1 || rowIndex !== Math.floor(rowIndex)) {
    throw new Error('rowIndex غير صالح: ' + rowIndex);
  }
  if (rowIndex + 1 > sheet.getLastRow()) {
    throw new Error('rowIndex خارج نطاق البيانات: ' + rowIndex);
  }
}

/**
 * التحقق من مصفوفة rowIndices
 */
function validateRowIndices_(rowIndices, sheet) {
  if (!Array.isArray(rowIndices) || rowIndices.length === 0) {
    throw new Error('rowIndices مطلوبة');
  }
  var lastRow = sheet.getLastRow();
  for (var i = 0; i < rowIndices.length; i++) {
    var ri = rowIndices[i];
    if (typeof ri !== 'number' || !isFinite(ri) || ri < 1 || ri !== Math.floor(ri) || ri + 1 > lastRow) {
      throw new Error('rowIndex غير صالح في المصفوفة: ' + ri);
    }
  }
}

// =================================================================
// ★ دوال مركزية موحدة لكتابة بيانات الغياب اليومي
// =================================================================

/**
 * الهيدرز الـ 17 الثابتة لشيت الغياب اليومي
 */
function getDailyAbsenceHeaders_() {
  return [
    'رقم_الطالب', 'اسم_الطالب', 'الصف', 'الفصل', 'رقم_الجوال',
    'نوع_الغياب', 'الحصة', 'التاريخ_هجري', 'اليوم', 'المسجل',
    'وقت_الإدخال', 'حالة_الاعتماد', 'نوع_العذر', 'تم_الإرسال',
    'حالة_التأخر', 'وقت_الحضور', 'ملاحظات'
  ];
}

/**
 * إنشاء/جلب شيت الغياب اليومي بتنسيق موحد
 */
function ensureDailyAbsenceSheet_(stage) {
  var ss = getSpreadsheet_();
  var sheetName = 'سجل_الغياب_اليومي_' + stage;
  var sheet = findSheet_(ss, sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.setRightToLeft(true);
    var headers = getDailyAbsenceHeaders_();
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#f3f4f6').setFontWeight('bold');
    sheet.setFrozenRows(1);
    // ★ تنسيق عمود التاريخ الهجري (8) كنص عادي لمنع التحويل التلقائي
    sheet.getRange(1, 8, sheet.getMaxRows(), 1).setNumberFormat('@');
    var _regColor = (SHEET_REGISTRY['الغياب_اليومي'] || {}).color;
    if (_regColor) sheet.setTabColor(_regColor);
  }

  return sheet;
}

/**
 * تطبيع نوع الغياب — دائماً يُرجع 'يوم كامل' أو 'حصة' فقط
 * يُرجع كائن { type: '...', period: '...' } لدعم نقل اسم الحصة
 */
function normalizeDailyAbsenceType_(rawType, existingPeriod) {
  if (!rawType) return { type: 'يوم كامل', period: existingPeriod || '' };

  var t = String(rawType).trim();

  // القيم المعيارية
  if (t === 'يوم كامل') return { type: 'يوم كامل', period: '' };
  if (t === 'حصة')      return { type: 'حصة', period: existingPeriod || '' };

  // قيم تُحوّل إلى 'يوم كامل'
  if (t === 'غياب بدون عذر' || t === 'غياب بعذر' || t === 'غائب' || t === 'absent') {
    return { type: 'يوم كامل', period: '' };
  }

  // قيم تُحوّل إلى 'حصة'
  if (t === 'غائب عن حصة') return { type: 'حصة', period: existingPeriod || '' };

  // أسماء الحصص → 'حصة' + نقل الاسم لعمود الحصة
  var periodNames = ['الأولى','الثانية','الثالثة','الرابعة','الخامسة','السادسة','السابعة'];
  if (periodNames.indexOf(t) !== -1) {
    return { type: 'حصة', period: t };
  }

  // افتراضي
  return { type: 'يوم كامل', period: '' };
}

/**
 * بناء صف غياب يومي موحد (17 عمود)
 * params: {
 *   studentId, studentName, grade, section, phone,
 *   absenceType, period, recorder, notes,
 *   dateOverride (اختياري — Date object)
 * }
 */
function buildDailyAbsenceRow_(params) {
  var now = params.dateOverride || new Date();
  var hijriDate = getHijriDate_(now);
  var dayName = getDayNameAr_(now);

  // تطبيع نوع الغياب
  var normalized = normalizeDailyAbsenceType_(params.absenceType, params.period);

  return [
    sanitizeInput_(String(params.studentId   || '')),  // 1:  رقم_الطالب
    sanitizeInput_(String(params.studentName || '')),  // 2:  اسم_الطالب
    sanitizeInput_(String(params.grade       || '')),  // 3:  الصف
    sanitizeInput_(String(params.section     || '')),  // 4:  الفصل
    sanitizeInput_(String(params.phone       || '')),  // 5:  رقم_الجوال
    normalized.type,                                    // 6:  نوع_الغياب ('يوم كامل'|'حصة')
    sanitizeInput_(String(normalized.period  || '')),  // 7:  الحصة
    hijriDate,                                          // 8:  التاريخ_هجري
    dayName,                                            // 9:  اليوم
    sanitizeInput_(String(params.recorder    || '')),  // 10: المسجل
    now,                                                // 11: وقت_الإدخال (Date object)
    'معلق',                                             // 12: حالة_الاعتماد
    '',                                                 // 13: نوع_العذر
    'لا',                                               // 14: تم_الإرسال
    'غائب',                                             // 15: حالة_التأخر
    '',                                                 // 16: وقت_الحضور
    sanitizeInput_(String(params.notes       || ''))   // 17: ملاحظات
  ];
}

/**
 * كتابة مصفوفة صفوف دفعة واحدة بـ setValues (أداء أفضل من appendRow)
 */
function writeDailyAbsenceRows_(sheet, rows) {
  if (!rows || rows.length === 0) return 0;
  var startRow = sheet.getLastRow() + 1;
  // ★ تنسيق عمود التاريخ الهجري كنص قبل الكتابة (يمنع Sheets من تحويله لـ Date)
  sheet.getRange(startRow, 8, rows.length, 1).setNumberFormat('@');
  sheet.getRange(startRow, 1, rows.length, 17).setValues(rows);
  return rows.length;
}

// =================================================================
// 1. حفظ سجل غياب يومي جديد
// =================================================================
function saveDailyAbsenceRecord(data) {
  try {
    var sheet = ensureDailyAbsenceSheet_(data.stage);

    var students = data.students;
    var rows = [];

    for (var i = 0; i < students.length; i++) {
      var student = students[i];
      rows.push(buildDailyAbsenceRow_({
        studentId:   student.id,
        studentName: student.name,
        grade:       student.grade,
        section:     student.class,
        phone:       student.phone,
        absenceType: data.absenceType,
        period:      data.period,
        recorder:    data.recorder || 'يدوي',
        notes:       data.notes
      }));
    }

    var savedCount = writeDailyAbsenceRows_(sheet, rows);
    return { success: true, message: 'تم حفظ ' + savedCount + ' سجل بنجاح' };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 2. جلب سجلات غياب اليوم
// =================================================================
function getTodayAbsenceRecords(stage) {
  try {
    var ss = getSpreadsheet_();
    var sheetName = "سجل_الغياب_اليومي_" + stage;
    var sheet = findSheet_(ss, sheetName);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, records: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var records = filterTodayRecords_(data, headers, 'وقت_الإدخال');
    
    return { success: true, records: records };
    
  } catch (e) {
    return { success: false, error: e.toString(), records: [] };
  }
}

// =================================================================
// 3. جلب أرشيف الغياب اليومي
// =================================================================
function getAbsenceArchive(stage, dateFrom, dateTo) {
  try {
    var ss = getSpreadsheet_();
    var sheetName = "سجل_الغياب_اليومي_" + stage;
    var sheet = findSheet_(ss, sheetName);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, records: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    // بناء خريطة الأعمدة الموحدة
    var headerMap = {};
    var seen = {};
    for (var h = 0; h < headers.length; h++) {
      var standardized = String(headers[h] || '').trim().replace(/\s+/g, '_');
      if (!standardized || seen[standardized]) continue;
      seen[standardized] = true;
      headerMap[h] = standardized;
    }
    
    var records = [];
    var fromDate = dateFrom ? new Date(dateFrom) : null;
    var toDate = dateTo ? new Date(dateTo) : null;
    if (toDate) toDate.setHours(23, 59, 59);
    
    var tz = Session.getScriptTimeZone();
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[1]) continue;
      
      // استخراج التاريخ بشكل مرن
      var dateValue = getRowDateValue_(row, headers, 'وقت_الإدخال');
      
      // تحويل لكائن Date للمقارنة
      var recordDate = null;
      if (dateValue instanceof Date) {
        recordDate = dateValue;
      } else if (typeof dateValue === 'string' && dateValue) {
        try { recordDate = new Date(dateValue); } catch(e) {}
      }
      
      // فلتر التاريخ (السجلات بدون تاريخ تُستبعد إذا فيه فلتر)
      if (fromDate || toDate) {
        if (!recordDate || isNaN(recordDate.getTime())) continue;
        if (fromDate && recordDate < fromDate) continue;
        if (toDate && recordDate > toDate) continue;
      }
      
      var record = { rowIndex: i };
      for (var j in headerMap) {
        var value = row[j];
        if (value instanceof Date) {
          value = Utilities.formatDate(value, tz, 'yyyy/MM/dd HH:mm');
        }
        record[headerMap[j]] = String(value || '');
      }
      records.push(record);
    }
    
    return { success: true, records: records };
    
  } catch (e) {
    return { success: false, error: e.toString(), records: [] };
  }
}

// =================================================================
// 4. تحديث حالة الاعتماد
// =================================================================
function updateAbsenceApprovalStatus(stage, rowIndex, status, excuseType) {
  try {
    var ss = getSpreadsheet_();
    var sheetName = "سجل_الغياب_اليومي_" + stage;
    var sheet = findSheet_(ss, sheetName);

    if (!sheet) {
      return { success: false, error: 'الشيت غير موجود' };
    }
    validateRowIndex_(rowIndex, sheet);

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var statusCol = headers.indexOf('حالة_الاعتماد') + 1;
    var excuseCol = headers.indexOf('نوع_العذر') + 1;
    
    if (statusCol > 0) {
      sheet.getRange(rowIndex + 1, statusCol).setValue(status);
    }
    if (excuseCol > 0 && excuseType) {
      sheet.getRange(rowIndex + 1, excuseCol).setValue(excuseType);
    }
    
    // تحديث الشيت التراكمي إذا تم الاعتماد
    if (status === 'معتمد' && excuseType) {
      updateCumulativeAbsence_(stage, rowIndex, excuseType);
    }
    
    return { success: true };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 5. تحديث الغياب التراكمي
// =================================================================
function updateCumulativeAbsence_(stage, rowIndex, excuseType) {
  try {
    var ss = getSpreadsheet_();
    
    // قراءة بيانات الطالب من السجل اليومي
    var dailySheetName = "سجل_الغياب_اليومي_" + stage;
    var dailySheet = findSheet_(ss, dailySheetName);
    var headers = dailySheet.getRange(1, 1, 1, dailySheet.getLastColumn()).getValues()[0];
    var rowData = dailySheet.getRange(rowIndex + 1, 1, 1, dailySheet.getLastColumn()).getValues()[0];
    
    var studentId = rowData[headers.indexOf('رقم_الطالب')];
    
    // الشيت التراكمي
    var cumulativeSheetName = "سجل_الغياب_" + stage;
    var cumulativeSheet = findSheet_(ss, cumulativeSheetName);
    
    if (!cumulativeSheet) return;
    
    var cumulativeData = cumulativeSheet.getDataRange().getValues();
    var cumulativeHeaders = cumulativeData[0];
    var idCol = cumulativeHeaders.indexOf('رقم الطالب');
    var excusedCol = cumulativeHeaders.indexOf('غياب بعذر');
    var unexcusedCol = cumulativeHeaders.indexOf('غياب بدون عذر');
    var updateCol = cumulativeHeaders.indexOf('آخر تحديث');
    
    // البحث عن الطالب
    for (var i = 1; i < cumulativeData.length; i++) {
      if (String(cumulativeData[i][idCol]) === String(studentId)) {
        var targetRow = i + 1;
        if (excuseType === 'بعذر') {
          var currentExcused = Number(cumulativeData[i][excusedCol]) || 0;
          cumulativeSheet.getRange(targetRow, excusedCol + 1).setValue(currentExcused + 1);
        } else {
          var currentUnexcused = Number(cumulativeData[i][unexcusedCol]) || 0;
          cumulativeSheet.getRange(targetRow, unexcusedCol + 1).setValue(currentUnexcused + 1);
        }
        cumulativeSheet.getRange(targetRow, updateCol + 1).setValue(new Date());
        break;
      }
    }
    
  } catch (e) {
    Logger.log('خطأ في تحديث الغياب التراكمي: ' + e.toString());
  }
}

// =================================================================
// 6. حذف سجل غياب
// =================================================================
function deleteAbsenceRecord(stage, rowIndex) {
  try {
    var ss = getSpreadsheet_();
    var sheetName = "سجل_الغياب_اليومي_" + stage;
    var sheet = findSheet_(ss, sheetName);

    if (!sheet) {
      return { success: false, error: 'الشيت غير موجود' };
    }
    validateRowIndex_(rowIndex, sheet);

    sheet.deleteRow(rowIndex + 1);
    return { success: true };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 7. تحديث حالة الإرسال
// =================================================================
function updateAbsenceSentStatus(stage, rowIndices) {
  try {
    var ss = getSpreadsheet_();
    var sheetName = "سجل_الغياب_اليومي_" + stage;
    var sheet = findSheet_(ss, sheetName);

    if (!sheet) {
      return { success: false, error: 'الشيت غير موجود' };
    }
    validateRowIndices_(rowIndices, sheet);

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var sentCol = headers.indexOf('تم_الإرسال') + 1;
    
    if (sentCol > 0) {
      for (var i = 0; i < rowIndices.length; i++) {
        sheet.getRange(rowIndices[i] + 1, sentCol).setValue('نعم');
      }
    }
    
    return { success: true };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 8. جلب إحصائيات الغياب للتقارير
// =================================================================
function getAbsenceStatistics(stage, dateFrom, dateTo, gradeFilter, classFilter) {
  try {
    var ss = getSpreadsheet_();
    var sheetName = "سجل_الغياب_اليومي_" + stage;
    var sheet = findSheet_(ss, sheetName);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, stats: getEmptyStats_() };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var fromDate = dateFrom ? new Date(dateFrom) : null;
    var toDate = dateTo ? new Date(dateTo) : null;
    if (toDate) toDate.setHours(23, 59, 59);
    
    var stats = {
      total: 0,
      approved: 0,
      pending: 0,
      rejected: 0,
      withExcuse: 0,
      withoutExcuse: 0,
      fullDay: 0,
      period: 0,
      byGrade: {},
      byClass: {},
      byDay: {},
      topStudents: []
    };
    
    var studentCounts = {};
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      
      var recordDate = row[headers.indexOf('وقت_الإدخال')];
      var grade = row[headers.indexOf('الصف')];
      var cls = row[headers.indexOf('الفصل')];
      var status = row[headers.indexOf('حالة_الاعتماد')];
      var excuseType = row[headers.indexOf('نوع_العذر')];
      var absenceType = row[headers.indexOf('نوع_الغياب')];
      var dayName = row[headers.indexOf('اليوم')];
      var studentId = row[headers.indexOf('رقم_الطالب')];
      var studentName = row[headers.indexOf('اسم_الطالب')];
      
      // فلترة التاريخ
      if (recordDate instanceof Date) {
        if (fromDate && recordDate < fromDate) continue;
        if (toDate && recordDate > toDate) continue;
      }
      
      // فلترة الصف والفصل
      if (gradeFilter && grade !== gradeFilter) continue;
      if (classFilter && String(cls) !== String(classFilter)) continue;
      
      stats.total++;
      
      // حسب الحالة
      if (status === 'معتمد') stats.approved++;
      else if (status === 'مرفوض') stats.rejected++;
      else stats.pending++;
      
      // حسب العذر
      if (excuseType === 'بعذر') stats.withExcuse++;
      else if (excuseType === 'بدون عذر') stats.withoutExcuse++;
      
      // حسب نوع الغياب
      if (absenceType === 'يوم كامل') stats.fullDay++;
      else stats.period++;
      
      // حسب الصف
      stats.byGrade[grade] = (stats.byGrade[grade] || 0) + 1;
      
      // حسب الفصل
      var classKey = grade + ' / ' + cls;
      stats.byClass[classKey] = (stats.byClass[classKey] || 0) + 1;
      
      // حسب اليوم
      stats.byDay[dayName] = (stats.byDay[dayName] || 0) + 1;
      
      // عد الطلاب
      if (!studentCounts[studentId]) {
        studentCounts[studentId] = { id: studentId, name: studentName, grade: grade, count: 0 };
      }
      studentCounts[studentId].count++;
    }
    
    // أكثر الطلاب غياباً
    var studentArray = Object.values(studentCounts);
    studentArray.sort(function(a, b) { return b.count - a.count; });
    stats.topStudents = studentArray.slice(0, 10);
    
    return { success: true, stats: stats };
    
  } catch (e) {
    return { success: false, error: e.toString(), stats: getEmptyStats_() };
  }
}

function getEmptyStats_() {
  return {
    total: 0, approved: 0, pending: 0, rejected: 0,
    withExcuse: 0, withoutExcuse: 0, fullDay: 0, period: 0,
    byGrade: {}, byClass: {}, byDay: {}, topStudents: []
  };
}

// =================================================================
// دوال مساعدة (محذوفة - تستخدم المركزية من Config.gs)
// getHijriDate_() → Config.gs
// getDayNameAr_() → Config.gs
// =================================================================

// =================================================================
// 9. استيراد بيانات من Excel (نظام نور)
// =================================================================
function importExcelAbsence(stage, students, notes) {
  try {
    var sheet = ensureDailyAbsenceSheet_(stage);

    // ★ تحديد المسجل والملاحظة حسب نوع الاستيراد
    var isPlatform = (notes && notes === 'منصة');
    var recorderLabel = isPlatform ? 'استيراد منصة' : 'مستورد من رابط الغياب';
    var finalNotes   = isPlatform ? 'منصة' : 'مستورد من نظام نور';

    var rows = [];

    for (var i = 0; i < students.length; i++) {
      var student = students[i];
      rows.push(buildDailyAbsenceRow_({
        studentId:   student.id,
        studentName: student.name,
        grade:       student.grade,
        section:     student.class,
        phone:       student.phone,
        absenceType: student.absenceType,  // يمر عبر normalizeDailyAbsenceType_ تلقائياً
        period:      '',
        recorder:    recorderLabel,
        notes:       finalNotes
      }));
    }

    var savedCount = writeDailyAbsenceRows_(sheet, rows);

    var msg = isPlatform
      ? 'تم استيراد ' + savedCount + ' طالب كغياب منصة'
      : 'تم استيراد ' + savedCount + ' طالب بنجاح';
    return { success: true, message: msg };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 10. تحديث حالة التأخر
// =================================================================
function updateAbsenceLateStatus(stage, rowIndex, status, lateTime) {
  try {
    var ss = getSpreadsheet_();
    var sheetName = "سجل_الغياب_اليومي_" + stage;
    var sheet = findSheet_(ss, sheetName);

    if (!sheet) {
      return { success: false, error: 'الشيت غير موجود' };
    }
    validateRowIndex_(rowIndex, sheet);

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var statusCol = headers.indexOf('حالة_التأخر') + 1;
    var timeCol = headers.indexOf('وقت_الحضور') + 1;
    
    // إضافة الأعمدة إذا لم تكن موجودة
    if (statusCol === 0) {
      statusCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, statusCol).setValue('حالة_التأخر');
    }
    if (timeCol === 0) {
      timeCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, timeCol).setValue('وقت_الحضور');
    }
    
    sheet.getRange(rowIndex + 1, statusCol).setValue(status);
    sheet.getRange(rowIndex + 1, timeCol).setValue(sanitizeInput_(lateTime || ''));
    
    return { success: true };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 11. إرسال إشعارات الغياب
// =================================================================
function sendAbsenceNotifications(stage, rowIndices) {
  try {
    var ss = getSpreadsheet_();
    var sheetName = "سجل_الغياب_اليومي_" + stage;
    var sheet = findSheet_(ss, sheetName);

    if (!sheet) {
      return { success: false, error: 'الشيت غير موجود' };
    }
    validateRowIndices_(rowIndices, sheet);

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var sentCol = headers.indexOf('تم_الإرسال') + 1;
    var phoneCol = headers.indexOf('رقم_الجوال') + 1;
    var nameCol = headers.indexOf('اسم_الطالب') + 1;
    var gradeCol = headers.indexOf('الصف') + 1;
    var classCol = headers.indexOf('الفصل') + 1;
    var idCol = headers.indexOf('رقم_الطالب') + 1;
    if (idCol < 1) idCol = headers.indexOf('رقم_الهوية') + 1;
    
    var today = new Date();
    var dayName = ['الأحد','الإثنين','الثلاثاء','الأربعاء','الخميس','الجمعة','السبت'][today.getDay()];
    var hijriDate = '';
    try { hijriDate = today.toLocaleDateString('ar-SA-u-ca-islamic', { day: 'numeric', month: 'long', year: 'numeric' }); } catch(e) { hijriDate = Utilities.formatDate(today, 'Asia/Riyadh', 'yyyy/MM/dd'); }
    
    var sentCount = 0;
    var failCount = 0;
    var noPhoneCount = 0;
    
    for (var i = 0; i < rowIndices.length; i++) {
      var rowIndex = rowIndices[i];
      var rowData = sheet.getRange(rowIndex + 1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      var phone = phoneCol > 0 ? String(rowData[phoneCol - 1] || '').trim() : '';
      var studentName = nameCol > 0 ? rowData[nameCol - 1] : '';
      var grade = gradeCol > 0 ? rowData[gradeCol - 1] : '';
      var cls = classCol > 0 ? rowData[classCol - 1] : '';
      var studentId = idCol > 0 ? String(rowData[idCol - 1] || '').trim() : '';
      
      // ★ fallback: إذا الشيت اليومي ما فيه جوال، ابحث في شيت الطلاب
      if (!phone && studentId) {
        try { phone = getStudentPhone_(studentId, stage); } catch(pe) { phone = ''; }
      }
      
      if (!phone) {
        noPhoneCount++;
        Logger.log('sendAbsenceNotifications: لا يوجد جوال للطالب ' + studentName + ' (ID: ' + studentId + ')');
        continue;
      }
      // ★ توليد رابط تقديم العذر لولي الأمر
      var parentLink = '';
      try {
        if (studentId) {
          parentLink = getParentExcuseLink_(String(studentId), stage);
          Logger.log('sendAbsenceNotifications: link for ' + studentName + ' = ' + parentLink);
        }
      } catch(linkErr) {
        Logger.log('sendAbsenceNotifications: link error for ' + studentName + ': ' + linkErr.toString());
        parentLink = '';
      }

      var message = '📋 *إشعار غياب*\n\nالسلام عليكم ورحمة الله وبركاته\nولي أمر الطالب: *' + studentName + '*\n\nنفيدكم بأن ابنكم *' + studentName + '* غائب اليوم\n📅 ' + dayName + ' - ' + hijriDate + '\nالصف: ' + grade + ' - الفصل: ' + cls;

      if (parentLink) {
        message += '\n\n📝 لكتابة العذر:\n' + parentLink + '\n(صالح لمدة ٢٤ ساعة)';
      }

      var _schoolName = '';
      try { _schoolName = getSchoolNameForLinks_(); } catch(e) { _schoolName = 'المدرسة'; }
      message += '\n\nمع تحيات إدارة مدرسة ' + _schoolName;

      // إرسال واتساب مع التوثيق
      var result = sendWhatsAppWithLog({
        phone: String(phone),
        message: message,
        studentName: studentName,
        studentId: String(studentId),
        grade: grade,
        class: cls,
        messageType: 'غياب',
        messageTitle: 'إشعار غياب'
      }, stage);
      
      if (result && result.success) {
        sentCount++;
        if (sentCol > 0) {
          sheet.getRange(rowIndex + 1, sentCol).setValue('نعم');
        }
      } else {
        failCount++;
      }
      
      // ★ تأخير 10 ثوان بين كل رسالة لحماية الرقم من الحظر
      if (i < rowIndices.length - 1) {
        Utilities.sleep(10000);
      }
    }
    
    var msg = 'تم إرسال ' + sentCount + ' إشعار';
    if (failCount > 0) msg += ' | فشل: ' + failCount;
    if (noPhoneCount > 0) msg += ' | بدون جوال: ' + noPhoneCount;
    return { success: true, message: msg, sent: sentCount, failed: failCount, noPhone: noPhoneCount };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// ★ جلب رابط عذر ولي الأمر (تُستدعى من الإرسال الفردي)
// =================================================================
function getParentExcuseLinkForStudent(studentId, stage) {
  try {
    if (!studentId) return { success: false, link: '', error: 'رقم الطالب فارغ' };
    var link = getParentExcuseLink_(String(studentId), stage);
    Logger.log('getParentExcuseLinkForStudent: studentId=' + studentId + ', stage=' + stage + ', link=' + link);
    return { success: true, link: link || '' };
  } catch(e) {
    Logger.log('getParentExcuseLinkForStudent ERROR: ' + e.toString());
    return { success: false, link: '', error: e.toString() };
  }
}

// =================================================================
// ★ إرسال فردي كامل من السيرفر (يبني الرسالة + الرابط + يرسل)
// =================================================================
function sendSingleAbsenceWithLink(stage, rowIndex) {
  try {
    var ss = getSpreadsheet_();
    var sheetName = "سجل_الغياب_اليومي_" + stage;
    var sheet = findSheet_(ss, sheetName);
    if (!sheet) return { success: false, error: 'الشيت غير موجود' };
    validateRowIndex_(rowIndex, sheet);

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var rowData = sheet.getRange(rowIndex + 1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var phoneCol = headers.indexOf('رقم_الجوال') + 1;
    var nameCol = headers.indexOf('اسم_الطالب') + 1;
    var gradeCol = headers.indexOf('الصف') + 1;
    var classCol = headers.indexOf('الفصل') + 1;
    var idCol = headers.indexOf('رقم_الطالب') + 1;
    if (idCol < 1) idCol = headers.indexOf('رقم_الهوية') + 1;
    var sentCol = headers.indexOf('تم_الإرسال') + 1;

    var phone = phoneCol > 0 ? String(rowData[phoneCol - 1] || '') : '';
    var studentName = nameCol > 0 ? String(rowData[nameCol - 1] || '') : '';
    var grade = gradeCol > 0 ? String(rowData[gradeCol - 1] || '') : '';
    var cls = classCol > 0 ? String(rowData[classCol - 1] || '') : '';
    var studentId = idCol > 0 ? String(rowData[idCol - 1] || '') : '';

    // إذا لم يوجد رقم جوال في الشيت اليومي، ابحث في شيت الطلاب
    if (!phone || !phone.trim()) {
      phone = getStudentPhone_(studentId, stage);
    }

    if (!phone || !phone.trim()) {
      return { success: false, error: 'لا يوجد رقم جوال لهذا الطالب' };
    }

    var today = new Date();
    var dayName = ['الأحد','الإثنين','الثلاثاء','الأربعاء','الخميس','الجمعة','السبت'][today.getDay()];
    var hijriDate = '';
    try { hijriDate = today.toLocaleDateString('ar-SA-u-ca-islamic', { day: 'numeric', month: 'long', year: 'numeric' }); } catch(e) { hijriDate = Utilities.formatDate(today, 'Asia/Riyadh', 'yyyy/MM/dd'); }

    // ★ توليد رابط ولي الأمر
    var parentLink = '';
    try {
      if (studentId) {
        parentLink = getParentExcuseLink_(studentId, stage);
        Logger.log('sendSingleAbsenceWithLink: link = ' + parentLink);
      }
    } catch(linkErr) {
      Logger.log('sendSingleAbsenceWithLink: link error: ' + linkErr.toString());
    }

    // ★ بناء الرسالة
    var message = '📋 *إشعار غياب*\n\nالسلام عليكم ورحمة الله وبركاته\nولي أمر الطالب: *' + studentName + '*\n\nنفيدكم بأن ابنكم *' + studentName + '* غائب اليوم\n📅 ' + dayName + ' - ' + hijriDate + '\nالصف: ' + grade + ' - الفصل: ' + cls;

    if (parentLink) {
      message += '\n\n📝 لكتابة العذر:\n' + parentLink + '\n(صالح لمدة ٢٤ ساعة)';
    }

    var _schoolName2 = '';
    try { _schoolName2 = getSchoolNameForLinks_(); } catch(e) { _schoolName2 = 'المدرسة'; }
    message += '\n\nمع تحيات إدارة مدرسة ' + _schoolName2;

    // ★ إرسال
    var result = sendWhatsAppWithLog({
      phone: phone,
      message: message,
      studentName: studentName,
      studentId: studentId,
      grade: grade,
      'class': cls,
      messageType: 'غياب',
      messageTitle: 'إشعار غياب'
    }, stage);

    if (result && result.success) {
      if (sentCol > 0) {
        sheet.getRange(rowIndex + 1, sentCol).setValue('نعم');
      }
      return { success: true, message: 'تم إرسال الإشعار لولي أمر ' + studentName, hasLink: !!parentLink };
    } else {
      return { success: false, error: (result && result.error) || 'فشل الإرسال' };
    }

  } catch(e) {
    Logger.log('sendSingleAbsenceWithLink ERROR: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 12. ترحيل الغياب اليومي (تشغيل تلقائي الساعة 12 صباحاً)
// =================================================================
function archiveDailyAbsence() {
  try {
    ensureStudentsSheetsLoaded_();
    var stages = Object.keys(STUDENTS_SHEETS); // ★ ديناميكي: يشمل أي مرحلة مُعرّفة
    var archivedCount = 0;
    
    stages.forEach(function(stage) {
      var ss = getSpreadsheet_();
      var dailySheetName = "سجل_الغياب_اليومي_" + stage;
      var dailySheet = findSheet_(ss, dailySheetName);
      
      if (!dailySheet || dailySheet.getLastRow() < 2) return;
      
      var data = dailySheet.getDataRange().getValues();
      var headers = data[0];
      var today = new Date();
      var todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      
      // ترحيل سجلات اليوم إلى السجل التراكمي
      for (var i = data.length - 1; i >= 1; i--) {
        var row = data[i];
        var recordDate = row[headers.indexOf('وقت_الإدخال')];
        
        if (recordDate instanceof Date) {
          var recordDateStr = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          
          // ترحيل سجلات الأمس وما قبله
          if (recordDateStr < todayStr) {
            // تحديث السجل التراكمي — فقط إذا لم يُعتمد مسبقاً (لتجنب العد المزدوج)
            var approvalStatus = String(row[headers.indexOf('حالة_الاعتماد')] || '').trim();
            var lateStatus = row[headers.indexOf('حالة_التأخر')] || 'غائب';

            if (lateStatus === 'غائب' && approvalStatus !== 'معتمد' && approvalStatus !== 'مقبول') {
              updateCumulativeAbsence_(stage, i, row[headers.indexOf('نوع_العذر')] || 'بدون عذر');
            }

            // حذف السجل من اليومي
            dailySheet.deleteRow(i + 1);
            archivedCount++;
          }
        }
      }
    });
    
    Logger.log('تم ترحيل ' + archivedCount + ' سجل');
    return { success: true, archived: archivedCount };
    
  } catch (e) {
    Logger.log('خطأ في الترحيل: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 13. إنشاء مشغل تلقائي للترحيل
// =================================================================
function createArchiveTrigger() {
  // حذف المشغلات القديمة
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'archiveDailyAbsence') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // إنشاء مشغل جديد يعمل الساعة 12 صباحاً
  ScriptApp.newTrigger('archiveDailyAbsence')
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .create();
    
  return { success: true, message: 'تم إنشاء المشغل التلقائي' };
}

// =================================================================
// ★ جلب قائمة الأعذار (سجلات الغياب المعلقة)
// =================================================================
function getExcusesList(stage) {
  try {
    var ss = getSpreadsheet_();
    var sheetName = "سجل_الغياب_اليومي_" + stage;
    var sheet = findSheet_(ss, sheetName);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, excuses: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim().replace(/\s+/g, '_'); });
    
    var excuses = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[1]) continue;
      
      var idIdx = headers.indexOf('رقم_الطالب');
      var nameIdx = headers.indexOf('اسم_الطالب');
      var gradeIdx = headers.indexOf('الصف');
      var classIdx = headers.indexOf('الفصل');
      var phoneIdx = headers.indexOf('رقم_الجوال');
      var typeIdx = headers.indexOf('نوع_الغياب');
      var dateIdx = headers.indexOf('التاريخ_هجري');
      var statusIdx = headers.indexOf('حالة_الاعتماد');
      var excuseIdx = headers.indexOf('نوع_العذر');
      var notesIdx = headers.indexOf('ملاحظات');
      var dayIdx = headers.indexOf('اليوم');
      
      var status = String(row[statusIdx] || '').trim();
      var mappedStatus = 'pending';
      if (status === 'معتمد' || status === 'مقبول') mappedStatus = 'approved';
      else if (status === 'مرفوض') mappedStatus = 'rejected';
      
      // ★ جلب التاريخ الميلادي من عمود وقت_الإدخال
      var inputTimeIdx = headers.indexOf('وقت_الإدخال');
      var miladiDate = '';
      try {
        var rawTime = row[inputTimeIdx >= 0 ? inputTimeIdx : 0];
        if (rawTime instanceof Date) {
          miladiDate = Utilities.formatDate(rawTime, 'Asia/Riyadh', 'yyyy-MM-dd');
        }
      } catch(de) {}

      excuses.push({
        id: String(i + 1),
        rowIndex: i + 1,
        studentId: String(row[idIdx] || ''),
        studentName: String(row[nameIdx] || ''),
        className: String(row[gradeIdx] || '') + ' / ' + String(row[classIdx] || ''),
        grade: String(row[gradeIdx] || ''),
        section: String(row[classIdx] || ''),
        phone: String(row[phoneIdx] || ''),
        absenceDate: String(row[dateIdx] || ''),
        day: String(row[dayIdx] || ''),
        reason: String(row[excuseIdx] || row[typeIdx] || 'بدون عذر'),
        details: String(row[notesIdx] || ''),
        status: mappedStatus,
        parentName: 'ولي الأمر',
        submittedAt: String(row[dateIdx] || ''),
        miladiDate: miladiDate
      });
    }
    
    // ★ جلب أعذار أولياء الأمور من شيت اعذار_اولياء_الامور
    try {
      var parentExcuseSheet = findSheet_(ss, 'اعذار_اولياء_الامور');
      if (parentExcuseSheet && parentExcuseSheet.getLastRow() >= 2) {
        var peData = parentExcuseSheet.getDataRange().getValues();
        var peHeaders = peData[0].map(function(h) { return String(h).trim().replace(/\s+/g, '_'); });

        var pe_idIdx = peHeaders.indexOf('رقم_الطالب');
        var pe_nameIdx = peHeaders.indexOf('اسم_الطالب');
        var pe_gradeIdx = peHeaders.indexOf('الصف');
        var pe_classIdx = peHeaders.indexOf('الفصل');
        var pe_stageIdx = peHeaders.indexOf('المرحلة');
        var pe_reasonIdx = peHeaders.indexOf('نص_العذر');
        var pe_attachIdx = peHeaders.indexOf('مرفقات');
        var pe_absDateIdx = peHeaders.indexOf('تاريخ_الغياب');
        var pe_dateIdx = peHeaders.indexOf('تاريخ_التقديم');
        var pe_timeIdx = peHeaders.indexOf('وقت_التقديم');
        var pe_statusIdx = peHeaders.indexOf('الحالة');
        var pe_notesIdx = peHeaders.indexOf('ملاحظات_المدرسة');

        for (var p = 1; p < peData.length; p++) {
          var peRow = peData[p];
          if (!peRow[pe_idIdx]) continue;

          // فلتر بالمرحلة
          var peStage = String(peRow[pe_stageIdx] || '').trim();
          if (peStage && peStage !== stage) continue;

          var peStatus = String(peRow[pe_statusIdx] || '').trim();
          var peMapped = 'pending';
          if (peStatus === 'معتمد' || peStatus === 'مقبول') peMapped = 'approved';
          else if (peStatus === 'مرفوض') peMapped = 'rejected';

          // ★ تاريخ ميلادي من وقت التقديم
          var peMiladiDate = '';
          try {
            var peTimeRaw = peRow[pe_timeIdx];
            if (peTimeRaw instanceof Date) {
              peMiladiDate = Utilities.formatDate(peTimeRaw, 'Asia/Riyadh', 'yyyy-MM-dd');
            } else {
              // إذا لم يكن Date، نستخدم تاريخ اليوم كاحتياط
              peMiladiDate = Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd');
            }
          } catch(pde) {}

          // ★ تاريخ الغياب (العمود الجديد) أو تاريخ التقديم كاحتياط
          var peAbsDate = (pe_absDateIdx >= 0 ? String(peRow[pe_absDateIdx] || '') : '') || String(peRow[pe_dateIdx] || '');

          excuses.push({
            id: 'pe_' + (p + 1),
            rowIndex: p + 1,
            source: 'parent',
            studentId: String(peRow[pe_idIdx] || ''),
            studentName: String(peRow[pe_nameIdx] || ''),
            className: String(peRow[pe_gradeIdx] || '') + ' / ' + String(peRow[pe_classIdx] || ''),
            grade: String(peRow[pe_gradeIdx] || ''),
            section: String(peRow[pe_classIdx] || ''),
            phone: '',
            absenceDate: peAbsDate,
            day: '',
            reason: String(peRow[pe_reasonIdx] || ''),
            details: String(peRow[pe_notesIdx] || ''),
            hasAttachment: String(peRow[pe_attachIdx] || '') === 'نعم - تُسلم مع الطالب',
            status: peMapped,
            parentName: 'ولي الأمر (عبر الرابط)',
            submittedAt: String(peRow[pe_dateIdx] || '') + ' ' + String(peRow[pe_timeIdx] || ''),
            miladiDate: peMiladiDate
          });
        }
      }
    } catch(peErr) { /* تجاهل أخطاء شيت الأعذار */ }

    return { success: true, excuses: excuses };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// ★ تحديث حالة العذر (قبول/رفض)
// =================================================================
function updateExcuseStatus(excuseId, status, reason, sendMessage) {
  try {
    ensureStudentsSheetsLoaded_();
    var ss = getSpreadsheet_();
    var arabicStatus = status === 'approved' ? 'مقبول' : (status === 'rejected' ? 'مرفوض' : 'معلق');

    // ★ التحقق إذا كان عذر ولي أمر (يبدأ بـ pe_)
    if (String(excuseId).indexOf('pe_') === 0) {
      var peRowNum = parseInt(String(excuseId).replace('pe_', ''));
      var peSheet = findSheet_(ss, 'اعذار_اولياء_الامور');
      if (!peSheet || peSheet.getLastRow() < peRowNum) {
        return { success: false, error: 'لم يتم العثور على سجل العذر' };
      }

      var peHeaders = peSheet.getRange(1, 1, 1, peSheet.getLastColumn()).getValues()[0]
        .map(function(h) { return String(h).trim().replace(/\s+/g, '_'); });

      var peStatusIdx = peHeaders.indexOf('الحالة');
      var peNotesIdx = peHeaders.indexOf('ملاحظات_المدرسة');

      if (peStatusIdx >= 0) {
        peSheet.getRange(peRowNum, peStatusIdx + 1).setValue(arabicStatus);
      }

      if (reason && peNotesIdx >= 0) {
        var existingNotes = String(peSheet.getRange(peRowNum, peNotesIdx + 1).getValue() || '');
        peSheet.getRange(peRowNum, peNotesIdx + 1).setValue(existingNotes + (existingNotes ? ' | ' : '') + sanitizeInput_(reason));
      }

      return { success: true, message: 'تم تحديث الحالة إلى: ' + arabicStatus };
    }

    // ★ عذر من سجل الغياب اليومي
    var rowNum = parseInt(excuseId);
    var stages = Object.keys(STUDENTS_SHEETS); // ★ ديناميكي: يشمل أي مرحلة مُعرّفة
    for (var s = 0; s < stages.length; s++) {
      var sheetName = "سجل_الغياب_اليومي_" + stages[s];
      var sheet = findSheet_(ss, sheetName);
      if (!sheet || sheet.getLastRow() < rowNum) continue;

      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
        .map(function(h) { return String(h).trim().replace(/\s+/g, '_'); });

      var statusIdx = headers.indexOf('حالة_الاعتماد');
      var notesIdx = headers.indexOf('ملاحظات');

      if (statusIdx < 0) continue;

      var currentVal = sheet.getRange(rowNum, 1).getValue();
      if (!currentVal) continue;

      var dailyArabicStatus = status === 'approved' ? 'معتمد' : (status === 'rejected' ? 'مرفوض' : 'معلق');
      sheet.getRange(rowNum, statusIdx + 1).setValue(dailyArabicStatus);

      var excuseTypeIdx = headers.indexOf('نوع_العذر');
      if (excuseTypeIdx >= 0) {
        if (status === 'approved') {
          sheet.getRange(rowNum, excuseTypeIdx + 1).setValue('بعذر');
        } else if (status === 'rejected') {
          sheet.getRange(rowNum, excuseTypeIdx + 1).setValue('بدون عذر');
        }
      }

      if (reason && notesIdx >= 0) {
        var existingNotes = String(sheet.getRange(rowNum, notesIdx + 1).getValue() || '');
        sheet.getRange(rowNum, notesIdx + 1).setValue(existingNotes + (existingNotes ? ' | ' : '') + 'سبب الرفض: ' + sanitizeInput_(reason));
      }

      return { success: true, message: 'تم تحديث الحالة إلى: ' + dailyArabicStatus };
    }

    return { success: false, error: 'لم يتم العثور على السجل' };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// ★ إرسال رسالة مخصصة لولي الأمر عن العذر
// =================================================================
function sendExcuseCustomMessage(excuseId, message) {
  try {
    message = sanitizeInput_(message);
    ensureStudentsSheetsLoaded_();
    var ss = getSpreadsheet_();

    // ★ عذر ولي أمر - جلب الرقم من شيت الطلاب
    if (String(excuseId).indexOf('pe_') === 0) {
      var peRowNum = parseInt(String(excuseId).replace('pe_', ''));
      var peSheet = findSheet_(ss, 'اعذار_اولياء_الامور');
      if (!peSheet || peSheet.getLastRow() < peRowNum) {
        return { success: false, error: 'لم يتم العثور على سجل العذر' };
      }
      var peHeaders = peSheet.getRange(1, 1, 1, peSheet.getLastColumn()).getValues()[0]
        .map(function(h) { return String(h).trim().replace(/\s+/g, '_'); });
      var peRow = peSheet.getRange(peRowNum, 1, 1, peSheet.getLastColumn()).getValues()[0];
      var studentId = String(peRow[peHeaders.indexOf('رقم_الطالب')] || '');
      var studentName = String(peRow[peHeaders.indexOf('اسم_الطالب')] || '');
      var stage = String(peRow[peHeaders.indexOf('المرحلة')] || '');

      // جلب رقم الجوال من شيت الطلاب
      var phone = getStudentPhone_(studentId, stage);
      if (!phone) return { success: false, error: 'لا يوجد رقم جوال لهذا الطالب' };

      if (typeof sendWhatsAppWithLog === 'function') {
        var result = sendWhatsAppWithLog({
          phone: phone, message: message, studentName: studentName,
          studentId: studentId, messageType: 'عذر', messageTitle: 'رد على عذر غياب'
        }, stage);
        return result;
      }
      return { success: true, message: 'تم إرسال الرسالة' };
    }

    // ★ عذر من سجل الغياب اليومي
    var rowNum = parseInt(excuseId);
    var stages = Object.keys(STUDENTS_SHEETS); // ★ ديناميكي: يشمل أي مرحلة مُعرّفة
    for (var s = 0; s < stages.length; s++) {
      var sheetName = "سجل_الغياب_اليومي_" + stages[s];
      var sheet = findSheet_(ss, sheetName);
      if (!sheet || sheet.getLastRow() < rowNum) continue;

      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
        .map(function(h) { return String(h).trim().replace(/\s+/g, '_'); });

      var phoneIdx = headers.indexOf('رقم_الجوال');
      var nameIdx = headers.indexOf('اسم_الطالب');
      var idIdx = headers.indexOf('رقم_الطالب');

      if (phoneIdx < 0) continue;

      var row = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (!row[0]) continue;

      var phone = String(row[phoneIdx] || '');
      var studentName = String(row[nameIdx] || '');

      if (!phone) {
        // محاولة جلب الرقم من شيت الطلاب
        var studentId = String(row[idIdx] || '');
        phone = getStudentPhone_(studentId, stages[s]);
        if (!phone) return { success: false, error: 'لا يوجد رقم جوال' };
      }

      if (typeof sendWhatsAppWithLog === 'function') {
        var result = sendWhatsAppWithLog({
          phone: phone, message: message, studentName: studentName,
          studentId: String(row[idIdx] || ''), messageType: 'عذر', messageTitle: 'رد على عذر غياب'
        }, stages[s]);
        return result;
      }
      return { success: true, message: 'تم إرسال الرسالة' };
    }

    return { success: false, error: 'لم يتم العثور على السجل' };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ★ جلب رقم جوال الطالب من شيت الطلاب
function getStudentPhone_(studentId, stage) {
  try {
    ensureStudentsSheetsLoaded_();
    var ss = getSpreadsheet_();
    var stages = stage ? [stage] : Object.keys(STUDENTS_SHEETS); // ★ ديناميكي
    for (var s = 0; s < stages.length; s++) {
      var sheetName = STUDENTS_SHEETS[stages[s]];
      if (!sheetName) continue;
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) continue;
      var data = sheet.getDataRange().getValues();
      var headers = data[0];
      var idCol = -1, phoneCol = -1;
      for (var h = 0; h < headers.length; h++) {
        var hdr = String(headers[h]).trim().replace(/\s+/g, '_');
        if (hdr === 'رقم_الطالب' || hdr === 'رقم_الهوية') idCol = h;
        if (hdr === 'رقم_الجوال' || hdr === 'جوال_ولي_الأمر' || hdr === 'الجوال') phoneCol = h;
      }
      if (idCol < 0 || phoneCol < 0) continue;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][idCol]).trim() === String(studentId).trim()) {
          var phone = String(data[i][phoneCol] || '').trim();
          if (phone) return phone;
        }
      }
    }
    return '';
  } catch(e) { return ''; }
}

// =================================================================
// ★ تحديث نوع العذر (بعذر / بدون عذر)
// =================================================================
function updateAbsenceExcuseType(stage, rowIndex, excuseType) {
  try {
    var ss = getSpreadsheet_();
    var sheetName = "سجل_الغياب_اليومي_" + stage;
    var sheet = findSheet_(ss, sheetName);

    if (!sheet) {
      return { success: false, error: 'الشيت غير موجود' };
    }
    validateRowIndex_(rowIndex, sheet);

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var excuseCol = headers.indexOf('نوع_العذر') + 1;
    
    if (excuseCol > 0) {
      sheet.getRange(rowIndex + 1, excuseCol).setValue(excuseType);
    }
    
    return { success: true };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
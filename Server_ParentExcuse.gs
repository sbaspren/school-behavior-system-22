// =================================================================
// Server_ParentExcuse.gs - دوال واجهة ولي الأمر لتقديم الأعذار
// =================================================================

// =================================================================
// 1. توليد رمز ولي الأمر (عند إرسال إشعار الغياب)
// =================================================================
function generateParentToken_(studentId, stage) {
  var token = Utilities.getUuid().replace(/-/g, '').substring(0, 16);
  var ss = getSpreadsheet_();
  var sheetName = 'رموز_اولياء_الامور';
  var sheet = findSheet_(ss, sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.setRightToLeft(true);
    var headers = ['الرمز', 'رقم_الطالب', 'المرحلة', 'تاريخ_الإنشاء', 'تاريخ_الانتهاء', 'مستخدم'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#7c3aed').setFontColor('#fff').setFontWeight('bold');
  }

  var now = new Date();
  var expiry = new Date(now.getTime() + 24 * 60 * 60 * 1000); // 24 ساعة

  sheet.appendRow([token, String(studentId), stage, now, expiry, 'لا']);

  // ★ تنظيف دوري: حذف التوكنات المنتهية (أقدم من 48 ساعة) لمنع تراكم الشيت
  try {
    var allData = sheet.getDataRange().getValues();
    var cutoff = new Date(now.getTime() - 48 * 60 * 60 * 1000);
    var rowsToDelete = [];
    for (var i = allData.length - 1; i >= 1; i--) {
      var created = new Date(allData[i][3]);
      if (created < cutoff) {
        rowsToDelete.push(i + 1);
      }
    }
    // حذف من الأسفل للأعلى حتى لا يتغير ترقيم الصفوف
    for (var d = 0; d < rowsToDelete.length && d < 50; d++) {
      sheet.deleteRow(rowsToDelete[d]);
    }
  } catch(cleanErr) {
    Logger.log('Token cleanup warning: ' + cleanErr.message);
  }

  return token;
}

// =================================================================
// 2. التحقق من رمز ولي الأمر وجلب بيانات الطالب
// =================================================================
function getParentExcusePageData_(token) {
  try {
    if (!token) return { success: false, error: 'الرمز مطلوب' };

    var ss = getSpreadsheet_();
    var tokenSheet = findSheet_(ss, 'رموز_اولياء_الامور');

    if (!tokenSheet || tokenSheet.getLastRow() < 2) {
      return { success: false, error: 'رابط غير صالح أو منتهي الصلاحية' };
    }

    var tokenData = tokenSheet.getDataRange().getValues();
    var headers = tokenData[0];
    var tokenRow = null;
    var tokenRowIndex = -1;

    for (var i = 1; i < tokenData.length; i++) {
      if (String(tokenData[i][0]).trim() === token) {
        tokenRow = tokenData[i];
        tokenRowIndex = i + 1;
        break;
      }
    }

    if (!tokenRow) return { success: false, error: 'رابط غير صالح' };

    // التحقق من الصلاحية (24 ساعة)
    var expiry = new Date(tokenRow[4]);
    if (new Date() > expiry) {
      return { success: false, error: 'انتهت صلاحية هذا الرابط (24 ساعة). سيتم إنشاء رابط جديد مع إشعار الغياب القادم.' };
    }

    // ★ إصلاح: إذا التوكن مستخدم — أخبر ولي الأمر
    var isUsed = String(tokenRow[5] || '').trim();
    if (isUsed === 'نعم') {
      return { success: false, error: 'تم تقديم العذر مسبقاً عبر هذا الرابط. شكراً لتعاونكم.' };
    }

    var studentId = String(tokenRow[1]).trim();
    var stage = String(tokenRow[2]).trim();

    // جلب بيانات الطالب من شيت الطلاب
    var studentInfo = getStudentInfoForParent_(studentId, stage);
    if (!studentInfo) {
      return { success: false, error: 'لم يتم العثور على بيانات الطالب' };
    }

    // جلب إحصائيات الغياب من شيت الغياب التراكمي
    var absenceStats = getStudentAbsenceStats_(studentId, stage);

    // جلب تاريخ اليوم هجري
    var now = new Date();
    var hijriDate = getHijriDateFull_(now).hijriStr || Utilities.formatDate(now, 'Asia/Riyadh', 'yyyy/MM/dd');
    var dayName = getDayNameAr_(now);

    // جلب اسم المدرسة من الإعدادات
    var schoolName = '';
    try {
      var settingsSheet = ss.getSheetByName('إعدادات_المدرسة');
      if (settingsSheet && settingsSheet.getLastRow() >= 2) {
        var sData = settingsSheet.getDataRange().getValues();
        for (var s = 1; s < sData.length; s++) {
          if (String(sData[s][0]).trim() === 'school_name') { schoolName = String(sData[s][1] || '').trim(); break; }
        }
      }
    } catch(e) {}

    return {
      success: true,
      schoolName: schoolName,
      student: {
        id: studentInfo.id,
        name: studentInfo.name,
        grade: studentInfo.grade,
        section: studentInfo.section,
        stage: stage
      },
      absence: {
        excused: absenceStats.excused || 0,
        unexcused: absenceStats.unexcused || 0,
        late: absenceStats.late || 0
      },
      today: {
        date: hijriDate,
        day: dayName
      }
    };

  } catch (e) {
    return { success: false, error: 'خطأ في تحميل البيانات: ' + e.toString() };
  }
}

// =================================================================
// 3. جلب بيانات الطالب من شيت الطلاب
// =================================================================
function getStudentInfoForParent_(studentId, stage) {
  try {
    ensureStudentsSheetsLoaded_();
    var ss = getSpreadsheet_();
    // البحث في كل شيتات الطلاب
    var stages = stage ? [stage] : Object.keys(STUDENTS_SHEETS);

    for (var s = 0; s < stages.length; s++) {
      var sheetName = STUDENTS_SHEETS[stages[s]];
      if (!sheetName) continue;
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) continue;

      var data = sheet.getDataRange().getValues();
      var headers = data[0];

      var idCol = -1, nameCol = -1, gradeCol = -1, classCol = -1;
      for (var h = 0; h < headers.length; h++) {
        var hdr = String(headers[h]).trim().replace(/\s+/g, '_');
        if (hdr === 'رقم_الطالب' || hdr === 'رقم_الهوية') idCol = h;
        if (hdr === 'اسم_الطالب') nameCol = h;
        if (hdr === 'الصف') gradeCol = h;
        if (hdr === 'الفصل') classCol = h;
      }

      if (idCol < 0) continue;

      for (var i = 1; i < data.length; i++) {
        if (String(data[i][idCol]).trim() === studentId) {
          return {
            id: studentId,
            name: String(data[i][nameCol] || ''),
            grade: String(data[i][gradeCol] || ''),
            section: String(data[i][classCol] || ''),
            stage: stages[s]
          };
        }
      }
    }

    return null;
  } catch (e) {
    return null;
  }
}

// =================================================================
// 4. جلب إحصائيات غياب الطالب من الشيت التراكمي
// =================================================================
function getStudentAbsenceStats_(studentId, stage) {
  try {
    var ss = getSpreadsheet_();
    var sheetName = 'سجل_الغياب_' + stage;
    var sheet = findSheet_(ss, sheetName);

    if (!sheet || sheet.getLastRow() < 2) {
      return { excused: 0, unexcused: 0, late: 0 };
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    // ★ إصلاح: البحث بأسماء الأعمدة بدل الأرقام الثابتة
    var idCol = -1, excusedCol = -1, unexcusedCol = -1, lateCol = -1;
    for (var h = 0; h < headers.length; h++) {
      var hdr = String(headers[h]).trim().replace(/\s+/g, '_');
      if (hdr === 'رقم_الطالب' || hdr === 'رقم_الهوية') idCol = h;
      if (hdr === 'غياب_بعذر' || hdr.indexOf('بعذر') > -1) excusedCol = h;
      if (hdr === 'غياب_بدون_عذر' || hdr.indexOf('بدون_عذر') > -1) unexcusedCol = h;
      if (hdr === 'تأخير' || hdr === 'تأخر') lateCol = h;
    }

    if (idCol < 0) return { excused: 0, unexcused: 0, late: 0 };

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]).trim() === studentId) {
        return {
          excused: excusedCol >= 0 ? (Number(data[i][excusedCol]) || 0) : 0,
          unexcused: unexcusedCol >= 0 ? (Number(data[i][unexcusedCol]) || 0) : 0,
          late: lateCol >= 0 ? (Number(data[i][lateCol]) || 0) : 0
        };
      }
    }

    return { excused: 0, unexcused: 0, late: 0 };
  } catch (e) {
    return { excused: 0, unexcused: 0, late: 0 };
  }
}

// =================================================================
// 5. حفظ عذر ولي الأمر
// =================================================================
function submitParentExcuse(data) {
  try {
    if (!data || !data.token || !data.reason) {
      return { success: false, error: 'البيانات غير مكتملة' };
    }

    // التحقق من الرمز
    var ss = getSpreadsheet_();
    var tokenSheet = findSheet_(ss, 'رموز_اولياء_الامور');
    if (!tokenSheet) return { success: false, error: 'خطأ في النظام' };

    var tokenData = tokenSheet.getDataRange().getValues();
    var tokenRow = null;
    var tokenRowIdx = -1;

    for (var i = 1; i < tokenData.length; i++) {
      if (String(tokenData[i][0]).trim() === data.token) {
        tokenRow = tokenData[i];
        tokenRowIdx = i + 1;
        break;
      }
    }

    if (!tokenRow) return { success: false, error: 'رابط غير صالح' };

    // التحقق من الصلاحية (24 ساعة)
    var expiry = new Date(tokenRow[4]);
    if (new Date() > expiry) {
      return { success: false, error: 'انتهت صلاحية الرابط (24 ساعة). سيتم إنشاء رابط جديد مع إشعار الغياب القادم.' };
    }

    // ★ إصلاح: التحقق إذا التوكن مستخدم مسبقاً (منع الإرسال المكرر)
    var isUsed = String(tokenRow[5] || '').trim();
    if (isUsed === 'نعم') {
      return { success: false, error: 'تم تقديم العذر مسبقاً عبر هذا الرابط. لا يمكن إرساله مرة أخرى.' };
    }

    var studentId = String(tokenRow[1]).trim();
    var stage = String(tokenRow[2]).trim();

    // جلب بيانات الطالب
    var studentInfo = getStudentInfoForParent_(studentId, stage);
    if (!studentInfo) return { success: false, error: 'لم يتم العثور على الطالب' };

    // حفظ العذر في شيت أعذار أولياء الأمور
    var excuseSheetName = 'اعذار_اولياء_الامور';
    var excuseSheet = findSheet_(ss, excuseSheetName);

    if (!excuseSheet) {
      excuseSheet = ss.insertSheet(excuseSheetName);
      excuseSheet.setRightToLeft(true);
      var headers = [
        'رقم_الطالب', 'اسم_الطالب', 'الصف', 'الفصل', 'المرحلة',
        'نص_العذر', 'مرفقات', 'تاريخ_الغياب', 'تاريخ_التقديم', 'وقت_التقديم',
        'الحالة', 'ملاحظات_المدرسة', 'الرمز'
      ];
      excuseSheet.appendRow(headers);
      excuseSheet.getRange(1, 1, 1, headers.length).setBackground('#7c3aed').setFontColor('#fff').setFontWeight('bold');
    }

    var now = new Date();
    var hijriDate = getHijriDateFull_(now).hijriStr || Utilities.formatDate(now, 'Asia/Riyadh', 'yyyy/MM/dd');

    // ★ تاريخ الغياب من النموذج (أو اليوم كافتراضي)
    var absenceDate = data.absenceDate || Utilities.formatDate(now, 'Asia/Riyadh', 'yyyy-MM-dd');

    // ★ ترقية تلقائية: إذا الشيت قديم (بدون عمود تاريخ_الغياب) — أضفه
    var existingHeaders = excuseSheet.getRange(1, 1, 1, excuseSheet.getLastColumn()).getValues()[0]
      .map(function(h) { return String(h).trim().replace(/\s+/g, '_'); });
    if (existingHeaders.indexOf('تاريخ_الغياب') === -1) {
      var insertAfter = existingHeaders.indexOf('مرفقات');
      if (insertAfter >= 0) {
        excuseSheet.insertColumnAfter(insertAfter + 1);
        excuseSheet.getRange(1, insertAfter + 2).setValue('تاريخ_الغياب')
          .setBackground('#7c3aed').setFontColor('#fff').setFontWeight('bold');
      }
    }

    var excuseRow = [
      studentId,
      studentInfo.name,
      studentInfo.grade,
      studentInfo.section,
      stage,
      sanitizeInput_(data.reason),
      data.hasAttachment ? 'نعم - تُسلم مع الطالب' : 'لا',
      absenceDate,
      hijriDate,
      Utilities.formatDate(now, 'Asia/Riyadh', 'HH:mm:ss'),
      'معلق',
      '',
      data.token
    ];

    excuseSheet.appendRow(excuseRow);

    // تحديث الرمز كمستخدم
    tokenSheet.getRange(tokenRowIdx, 6).setValue('نعم');

    return {
      success: true,
      message: 'تم إرسال العذر بنجاح. سيتم مراجعته من قبل إدارة المدرسة.'
    };

  } catch (e) {
    return { success: false, error: 'خطأ: ' + e.toString() };
  }
}

// =================================================================
// 6. عدد الأعذار المعلقة (لإشعار الداشبورد)
// =================================================================
function getPendingExcusesCount(stage) {
  try {
    var ss = getSpreadsheet_();
    var sheet = findSheet_(ss, 'اعذار_اولياء_الامور');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, count: 0 };

    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim().replace(/\s+/g, '_'); });
    var stageIdx = headers.indexOf('المرحلة');
    var statusIdx = headers.indexOf('الحالة');
    if (statusIdx < 0) return { success: true, count: 0 };

    var count = 0;
    for (var i = 1; i < data.length; i++) {
      var rowStage = stageIdx >= 0 ? String(data[i][stageIdx]).trim() : '';
      var rowStatus = String(data[i][statusIdx]).trim();
      if (rowStatus === 'معلق' && (!stage || rowStage === stage)) {
        count++;
      }
    }
    return { success: true, count: count };
  } catch (e) {
    return { success: true, count: 0 };
  }
}

// =================================================================
// 7. توليد رابط ولي الأمر (يُستدعى من sendAbsenceNotifications)
// =================================================================
function getParentExcuseLink_(studentId, stage) {
  var token = generateParentToken_(studentId, stage);
  var baseUrl = ScriptApp.getService().getUrl().replace(/\/dev$/, '/exec');
  return baseUrl + '?page=parent&token=' + token;
}

// =================================================================
// ★★★ دالة اختبار - شغّلها من محرر Apps Script ★★★
// تولّد رابط تجريبي لأول طالب في شيت الطلاب
// بدون إرسال واتساب - فقط تعطيك الرابط لتجربته بنفسك
// =================================================================
function TEST_ParentExcuseLink() {
  Logger.log('========= اختبار واجهة ولي الأمر =========');

  // 1. جلب أول طالب من شيت الطلاب
  ensureStudentsSheetsLoaded_();
  var ss = getSpreadsheet_();
  var stages = Object.keys(STUDENTS_SHEETS);
  var testStudent = null;
  var testStage = '';

  for (var s = 0; s < stages.length; s++) {
    var sheetName = STUDENTS_SHEETS[stages[s]];
    if (!sheetName) continue;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) continue;

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idCol = -1, nameCol = -1;
    for (var h = 0; h < headers.length; h++) {
      var hdr = String(headers[h]).trim().replace(/\s+/g, '_');
      if (hdr === 'رقم_الطالب' || hdr === 'رقم_الهوية') idCol = h;
      if (hdr === 'اسم_الطالب') nameCol = h;
    }
    if (idCol >= 0 && data.length > 1 && data[1][idCol]) {
      testStudent = { id: String(data[1][idCol]).trim(), name: String(data[1][nameCol] || '') };
      testStage = stages[s];
      break;
    }
  }

  if (!testStudent) {
    Logger.log('❌ لا يوجد طلاب في الشيتات!');
    return;
  }

  Logger.log('✅ الطالب: ' + testStudent.name + ' | الرقم: ' + testStudent.id + ' | المرحلة: ' + testStage);

  // 2. توليد الرابط
  var link = getParentExcuseLink_(testStudent.id, testStage);
  Logger.log('');
  Logger.log('🔗 رابط ولي الأمر (افتحه في المتصفح):');
  Logger.log(link);
  Logger.log('');
  Logger.log('⏳ صالح لمدة: 24 ساعة');

  // 3. اختبار جلب البيانات
  var token = link.split('token=')[1];
  var pageData = getParentExcusePageData_(token);
  Logger.log('');
  Logger.log('📋 بيانات الصفحة:');
  Logger.log(JSON.stringify(pageData, null, 2));

  Logger.log('');
  Logger.log('========= انتهى الاختبار =========');
  Logger.log('✅ افتح الرابط أعلاه في المتصفح لتجربة الواجهة');
  Logger.log('✅ اكتب عذر تجريبي وأرسله');
  Logger.log('✅ ثم تحقق من شيت "اعذار_اولياء_الامور" لترى العذر');
}
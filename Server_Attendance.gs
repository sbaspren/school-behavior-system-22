// =================================================================
// Server_Attendance.gs - دوال التأخر والاستئذان
// الإصدار المحدث - إنشاء تلقائي للشيتات
// =================================================================

// =================================================================
// الحصول على شيت التأخر (إنشاء تلقائي إذا غير موجود)
// =================================================================
function getLateSheet(stage) {
  var ss = getSpreadsheet_();
  var sheetName = 'سجل_التأخر_' + stage;
  var sheet = findSheet_(ss, sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.setRightToLeft(true);
    
    var headers = ['رقم_الطالب', 'اسم_الطالب', 'الصف', 'الفصل', 'رقم_الجوال', 'نوع_التأخر', 'الحصة', 'التاريخ_هجري', 'المسجل', 'وقت_الإدخال', 'تم_الإرسال'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#dc2626')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    // حماية عمود التاريخ الهجري من التحويل التلقائي
    sheet.getRange(1, 8, sheet.getMaxRows(), 1).setNumberFormat('@');

    // ضبط عرض الأعمدة
    sheet.setColumnWidth(1, 100);  // رقم الطالب
    sheet.setColumnWidth(2, 150);  // اسم الطالب
    sheet.setColumnWidth(3, 80);   // الصف
    sheet.setColumnWidth(4, 60);   // الفصل
    sheet.setColumnWidth(5, 110);  // رقم الجوال
    sheet.setColumnWidth(6, 100);  // نوع التأخر
    sheet.setColumnWidth(7, 60);   // الحصة
    sheet.setColumnWidth(8, 100);  // التاريخ هجري
    sheet.setColumnWidth(9, 100);  // المسجل
    sheet.setColumnWidth(10, 120); // وقت الإدخال
    sheet.setColumnWidth(11, 80);  // تم الإرسال
    
    // لون التبويب من SHEET_REGISTRY
    var _regColor = (SHEET_REGISTRY['التأخر'] || {}).color;
    if (_regColor) sheet.setTabColor(_regColor);
  }

  return sheet;
}

// =================================================================
// الحصول على شيت الاستئذان (إنشاء تلقائي إذا غير موجود)
// =================================================================
function getPermissionSheet(stage) {
  var ss = getSpreadsheet_();
  var sheetName = 'سجل_الاستئذان_' + stage;
  var sheet = findSheet_(ss, sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.setRightToLeft(true);
    
    var headers = ['رقم_الطالب', 'اسم_الطالب', 'الصف', 'الفصل', 'رقم_الجوال', 'وقت_الخروج', 'السبب', 'المستلم', 'المسؤول', 'التاريخ_هجري', 'المسجل', 'وقت_الإدخال', 'وقت_التأكيد', 'تم_الإرسال'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#7c3aed')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    // حماية عمود التاريخ الهجري من التحويل التلقائي
    sheet.getRange(1, 10, sheet.getMaxRows(), 1).setNumberFormat('@');

    // ضبط عرض الأعمدة
    sheet.setColumnWidth(1, 100);  // رقم الطالب
    sheet.setColumnWidth(2, 150);  // اسم الطالب
    sheet.setColumnWidth(3, 80);   // الصف
    sheet.setColumnWidth(4, 60);   // الفصل
    sheet.setColumnWidth(5, 110);  // رقم الجوال
    sheet.setColumnWidth(6, 80);   // وقت الخروج
    sheet.setColumnWidth(7, 150);  // السبب
    sheet.setColumnWidth(8, 120);  // المستلم
    sheet.setColumnWidth(9, 100);  // المسؤول
    sheet.setColumnWidth(10, 100); // التاريخ هجري
    sheet.setColumnWidth(11, 100); // المسجل
    sheet.setColumnWidth(12, 120); // وقت الإدخال
    sheet.setColumnWidth(13, 80);  // وقت التأكيد
    sheet.setColumnWidth(14, 80);  // تم الإرسال
    
    // لون التبويب من SHEET_REGISTRY
    var _regColor = (SHEET_REGISTRY['الاستئذان'] || {}).color;
    if (_regColor) sheet.setTabColor(_regColor);
  }

  return sheet;
}

// =================================================================
// جلب سجلات تأخر اليوم
// =================================================================
function getTodayLateRecords(stage) {
  var sheet = getLateSheet(stage);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: true, records: [] };
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var records = filterTodayRecords_(data, headers, 'وقت_الإدخال');
  
  return { success: true, records: records };
}

// =================================================================
// جلب سجلات استئذان اليوم
// =================================================================
function getTodayPermissionRecords(stage) {
  var sheet = getPermissionSheet(stage);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: true, records: [] };
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var records = filterTodayRecords_(data, headers, 'وقت_الإدخال');
  
  return { success: true, records: records };
}

// =================================================================
// حفظ سجلات تأخر (متعدد)
// =================================================================
function saveLateRecords(data) {
  try {
    var stage = data.stage;
    if (!stage) throw new Error("المرحلة الدراسية مطلوبة");
    var sheet = getLateSheet(stage);
    
    var currentUser = 'الوكيل';
    
    var now = new Date();
    var hijriDate = getHijriDate_(now);
    var savedCount = 0;
    
    var students = data.students || [];
    
    for (var i = 0; i < students.length; i++) {
      var student = students[i];
      sheet.appendRow([
        student.studentId,
        sanitizeInput_(student.studentName),
        student.grade,
        student.class,
        student.phone || '',
        sanitizeInput_(data.lateType || 'تأخر صباحي'),
        sanitizeInput_(data.period || ''),
        hijriDate,
        currentUser,
        now,
        'لا'
      ]);
      savedCount++;
    }
    
    return { success: true, message: 'تم تسجيل ' + savedCount + ' طالب متأخر', count: savedCount };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// حفظ سجلات استئذان (متعدد)
// =================================================================
function savePermissionRecords(data) {
  try {
    var stage = data.stage;
    if (!stage) throw new Error("المرحلة الدراسية مطلوبة");
    var sheet = getPermissionSheet(stage);
    
    var currentUser = 'الوكيل';
    
    var now = new Date();
    var hijriDate = getHijriDate_(now);
    var savedCount = 0;
    
    var students = data.students || [];
    
    for (var i = 0; i < students.length; i++) {
      var student = students[i];
      sheet.appendRow([
        student.studentId,
        sanitizeInput_(student.studentName),
        student.grade,
        student.class,
        student.phone || '',
        data.exitTime || Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm'),
        sanitizeInput_(data.reason || ''),
        sanitizeInput_(data.receiver || ''),
        sanitizeInput_(data.responsible || ''),
        hijriDate,
        currentUser,
        now,
        '',
        'لا'
      ]);
      savedCount++;
    }
    
    return { success: true, message: 'تم تسجيل ' + savedCount + ' طالب مستأذن', count: savedCount };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// حفظ سجل تأخر واحد
// =================================================================
function saveLateRecord(data) {
  var singleData = {
    stage: data.stage,
    lateType: data.lateType || 'تأخر صباحي',
    period: data.period || '',
    students: [{
      studentId: data.studentId,
      studentName: data.studentName,
      grade: data.grade,
      class: data.class,
      phone: data.phone
    }]
  };
  return saveLateRecords(singleData);
}

// =================================================================
// حفظ سجل استئذان واحد
// =================================================================
function savePermissionRecord(data) {
  var singleData = {
    stage: data.stage,
    exitTime: data.exitTime,
    reason: data.reason,
    receiver: data.receiver,
    responsible: data.responsible,
    students: [{
      studentId: data.studentId,
      studentName: data.studentName,
      grade: data.grade,
      class: data.class,
      phone: data.phone
    }]
  };
  return savePermissionRecords(singleData);
}

// =================================================================
// جلب أرشيف التأخر
// =================================================================
function getLateArchive(stage, dateFrom, dateTo) {
  var sheet = getLateSheet(stage);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: true, records: [] };
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var records = [];
  
  var fromDate = dateFrom ? new Date(dateFrom) : null;
  var toDate = dateTo ? new Date(dateTo) : null;
  if (toDate) toDate.setHours(23, 59, 59);
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    
    var recordDate = row[headers.indexOf('وقت_الإدخال')];
    if (recordDate instanceof Date) {
      if (fromDate && recordDate < fromDate) continue;
      if (toDate && recordDate > toDate) continue;
    }
    
    var record = { rowIndex: i };
    for (var j = 0; j < headers.length; j++) {
      var value = row[j];
      if (value instanceof Date) {
        if (headers[j] === 'التاريخ_هجري') {
          value = readHijriCellValue_(value);
        } else {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
        }
      }
      record[headers[j]] = String(value || '');
    }
    records.push(record);
  }

  return { success: true, records: records };
}

// =================================================================
// جلب أرشيف الاستئذان
// =================================================================
function getPermissionArchive(stage, dateFrom, dateTo) {
  var sheet = getPermissionSheet(stage);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: true, records: [] };
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var records = [];
  
  var fromDate = dateFrom ? new Date(dateFrom) : null;
  var toDate = dateTo ? new Date(dateTo) : null;
  if (toDate) toDate.setHours(23, 59, 59);
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    
    var recordDate = row[headers.indexOf('وقت_الإدخال')];
    if (recordDate instanceof Date) {
      if (fromDate && recordDate < fromDate) continue;
      if (toDate && recordDate > toDate) continue;
    }
    
    var record = { rowIndex: i };
    for (var j = 0; j < headers.length; j++) {
      var value = row[j];
      if (value instanceof Date) {
        if (headers[j] === 'التاريخ_هجري') {
          value = readHijriCellValue_(value);
        } else if (headers[j] === 'وقت_الخروج' || headers[j] === 'وقت_التأكيد') {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
        } else {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
        }
      }
      record[headers[j]] = String(value || '');
    }
    records.push(record);
  }

  return { success: true, records: records };
}

// =================================================================
// حذف سجل تأخر
// =================================================================
function deleteLateRecord(stage, rowIndex) {
  try {
    var sheet = getLateSheet(stage);
    sheet.deleteRow(rowIndex + 1);
    return { success: true, message: 'تم الحذف' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// حذف سجل استئذان
// =================================================================
function deletePermissionRecord(stage, rowIndex) {
  try {
    var sheet = getPermissionSheet(stage);
    sheet.deleteRow(rowIndex + 1);
    return { success: true, message: 'تم الحذف' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// تحديث حالة الإرسال - تأخر
// =================================================================
function updateLateSentStatus(stage, rowIndex) {
  try {
    var sheet = getLateSheet(stage);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var sentCol = headers.indexOf('تم_الإرسال') + 1;

    if (sentCol > 0) {
      sheet.getRange(rowIndex + 1, sentCol).setValue('نعم');
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// تحديث حالة الإرسال - تأخر (جماعي)
// =================================================================
function updateLateSentStatusBatch(stage, rowIndices) {
  try {
    if (!rowIndices || rowIndices.length === 0) return { success: true };
    var sheet = getLateSheet(stage);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var sentCol = headers.indexOf('تم_الإرسال') + 1;
    if (sentCol <= 0) return { success: false, error: 'عمود تم_الإرسال غير موجود' };
    var lastRow = sheet.getLastRow();
    for (var i = 0; i < rowIndices.length; i++) {
      var row = parseInt(rowIndices[i]);
      if (isNaN(row) || row < 1 || row + 1 > lastRow) continue;
      sheet.getRange(row + 1, sentCol).setValue('نعم');
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// تحديث حالة الإرسال - استئذان
// =================================================================
function updatePermissionSentStatus(stage, rowIndex) {
  try {
    var sheet = getPermissionSheet(stage);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var sentCol = headers.indexOf('تم_الإرسال') + 1;

    if (sentCol > 0) {
      sheet.getRange(rowIndex + 1, sentCol).setValue('نعم');
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// تحديث حالة الإرسال - استئذان (جماعي)
// =================================================================
function updatePermissionSentStatusBatch(stage, rowIndices) {
  try {
    if (!rowIndices || rowIndices.length === 0) return { success: true };
    var sheet = getPermissionSheet(stage);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var sentCol = headers.indexOf('تم_الإرسال') + 1;
    if (sentCol <= 0) return { success: false, error: 'عمود تم_الإرسال غير موجود' };
    var lastRow = sheet.getLastRow();
    for (var i = 0; i < rowIndices.length; i++) {
      var row = parseInt(rowIndices[i]);
      if (isNaN(row) || row < 1 || row + 1 > lastRow) continue;
      sheet.getRange(row + 1, sentCol).setValue('نعم');
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// تأكيد خروج الطالب (للحارس)
// =================================================================
// ★ confirmStudentExit — محذوفة من هنا لأنها مكررة
// النسخة الرسمية في Server_StaffInput.gs بترتيب (rowIndex, stage)
// وهي المستخدمة فعلياً من GuardDisplay.html

// =================================================================
// جلب إحصائيات اليوم
// =================================================================
function getTodayAttendanceStats(stage) {
  var lateResult = getTodayLateRecords(stage);
  var permResult = getTodayPermissionRecords(stage);
  
  var lateRecords = lateResult.success ? lateResult.records : [];
  var permRecords = permResult.success ? permResult.records : [];
  
  return {
    success: true,
    stats: {
      lateCount: lateRecords.length,
      lateSent: lateRecords.filter(function(r) { return r['تم_الإرسال'] === 'نعم'; }).length,
      permissionCount: permRecords.length,
      permissionSent: permRecords.filter(function(r) { return r['تم_الإرسال'] === 'نعم'; }).length
    }
  };
}

// =================================================================
// جلب الاستئذانات المعلقة (للحارس)
// =================================================================
function getPendingPermissions(stage) {
  var result = getTodayPermissionRecords(stage);
  
  if (!result.success) return result;
  
  var pending = result.records.filter(function(r) {
    return !r['وقت_التأكيد'] || r['وقت_التأكيد'] === '';
  });
  
  return { success: true, records: pending };
}

// getHijriDate_() → مركزية في Config.gs

// =================================================================
// دالة اختبار
// =================================================================
function TEST_Attendance() {
  Logger.log("=== اختبار التأخر ===");
  var lateRecords = getTodayLateRecords('متوسط');
  Logger.log("سجلات التأخر: " + JSON.stringify(lateRecords));
  
  Logger.log("=== اختبار الاستئذان ===");
  var permRecords = getTodayPermissionRecords('متوسط');
  Logger.log("سجلات الاستئذان: " + JSON.stringify(permRecords));
}
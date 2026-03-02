// =================================================================
// Server_EducationalNotes.gs - دوال الملاحظات التربوية
// الإصدار المحدث - إنشاء تلقائي للشيتات
// =================================================================

// =================================================================
// الحصول على شيت الملاحظات التربوية (إنشاء تلقائي إذا غير موجود)
// =================================================================
function getEducationalNotesSheet(stage) {
  var ss = getSpreadsheet_();
  var sheetName = 'سجل_الملاحظات_التربوية_' + stage;
  var sheet = findSheet_(ss, sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.setRightToLeft(true);
    
    var headers = [
      'رقم الطالب',
      'اسم الطالب', 
      'الصف',
      'الفصل',
      'رقم الجوال',
      'نوع الملاحظة',
      'التفاصيل',
      'المعلم/المسجل',
      'التاريخ',
      'وقت الإدخال',
      'تم الإرسال'
    ];
    
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1e3a5f')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    // حماية عمود التاريخ من التحويل التلقائي
    sheet.getRange(1, 9, sheet.getMaxRows(), 1).setNumberFormat('@');

    // ضبط عرض الأعمدة
    sheet.setColumnWidth(1, 100);  // رقم الطالب
    sheet.setColumnWidth(2, 150);  // اسم الطالب
    sheet.setColumnWidth(3, 80);   // الصف
    sheet.setColumnWidth(4, 60);   // الفصل
    sheet.setColumnWidth(5, 110);  // رقم الجوال
    sheet.setColumnWidth(6, 150);  // نوع الملاحظة
    sheet.setColumnWidth(7, 200);  // التفاصيل
    sheet.setColumnWidth(8, 100);  // المعلم
    sheet.setColumnWidth(9, 100);  // التاريخ
    sheet.setColumnWidth(10, 100); // وقت الإدخال
    sheet.setColumnWidth(11, 80);  // تم الإرسال
    
    // لون التبويب من SHEET_REGISTRY
    var _regColor = (SHEET_REGISTRY['الملاحظات_التربوية'] || {}).color;
    if (_regColor) sheet.setTabColor(_regColor);
  }
  
  return sheet;
}

// =================================================================
// جلب سجلات الملاحظات التربوية (الكل - للأرشيف)
// =================================================================
function getEducationalNotesRecords(stage) {
  var sheet = getEducationalNotesSheet(stage);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: true, records: [] };
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var records = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] && !row[1]) continue;
    
    var record = { rowIndex: i };
    for (var j = 0; j < headers.length; j++) {
      var key = String(headers[j] || '').trim().replace(/\s+/g, '_');
      if (!key) continue;
      var value = row[j];
      if (value instanceof Date) {
        if (key === 'التاريخ') {
          value = readHijriCellValue_(value);
        } else {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
        }
      }
      record[key] = String(value || '');
    }
    records.push(record);
  }
  
  return { success: true, records: records };
}

// =================================================================
// جلب ملاحظات اليوم فقط (نفس نمط getTodayLateRecords)
// يستخدم filterTodayRecords_ المركزية من Config.gs
// =================================================================
function getTodayEducationalNotesRecords(stage) {
  try {
    var sheet = getEducationalNotesSheet(stage);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, records: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var records = filterTodayRecords_(data, headers, 'وقت_الإدخال');
    
    return { success: true, records: records };
  } catch (e) {
    Logger.log('خطأ في تحميل الملاحظات: ' + e.toString());
    return { success: false, records: [], error: e.toString() };
  }
}

function saveEducationalNote(noteData) {
  try {
    var stage = noteData.stage;
    if (!stage) throw new Error("المرحلة الدراسية مطلوبة");
    var sheet = getEducationalNotesSheet(stage);
    
    var currentUser = 'الوكيل';
    
    var now = new Date();
    var hijriDate = getHijriDate_(now);
    
    sheet.appendRow([
      noteData.studentId,
      sanitizeInput_(noteData.studentName),
      sanitizeInput_(noteData.grade),
      sanitizeInput_(noteData.class),
      noteData.phone || '',
      sanitizeInput_(noteData.noteType),
      sanitizeInput_(noteData.details || ''),
      currentUser,
      hijriDate,
      now,
      'لا'
    ]);
    
    return { success: true, message: 'تم حفظ الملاحظة بنجاح' };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// BATCH SAVE - حفظ ملاحظة لعدة طلاب دفعة واحدة
// =================================================================
function saveEducationalNotesBatch(data) {
  try {
    var stage = data.stage;
    if (!stage) throw new Error("المرحلة الدراسية مطلوبة");
    var sheet = getEducationalNotesSheet(stage);

    var currentUser = 'الوكيل';
    var now = new Date();
    var hijriDate = getHijriDate_(now);

    var students = data.students || [];
    if (students.length === 0) throw new Error("لم يتم اختيار طلاب");

    var rows = [];
    for (var i = 0; i < students.length; i++) {
      var s = students[i];
      rows.push([
        s.studentId,
        sanitizeInput_(s.studentName),
        sanitizeInput_(s.grade),
        sanitizeInput_(s.class),
        s.phone || '',
        sanitizeInput_(data.noteType),
        sanitizeInput_(data.details || ''),
        currentUser,
        hijriDate,
        now,
        'لا'
      ]);
    }

    if (rows.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 9, rows.length, 1).setNumberFormat('@');
      sheet.getRange(startRow, 1, rows.length, 11).setValues(rows);
    }

    return { success: true, message: 'تم حفظ ' + rows.length + ' ملاحظة بنجاح', count: rows.length };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// تحديث حالة الإرسال
// =================================================================
function updateEduNoteSentStatus(stage, rowIndicesOrStudentId, noteType, date) {
  try {
    var sheet = getEducationalNotesSheet(stage);
    if (!sheet) return { success: false, error: 'الشيت غير موجود' };
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var sentCol = headers.indexOf('تم الإرسال');
    if (sentCol === -1) sentCol = headers.indexOf('تم_الإرسال');
    if (sentCol === -1) return { success: false, error: 'عمود الإرسال غير موجود' };
    
    // دعم الطريقتين: row indices (array) أو student ID (string/number)
    if (Array.isArray(rowIndicesOrStudentId)) {
      // طريقة row indices — الفهرس + 1 = رقم الصف الفعلي في الشيت
      var rowIndices = rowIndicesOrStudentId;
      var lastRow = sheet.getLastRow();
      for (var i = 0; i < rowIndices.length; i++) {
        var row = parseInt(rowIndices[i]);
        if (isNaN(row) || row < 1 || row + 1 > lastRow) continue;
        sheet.getRange(row + 1, sentCol + 1).setValue('نعم');
      }
    } else {
      // طريقة student ID القديمة (للتوافقية)
      var studentId = rowIndicesOrStudentId;
      var data = sheet.getDataRange().getValues();
      var studentIdCol = headers.indexOf('رقم الطالب');
      if (studentIdCol === -1) studentIdCol = headers.indexOf('رقم_الطالب');
      for (var j = 1; j < data.length; j++) {
        if (String(data[j][studentIdCol]) === String(studentId)) {
          sheet.getRange(j + 1, sentCol + 1).setValue('نعم');
        }
      }
    }
    
    // مسح الكاش
    try {
      var cacheKey = 'eduNotes_' + stage + '_' + new Date().toLocaleDateString('en-US');
      CacheService.getScriptCache().remove(cacheKey);
    } catch(ce) {}
    
    return { success: true };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// حذف ملاحظة تربوية
// =================================================================
function deleteEducationalNote(stage, rowIndex) {
  try {
    var sheet = getEducationalNotesSheet(stage);
    
    if (!sheet) return { success: false, error: 'الشيت غير موجود' };
    
    sheet.deleteRow(rowIndex + 1);
    
    return { success: true, message: 'تم حذف الملاحظة' };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// جلب إحصائيات الملاحظات التربوية
// =================================================================
function getEducationalNotesStats(stage) {
  try {
    var result = getEducationalNotesRecords(stage);
    var records = result.records || [];
    
    var stats = {
      total: records.length,
      sent: 0,
      notSent: 0,
      byType: {},
      byGrade: {},
      todayCount: 0
    };
    
    var today = new Date().toDateString();
    
    for (var i = 0; i < records.length; i++) {
      var r = records[i];
      
      if (r['تم_الإرسال'] === 'نعم') {
        stats.sent++;
      } else {
        stats.notSent++;
      }

      var type = r['نوع_الملاحظة'] || 'غير محدد';
      stats.byType[type] = (stats.byType[type] || 0) + 1;

      var grade = r['الصف'] || 'غير محدد';
      stats.byGrade[grade] = (stats.byGrade[grade] || 0) + 1;

      var recordDate = new Date(r['وقت_الإدخال'] || r['التاريخ']);
      if (recordDate.toDateString() === today) {
        stats.todayCount++;
      }
    }
    
    return { success: true, stats: stats };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// جلب أنواع الملاحظات التربوية
// =================================================================
function getEducationalNotesTypes(stage) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName("أنواع_الملاحظات_التربوية");
    
    if (!sheet || sheet.getLastRow() < 2) {
      return getDefaultEduNotesTypes();
    }
    
    var data = sheet.getDataRange().getValues();
    var types = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0] === stage && row[1]) {
        types.push(row[1]);
      }
    }
    
    return types.length > 0 ? types : getDefaultEduNotesTypes();
    
  } catch (e) {
    return getDefaultEduNotesTypes();
  }
}

// =================================================================
// حفظ أنواع الملاحظات التربوية
// =================================================================
function saveEducationalNotesTypes(stage, types) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName("أنواع_الملاحظات_التربوية");
    
    if (!sheet) {
      sheet = ss.insertSheet("أنواع_الملاحظات_التربوية");
      sheet.setRightToLeft(true);
      sheet.appendRow(['المرحلة', 'نوع الملاحظة', 'تاريخ الإضافة']);
      sheet.getRange(1, 1, 1, 3).setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
    }
    
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === stage) {
        sheet.deleteRow(i + 1);
      }
    }
    
    var now = new Date();
    for (var j = 0; j < types.length; j++) {
      sheet.appendRow([stage, sanitizeInput_(types[j]), now]);
    }
    
    return { success: true, message: 'تم حفظ الأنواع بنجاح' };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// الأنواع الافتراضية
// =================================================================
function getDefaultEduNotesTypes() {
  return [
    'عدم حل الواجب',
    'عدم الحفظ',
    'عدم المشاركة والتفاعل',
    'عدم إحضار الكتاب الدراسي',
    'عدم إحضار الدفتر',
    'كثرة السرحان داخل الفصل',
    'عدم إحضار أدوات الرسم',
    'عدم إحضار الأدوات الهندسية',
    'عدم إحضار الملابس الرياضية',
    'النوم داخل الفصل',
    'عدم تدوين الملاحظات مع المعلم',
    'إهمال تسليم البحوث والمشاريع',
    'عدم المذاكرة للاختبارات القصيرة',
    'الانشغال بمادة أخرى أثناء الحصة',
    'عدم تصحيح الأخطاء في الدفتر',
    'عدم إحضار ملف الإنجاز'
  ];
}

// getHijriDate_() → مركزية في Config.gs

// =================================================================
// دالة اختبار
// =================================================================
function TEST_EduNotes() {
  Logger.log("=== اختبار الملاحظات التربوية ===");
  var records = getEducationalNotesRecords('متوسط');
  Logger.log("النتيجة: " + JSON.stringify(records));
  Logger.log("عدد السجلات: " + (records.records ? records.records.length : 0));
}
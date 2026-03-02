// =================================================================
// نظام توثيق التواصل - Server_Communication.gs
// =================================================================

/**
 * الحصول على أو إنشاء ورقة سجل التواصل
 */
// =================================================================
// الحصول على شيت سجل التواصل حسب المرحلة (إنشاء تلقائي إذا غير موجود)
// =================================================================
function getCommunicationSheet(stage) {
  var ss = getSpreadsheet_();
  
  // تحديد اسم الشيت حسب المرحلة
  var sheetName = 'سجل_التواصل_' + stage;
  var sheet = findSheet_(ss, sheetName);
  
  // إذا الشيت غير موجود، أنشئه تلقائياً
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    
    // إنشاء الترويسة
    var headers = [
      'م',
      'التاريخ الهجري',
      'التاريخ الميلادي', 
      'الوقت',
      'رقم الطالب',
      'اسم الطالب',
      'الصف',
      'الفصل',
      'رقم الجوال',
      'نوع الرسالة',
      'عنوان الرسالة',
      'نص الرسالة',
      'حالة الإرسال',
      'المرسل',
      'ملاحظات'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4a5568')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    // تجميد الصف الأول
    sheet.setFrozenRows(1);

    // حماية عمود التاريخ الهجري من التحويل التلقائي
    sheet.getRange(1, 2, sheet.getMaxRows(), 1).setNumberFormat('@');

    // ضبط عرض الأعمدة
    sheet.setColumnWidth(1, 40);   // م
    sheet.setColumnWidth(2, 100);  // التاريخ الهجري
    sheet.setColumnWidth(3, 100);  // التاريخ الميلادي
    sheet.setColumnWidth(4, 70);   // الوقت
    sheet.setColumnWidth(5, 100);  // رقم الطالب
    sheet.setColumnWidth(6, 150);  // اسم الطالب
    sheet.setColumnWidth(7, 80);   // الصف
    sheet.setColumnWidth(8, 60);   // الفصل
    sheet.setColumnWidth(9, 110);  // رقم الجوال
    sheet.setColumnWidth(10, 80);  // نوع الرسالة
    sheet.setColumnWidth(11, 150); // عنوان الرسالة
    sheet.setColumnWidth(12, 300); // نص الرسالة
    sheet.setColumnWidth(13, 80);  // حالة الإرسال
    sheet.setColumnWidth(14, 100); // المرسل
    sheet.setColumnWidth(15, 150); // ملاحظات
    
    // لون التبويب من SHEET_REGISTRY
    var _regColor = (SHEET_REGISTRY['التواصل'] || {}).color;
    if (_regColor) sheet.setTabColor(_regColor);
  }
  
  return sheet;
}

/**
 * تسجيل رسالة في سجل التواصل
 */
// =================================================================
// تسجيل رسالة في سجل التواصل
// =================================================================
function logCommunication(data, stageParam) {
  try {
    var stage = stageParam || data.stage;
    if (!stage) throw new Error("المرحلة الدراسية مطلوبة");
    var sheet = getCommunicationSheet(stage);
    var lastRow = sheet.getLastRow();
    var newId = lastRow;
    
    var now = new Date();
    var hijriDate = getHijriDate_(now);
    var miladiDate = Utilities.formatDate(now, 'Asia/Riyadh', 'yyyy/MM/dd');
    var time = Utilities.formatDate(now, 'Asia/Riyadh', 'HH:mm');
    
    var row = [
      newId,
      hijriDate,
      miladiDate,
      time,
      data.studentId || '',
      data.studentName || '',
      data.grade || '',
      data.class || '',
      data.phone || '',
      data.messageType || '',
      sanitizeInput_(data.messageTitle || ''),
      sanitizeInput_(data.messageContent || ''),
      data.status || 'جاري الإرسال',
      sanitizeInput_(data.sender || ''),
      sanitizeInput_(data.notes || '')
    ];
    
    sheet.appendRow(row);
    
    return { success: true, logId: newId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * تحديث حالة الإرسال في السجل
 */
function updateCommunicationStatus(logId, status, notes, stage) {
  try {
    if (!stage) throw new Error("المرحلة الدراسية مطلوبة");
    const sheet = getCommunicationSheet(stage);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == logId) {
        sheet.getRange(i + 1, 13).setValue(status); // حالة الإرسال
        if (notes) {
          sheet.getRange(i + 1, 15).setValue(sanitizeInput_(notes)); // ملاحظات
        }
        return { success: true };
      }
    }
    
    return { success: false, error: 'لم يتم العثور على السجل' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ملاحظة: دالة sendWhatsAppWithLog موجودة في Server_WhatsApp.gs

/**
 * تحديث حالة الإرسال في السجل الأصلي
 */
function updateOriginalRecordSentStatus(sheetName, rowIndex) {
  try {
    const ss = getSpreadsheet_();
    const sheet = findSheet_(ss, sheetName);
    if (!sheet) return;
    
    // البحث عن عمود "تم الإرسال"
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const sentColIndex = headers.indexOf('تم الإرسال');
    
    if (sentColIndex !== -1) {
      sheet.getRange(rowIndex, sentColIndex + 1).setValue('نعم');
    }
  } catch (e) {
    console.error('خطأ في تحديث حالة الإرسال الأصلية:', e);
  }
}

/**
 * جلب سجل التواصل مع الفلترة
 */
function getCommunicationLog(filters) {
  try {
    const stage = (filters && filters.stage) ? filters.stage : '';
    if (!stage) throw new Error("المرحلة الدراسية مطلوبة");
    const sheet = getCommunicationSheet(stage);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return [];
    }
    
    let records = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const record = {
        id: row[0],
        hijriDate: row[1],
        miladiDate: row[2],
        time: row[3],
        studentId: row[4],
        studentName: row[5],
        grade: row[6],
        class: row[7],
        phone: row[8],
        messageType: row[9],
        messageTitle: row[10],
        messageContent: row[11],
        status: row[12],
        sender: row[13],
        notes: row[14]
      };
      
      let include = true;
      
      if (filters) {
        if (filters.studentId && record.studentId != filters.studentId) include = false;
        if (filters.messageType && record.messageType !== filters.messageType) include = false;
        if (filters.status && record.status !== filters.status) include = false;
        if (filters.dateFrom && record.miladiDate < filters.dateFrom) include = false;
        if (filters.dateTo && record.miladiDate > filters.dateTo) include = false;
        if (filters.grade && record.grade !== filters.grade) include = false;
      }
      
      if (include) {
        records.push(record);
      }
    }
    
    records.reverse();
    
    return records;
  } catch (e) {
    console.error('خطأ في getCommunicationLog:', e);
    return [];
  }
}
/**
 * جلب إحصائيات التواصل
 */
// =================================================================
// جلب إحصائيات التواصل حسب المرحلة
// =================================================================
function getCommunicationStats(stage) {
  try {
    var sheet = getCommunicationSheet(stage);
    var data = sheet.getDataRange().getValues();
    
    var stats = {
      total: 0,
      sent: 0,
      failed: 0,
      byType: {
        'مخالفة': 0,
        'ملاحظة': 0,
        'غياب': 0,
        'تأخر': 0,
        'أخرى': 0
      },
      todayCount: 0,
      weekCount: 0
    };
    
    var today = new Date();
    var todayStr = Utilities.formatDate(today, 'Asia/Riyadh', 'yyyy/MM/dd');
    
    var weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
    var weekAgoStr = Utilities.formatDate(weekAgo, 'Asia/Riyadh', 'yyyy/MM/dd');
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[5]) continue;
      
      stats.total++;
      
      // حالة الإرسال
      var status = String(row[12] || '');
      if (status.indexOf('تم') !== -1) stats.sent++;
      else if (status.indexOf('فشل') !== -1) stats.failed++;
      
      // نوع الرسالة
      var msgType = row[9] || 'أخرى';
      if (stats.byType[msgType] !== undefined) {
        stats.byType[msgType]++;
      } else {
        stats.byType['أخرى']++;
      }
      
      // إحصائيات الفترة
      var recordDate = row[2];
      if (recordDate instanceof Date) {
        recordDate = Utilities.formatDate(recordDate, 'Asia/Riyadh', 'yyyy/MM/dd');
      }
      if (recordDate === todayStr) stats.todayCount++;
      if (recordDate >= weekAgoStr) stats.weekCount++;
    }
    
    return { success: true, stats: stats };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * حذف سجل تواصل (للمدير فقط)
 */
function deleteCommunicationRecord(logId, stage) {
  try {
    var authCheck = checkUserPermission('admin');
    if (!authCheck.hasPermission) {
      return { success: false, error: 'غير مصرح: ' + authCheck.reason };
    }
    if (!stage) throw new Error("المرحلة الدراسية مطلوبة");
    const sheet = getCommunicationSheet(stage);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == logId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    
    return { success: false, error: 'لم يتم العثور على السجل' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * تصدير سجل التواصل لملف Excel
 */
function exportCommunicationLog(filters) {
  try {
    const records = getCommunicationLog(filters);
    if (!records || records.length === 0) {
      return { success: false, error: 'لا توجد سجلات للتصدير' };
    }

    // إنشاء ملف جديد
    const ss = SpreadsheetApp.create('سجل_التواصل_' + new Date().toLocaleDateString('ar-SA'));
    const sheet = ss.getActiveSheet();

    // الترويسة
    const headers = ['م', 'التاريخ', 'الوقت', 'الطالب', 'الصف', 'الجوال', 'النوع', 'العنوان', 'الرسالة', 'الحالة', 'المرسل'];
    sheet.appendRow(headers);

    // البيانات
    records.forEach((rec, i) => {
      sheet.appendRow([
        i + 1,
        rec.hijriDate,
        rec.time,
        rec.studentName,
        rec.grade + '/' + rec.class,
        rec.phone,
        rec.messageType,
        rec.messageTitle,
        rec.messageContent,
        rec.status,
        rec.sender
      ]);
    });
    
    return { success: true, url: ss.getUrl() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// getHijriDate_() → مركزية في Config.gs

function TEST_LogCommunication() {
  Logger.log("=== اختبار تسجيل التواصل ===");
  
  const testData = {
    studentId: '123',
    studentName: 'طالب تجريبي',
    grade: 'الأول',
    class: 'أ',
    phone: '0551234567',
    messageType: 'اختبار',
    messageTitle: 'رسالة تجريبية',
    messageContent: 'هذا نص تجريبي',
    sender: 'الوكيل',
    status: 'تجربة'
  };
  
  const result = logCommunication(testData);
  Logger.log("النتيجة: " + JSON.stringify(result));
}
function TEST_GetCommunicationLog() {
  const result = getCommunicationLog({});
  Logger.log("النتيجة: " + JSON.stringify(result));
}

function getCommunicationRecords(stage) {
  try {
    if (!stage) throw new Error("المرحلة الدراسية مطلوبة");
    var sheet = getCommunicationSheet(stage);
    var data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return { success: true, records: [] };
    }
    
    var records = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      records.push({
        id: row[0],
        hijriDate: row[1],
        miladiDate: row[2],
        time: row[3],
        studentId: row[4],
        studentName: row[5],
        grade: row[6],
        class: row[7],
        phone: row[8],
        messageType: row[9],
        messageTitle: row[10],
        messageContent: row[11],
        status: row[12],
        sender: row[13],
        notes: row[14]
      });
    }
    
    records.reverse();
    return { success: true, records: records };
  } catch (e) {
    return { success: false, records: [], error: e.toString() };
  }
}

function TEST_GetCommunicationRecords() {
  const result = getCommunicationRecords();
  Logger.log("النتيجة: " + JSON.stringify(result));
  Logger.log("عدد السجلات: " + result.length);
}

// =================================================================
// جلب سجلات التواصل حسب المرحلة
// =================================================================
function loadCommRecords(stage) {
  var ss = getSpreadsheet_();
  
  // تحديد اسم الشيت حسب المرحلة
  var sheetName = 'سجل_التواصل_' + stage;
  var sheet = findSheet_(ss, sheetName);
  
  // إذا الشيت غير موجود، أنشئه
  if (!sheet) {
    sheet = getCommunicationSheet(stage);
  }
  
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: true, records: [] };
  }
  
  var data = sheet.getDataRange().getValues();
  var records = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] && !row[5]) continue;
    
    var hijriDate = row[1];
    var miladiDate = row[2];
    var time = row[3];
    
    if (hijriDate instanceof Date) {
      hijriDate = readHijriCellValue_(hijriDate);
    }
    if (miladiDate instanceof Date) {
      miladiDate = Utilities.formatDate(miladiDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    }
    if (time instanceof Date) {
      time = Utilities.formatDate(time, Session.getScriptTimeZone(), 'HH:mm');
    }
    
    records.push({
      id: String(row[0] || i),
      hijriDate: String(hijriDate || ''),
      miladiDate: String(miladiDate || ''),
      time: String(time || ''),
      studentId: String(row[4] || ''),
      studentName: String(row[5] || ''),
      grade: String(row[6] || ''),
      class: String(row[7] || ''),
      phone: String(row[8] || ''),
      messageType: String(row[9] || ''),
      messageTitle: String(row[10] || ''),
      messageContent: String(row[11] || ''),
      status: String(row[12] || ''),
      sender: String(row[13] || ''),
      notes: String(row[14] || '')
    });
  }
  
  records.reverse();
  return { success: true, records: records };
}

function TEST_loadCommRecords() {
  var result = loadCommRecords();
  Logger.log("النتيجة: " + JSON.stringify(result));
  Logger.log("العدد: " + result.length);
}
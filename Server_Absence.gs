// =================================================================
// Server_Absence.gs - معالجة الغياب
// الإصدار المحدث - إنشاء تلقائي للشيتات
// =================================================================

// =================================================================
// الحصول على شيت الغياب (إنشاء تلقائي إذا غير موجود)
// =================================================================
function getAbsenceSheet(stage) {
  var ss = getSpreadsheet_();
  var sheetName = 'سجل_الغياب_' + stage;
  var sheet = findSheet_(ss, sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.setRightToLeft(true);
    
    var headers = ['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'غياب بعذر', 'غياب بدون عذر', 'تأخير', 'آخر تحديث'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#f59e0b')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    
    // ضبط عرض الأعمدة
    sheet.setColumnWidth(1, 100);  // رقم الطالب
    sheet.setColumnWidth(2, 150);  // اسم الطالب
    sheet.setColumnWidth(3, 80);   // الصف
    sheet.setColumnWidth(4, 60);   // الفصل
    sheet.setColumnWidth(5, 100);  // غياب بعذر
    sheet.setColumnWidth(6, 120);  // غياب بدون عذر
    sheet.setColumnWidth(7, 80);   // تأخير
    sheet.setColumnWidth(8, 120);  // آخر تحديث
    
    // لون التبويب من SHEET_REGISTRY
    var _regColor = (SHEET_REGISTRY['الغياب'] || {}).color;
    if (_regColor) sheet.setTabColor(_regColor);
  }
  
  return sheet;
}

// =================================================================
// معالجة ملف نور وحفظ البيانات
// =================================================================
function processNoorAbsenceFile(fileContent, stage) {
  try {
    var ss = getSpreadsheet_();
    
    // جلب الطلاب من شيتات المراحل
    var allSheets = getAllStudentsSheets_();
    var masterMap = {};
    
    for (var sh = 0; sh < allSheets.length; sh++) {
      var studentsSheet = allSheets[sh].sheet;
      var stg = allSheets[sh].stage;
      var studentsData = studentsSheet.getDataRange().getValues();
      if (studentsData.length < 2) continue;
      var headers = studentsData[0];
      
      var idCol = headers.indexOf('رقم الطالب'); if (idCol < 0) idCol = headers.indexOf('رقم_الطالب'); if (idCol < 0) idCol = 0;
      var nameCol = headers.indexOf('اسم الطالب'); if (nameCol < 0) nameCol = headers.indexOf('اسم_الطالب'); if (nameCol < 0) nameCol = 1;
      var gradeCol = headers.indexOf('الصف'); if (gradeCol < 0) gradeCol = headers.indexOf('رقم الصف'); if (gradeCol < 0) gradeCol = 2;
      var classCol = headers.indexOf('الفصل'); if (classCol < 0) classCol = 3;
      
      for (var i = 1; i < studentsData.length; i++) {
        var row = studentsData[i];
        var name = normalizeName_(row[nameCol]);
        if(!name) continue;
        
        masterMap[name] = {
          id: String(row[idCol] || '').trim(),
          name: row[nameCol],
          grade: cleanGradeName_(row[gradeCol]),
          class: row[classCol],
          stage: stg,
          excused: 0,
          unexcused: 0,
          late: 0,
          updated: new Date()
        };
      }
    }

    // الحفاظ على البيانات القديمة — ديناميكي حسب المراحل المفعّلة
    ensureStudentsSheetsLoaded_();
    var enabledStages = STUDENTS_SHEETS ? Object.keys(STUDENTS_SHEETS) : [];

    for (var s = 0; s < enabledStages.length; s++) {
      var sheetName = 'سجل_الغياب_' + enabledStages[s];
      var sheet = findSheet_(ss, sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        var data = sheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          var name = normalizeName_(data[i][1]);
          if (masterMap[name]) {
            masterMap[name].excused = Number(data[i][4]) || 0;
            masterMap[name].unexcused = Number(data[i][5]) || 0;
            masterMap[name].late = Number(data[i][6]) || 0;
          }
        }
      }
    }

    // قراءة ملف نور
    var rows = Utilities.parseCsv(fileContent);
    if (rows.length > 0 && rows[0].length < 5) rows = Utilities.parseCsv(fileContent, ';');
    
    var stats = { updated: 0 };
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var rawName = row[9] ? row[9].toString().trim() : ''; 
      if (!rawName || rawName.indexOf('الاســـم') !== -1) continue;

      var cleanName = normalizeName_(rawName);
      if (masterMap[cleanName]) {
        masterMap[cleanName].late = parseInt(row[0]) || 0;
        masterMap[cleanName].unexcused = parseInt(row[1]) || 0;
        masterMap[cleanName].excused = parseInt(row[2]) || 0;
        masterMap[cleanName].updated = new Date();
        stats.updated++;
      }
    }

    // الفرز والحفظ — ديناميكي حسب المراحل المفعّلة
    var listByStage = {};
    for (var es = 0; es < enabledStages.length; es++) {
      listByStage[enabledStages[es]] = {};
    }

    for (var key in masterMap) {
      var s = masterMap[key];
      var studentStage = s.stage || '';
      // البحث عن المرحلة المطابقة
      var matched = false;
      for (var es = 0; es < enabledStages.length; es++) {
        var stg = enabledStages[es];
        if (studentStage.indexOf(stg) !== -1 || (s.grade && s.grade.indexOf(stg) !== -1)) {
          listByStage[stg][key] = s;
          matched = true;
          break;
        }
      }
      // إذا لم تُطابق أي مرحلة، ضعها في أول مرحلة متاحة
      if (!matched && enabledStages.length > 0) {
        listByStage[enabledStages[0]][key] = s;
      }
    }

    for (var es = 0; es < enabledStages.length; es++) {
      saveAbsenceToSheet_(ss, 'سجل_الغياب_' + enabledStages[es], listByStage[enabledStages[es]]);
    }

    return { 
      success: true, 
      message: 'تمت المعالجة بنجاح:\n- تم تحديث بيانات: ' + stats.updated + ' طالب'
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// دالة مساعدة للحفظ
// =================================================================
function saveAbsenceToSheet_(ss, sheetName, dataMap) {
  var sheet = findSheet_(ss, sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.setRightToLeft(true);
  }
  
  sheet.clear();
  var headers = ['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'غياب بعذر', 'غياب بدون عذر', 'تأخير', 'آخر تحديث'];
  sheet.appendRow(headers);
  
  var rowsToWrite = [];
  for (var key in dataMap) {
    var student = dataMap[key];
    rowsToWrite.push([
      student.id,
      student.name,
      student.grade,
      student.class,
      student.excused,
      student.unexcused,
      student.late,
      student.updated
    ]);
  }
  
  if (rowsToWrite.length > 0) {
    sheet.getRange(2, 1, rowsToWrite.length, 8).setValues(rowsToWrite);
    sheet.getRange(1, 1, 1, 8)
      .setBackground('#f59e0b')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  }
}

// =================================================================
// جلب البيانات للواجهة
// =================================================================
function getAbsenceDashboardData(stage) {
  try {
    var sheet = getAbsenceSheet(stage);
    var allData = [];
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, records: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      allData.push({
        id: String(data[i][0] || ''),
        name: String(data[i][1] || ''),
        grade: String(data[i][2] || ''),
        class: String(data[i][3] || ''),
        excused: Number(data[i][4]) || 0,
        unexcused: Number(data[i][5]) || 0,
        late: Number(data[i][6]) || 0,
        lastUpdate: data[i][7] ? Utilities.formatDate(new Date(data[i][7]), Session.getScriptTimeZone(), 'yyyy/MM/dd') : ''
      });
    }
    
    return { success: true, records: allData };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// التعديل اليدوي
// =================================================================
function updateStudentAbsence(data) {
  try {
    if (!data || !data.id) throw new Error("بيانات غير مكتملة");
    
    var stage = data.stage;
    if (!stage) throw new Error("المرحلة الدراسية مطلوبة");
    var sheet = getAbsenceSheet(stage);
    
    var allData = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][0]) === String(data.id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("لم يتم العثور على الطالب");
    
    sheet.getRange(rowIndex, 5, 1, 4).setValues([[
      data.excused, 
      data.unexcused, 
      data.late, 
      new Date()
    ]]);
    
    return { success: true };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// جلب إحصائيات الغياب
// =================================================================
function getAbsenceStats(stage) {
  try {
    var result = getAbsenceDashboardData(stage);
    var records = result.records || [];
    
    var stats = {
      total: records.length,
      withExcused: 0,
      withUnexcused: 0,
      withLate: 0,
      totalExcused: 0,
      totalUnexcused: 0,
      totalLate: 0
    };
    
    for (var i = 0; i < records.length; i++) {
      var r = records[i];
      if (r.excused > 0) stats.withExcused++;
      if (r.unexcused > 0) stats.withUnexcused++;
      if (r.late > 0) stats.withLate++;
      stats.totalExcused += r.excused;
      stats.totalUnexcused += r.unexcused;
      stats.totalLate += r.late;
    }
    
    return { success: true, stats: stats };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// دالة تنظيف الأسماء
// =================================================================
function normalizeName_(name) {
  if (!name) return "";
  var n = String(name).trim();
  n = n.replace(/\s(بن|ابن)\s/g, ' '); 
  n = n.replace(/عبد\s+/g, 'عبد');
  n = n.replace(/[أإآ]/g, 'ا').replace(/ى/g, 'ي').replace(/ة/g, 'ه');
  n = n.replace(/\s+/g, ' ');
  return n;
}

// =================================================================
// دالة اختبار
// =================================================================
function TEST_Absence() {
  Logger.log("=== اختبار الغياب ===");
  var result = getAbsenceDashboardData('متوسط');
  Logger.log("النتيجة: " + JSON.stringify(result));
  Logger.log("عدد السجلات: " + (result.records ? result.records.length : 0));
}
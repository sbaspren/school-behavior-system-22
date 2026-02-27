// =================================================================
// Server_Academic.gs - نظام التحصيل الدراسي (الإصدار الشامل)
// شيتين: ملخص + درجات تفصيلية
// استيراد من شهادات نور (Excel)
// =================================================================

// ─── الثوابت ───
var ACADEMIC_SUMMARY_PREFIX = 'تحصيل_ملخص_';
var ACADEMIC_GRADES_PREFIX  = 'تحصيل_درجات_';

var SUMMARY_HEADERS = [
  'رقم_الهوية', 'اسم_الطالب', 'الصف', 'الفصل',
  'الفصل_الدراسي', 'الفترة',
  'المعدل', 'التقدير_العام',
  'ترتيب_الصف', 'ترتيب_الفصل',
  'الغياب', 'التأخر',
  'السلوك_متميز', 'السلوك_إيجابي'
];

var GRADES_HEADERS = [
  'رقم_الهوية', 'اسم_الطالب', 'الصف', 'الفصل',
  'الفصل_الدراسي', 'الفترة',
  'المادة', 'المجموع', 'اختبار_نهائي', 'أدوات_تقييم', 'اختبارات_قصيرة', 'التقدير'
];

// المواد غير الأكاديمية
var NON_ACADEMIC_SUBJECTS = ['السلوك', 'المواظبة', 'النشاط'];

// =================================================================
// 1. إنشاء / جلب الشيتات
// =================================================================
function getAcademicSummarySheet_(stage) {
  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var sheetName = ACADEMIC_SUMMARY_PREFIX + stage;
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.setRightToLeft(true);
    sheet.appendRow(SUMMARY_HEADERS);
    sheet.getRange(1, 1, 1, SUMMARY_HEADERS.length)
      .setBackground('#059669').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 110);
    sheet.setColumnWidth(2, 180);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 60);
    sheet.setColumnWidth(5, 80);
    sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 70);
    sheet.setColumnWidth(8, 80);
    sheet.setColumnWidth(9, 70);
    sheet.setColumnWidth(10, 70);
    sheet.setColumnWidth(11, 60);
    sheet.setColumnWidth(12, 60);
    sheet.setColumnWidth(13, 80);
    sheet.setColumnWidth(14, 80);
    
    var STAGE_TAB_COLORS = { 'متوسط': '#4285f4', 'ثانوي': '#0f9d58', 'ابتدائي': '#f4b400', 'طفولة مبكرة': '#e91e63' };
    sheet.setTabColor(STAGE_TAB_COLORS[stage] || '#9e9e9e');
  }
  return sheet;
}

function getAcademicGradesSheet_(stage) {
  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var sheetName = ACADEMIC_GRADES_PREFIX + stage;
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.setRightToLeft(true);
    sheet.appendRow(GRADES_HEADERS);
    sheet.getRange(1, 1, 1, GRADES_HEADERS.length)
      .setBackground('#0d9488').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 110);
    sheet.setColumnWidth(2, 180);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 60);
    sheet.setColumnWidth(5, 80);
    sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 180);
    sheet.setColumnWidth(8, 70);
    sheet.setColumnWidth(9, 80);
    sheet.setColumnWidth(10, 80);
    sheet.setColumnWidth(11, 80);
    sheet.setColumnWidth(12, 80);
    
    var STAGE_TAB_COLORS = { 'متوسط': '#3b82f6', 'ثانوي': '#10b981', 'ابتدائي': '#eab308', 'طفولة مبكرة': '#ec4899' };
    sheet.setTabColor(STAGE_TAB_COLORS[stage] || '#9e9e9e');
  }
  return sheet;
}

// =================================================================
// 2. استيراد شهادات نور (Excel)
// =================================================================
function importAcademicFromExcel(fileBlob, stage, semester, period) {
  try {
    var startTime = Date.now();
    Logger.log('بدء استيراد التحصيل: ' + stage + ' | ' + period);
    
    var tempFile = Drive.Files.insert(
      { title: 'temp_academic_import_' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS },
      fileBlob,
      { convert: true }
    );
    
    Logger.log('تحويل Excel اكتمل: ' + ((Date.now() - startTime) / 1000).toFixed(1) + ' ثانية');
    
    var tempSS = SpreadsheetApp.openById(tempFile.id);
    var allSheets = tempSS.getSheets();
    
    var summarySheet = getAcademicSummarySheet_(stage);
    var gradesSheet  = getAcademicGradesSheet_(stage);
    
    var importedCount = 0;
    var skippedCount = 0;
    var errors = [];
    var summaryRows = [];
    var gradesRows  = [];
    var detectedSemester = semester;
    var MAX_TIME = 300000; // 5 دقائق (حد Apps Script = 6)
    
    for (var s = 0; s < allSheets.length; s++) {
      // ── مراقبة الوقت ──
      if (Date.now() - startTime > MAX_TIME) {
        errors.push('تم إيقاف الاستيراد عند الشيت ' + (s+1) + ' بسبب حد الوقت. تم استيراد ' + importedCount + ' طالب.');
        break;
      }
      
      try {
        var ws = allSheets[s];
        
        // ── تخطي سريع: فحص إذا الشيت فيه بيانات كافية ──
        var lastRow = ws.getLastRow();
        var lastCol = ws.getLastColumn();
        if (lastRow < 30 || lastCol < 10) { skippedCount++; continue; }
        
        var studentData = parseStudentSheet_(ws, stage, semester, period);
        if (!studentData || !studentData.identity) { skippedCount++; continue; }
        
        if (!detectedSemester && studentData.detectedSemester) {
          detectedSemester = studentData.detectedSemester;
        }
        
        var useSemester = studentData.detectedSemester || detectedSemester || 'غير محدد';
        
        summaryRows.push([
          studentData.identity,
          studentData.name,
          studentData.grade,
          studentData.classNum,
          useSemester,
          period,
          studentData.average,
          studentData.generalGrade,
          studentData.rankGrade,
          studentData.rankClass,
          studentData.absence,
          studentData.tardiness,
          studentData.behaviorExcellent,
          studentData.behaviorPositive
        ]);
        
        for (var m = 0; m < studentData.subjects.length; m++) {
          var subj = studentData.subjects[m];
          gradesRows.push([
            studentData.identity,
            studentData.name,
            studentData.grade,
            studentData.classNum,
            useSemester,
            period,
            subj.name,
            subj.total,
            subj.finalExam,
            subj.evalTools,
            subj.shortTests,
            subj.grade
          ]);
        }
        
        importedCount++;
        
      } catch (sheetErr) {
        errors.push('شيت ' + (s + 1) + ': ' + sheetErr.toString());
      }
    }
    
    // ── حذف البيانات القديمة لنفس الفترة (بعد اكتشاف الفصل) ──
    if (detectedSemester) {
      deletePeriodData_(summarySheet, detectedSemester, period);
      deletePeriodData_(gradesSheet, detectedSemester, period);
    }
    
    // كتابة دفعة واحدة
    if (summaryRows.length > 0) {
      summarySheet.getRange(summarySheet.getLastRow() + 1, 1, summaryRows.length, SUMMARY_HEADERS.length)
        .setValues(summaryRows);
    }
    if (gradesRows.length > 0) {
      gradesSheet.getRange(gradesSheet.getLastRow() + 1, 1, gradesRows.length, GRADES_HEADERS.length)
        .setValues(gradesRows);
    }
    
    DriveApp.getFileById(tempFile.id).setTrashed(true);
    
    var totalTime = ((Date.now() - startTime) / 1000).toFixed(1);
    Logger.log('اكتمل الاستيراد: ' + importedCount + ' طالب في ' + totalTime + ' ثانية');
    
    return {
      success: true,
      imported: importedCount,
      totalSheets: allSheets.length,
      skipped: skippedCount,
      errors: errors,
      detectedSemester: detectedSemester,
      timeSeconds: totalTime,
      message: 'تم استيراد ' + importedCount + ' طالب في ' + totalTime + ' ثانية'
        + ' (' + (detectedSemester || 'غير محدد') + ' — ' + period + ')'
        + (errors.length > 0 ? ' | ' + errors.length + ' تنبيه' : '')
    };
    
  } catch (e) {
    Logger.log('خطأ في الاستيراد: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 3. تحليل شيت طالب واحد (بنية شهادة نور المتوسطة)
// =================================================================
function parseStudentSheet_(ws, stage, semester, period) {
  // ══ قراءة كل البيانات دفعة واحدة (طلب واحد فقط!) ══
  var lastRow = ws.getLastRow();
  var lastCol = ws.getLastColumn();
  if (lastRow < 30 || lastCol < 10) return null;
  
  var allData = ws.getRange(1, 1, Math.min(lastRow, 70), Math.min(lastCol, 50)).getValues();
  
  // دالة مساعدة: قراءة خلية من المصفوفة (صف 1-based, عمود 1-based)
  function cell(row, col) {
    if (row < 1 || row > allData.length || col < 1 || col > allData[0].length) return '';
    return String(allData[row - 1][col - 1] || '');
  }
  
  // ── البيانات الأساسية ──
  var nameAr = cell(28, 35); // AI28
  if (nameAr.indexOf('اسم الطالب:') >= 0) nameAr = nameAr.replace('اسم الطالب:', '').trim();
  if (!nameAr) return null;
  
  var classNum = cell(28, 27); // AA28
  var identity = cell(30, 19); // S30
  if (!identity || identity === '0' || identity === 'undefined') return null;
  
  // تحديد الصف
  var gradeText = cell(19, 10); // J19
  var gradeName = '';
  if (gradeText.indexOf('الأول') >= 0) gradeName = 'الأول المتوسط';
  else if (gradeText.indexOf('الثاني') >= 0) gradeName = 'الثاني المتوسط';
  else if (gradeText.indexOf('الثالث') >= 0) gradeName = 'الثالث المتوسط';
  else gradeName = gradeText;
  
  // ── استخراج الفصل الدراسي تلقائياً ──
  var detectedSemester = semester;
  if (!detectedSemester) {
    for (var r = 0; r < Math.min(allData.length, 25); r++) {
      for (var c = 0; c < allData[r].length; c++) {
        var v = String(allData[r][c] || '');
        if (v.indexOf('الفصل الدراسي الأول') >= 0 || v.indexOf('الفصل الأول') >= 0) {
          detectedSemester = 'الفصل الأول'; break;
        } else if (v.indexOf('الفصل الدراسي الثاني') >= 0 || v.indexOf('الفصل الثاني') >= 0) {
          detectedSemester = 'الفصل الثاني'; break;
        } else if (v.indexOf('الفصل الدراسي الثالث') >= 0 || v.indexOf('الفصل الثالث') >= 0) {
          detectedSemester = 'الفصل الثالث'; break;
        }
      }
      if (detectedSemester) break;
    }
    if (!detectedSemester) detectedSemester = 'غير محدد';
  }
  
  // ── المواد ──
  var subjects = [];
  var average = null, generalGrade = '';
  var maxRow = Math.min(allData.length, 60);
  
  for (var r = 35; r <= maxRow; r++) {
    var auVal = cell(r, 47).trim(); // AU column
    if (!auVal || auVal === 'المواد الدراسية' || auVal === 'مجموع الدرجات الموزونة') continue;
    
    if (auVal === 'المعدل') {
      var avgRaw = cell(r, 24); // X column
      average = parseFloat(avgRaw.replace('%', '').trim()) || null;
      continue;
    }
    if (auVal === 'التقدير العام') {
      generalGrade = cell(r, 37) || cell(r, 24); // AK or X
      continue;
    }
    
    subjects.push({
      name: auVal,
      total: toNum_(cell(r, 34)),     // AH
      finalExam: toNum_(cell(r, 38)),  // AL
      evalTools: toNum_(cell(r, 40)),  // AN
      shortTests: toNum_(cell(r, 45)), // AS
      grade: cell(r, 24)              // X
    });
  }
  
  // ── الترتيب والغياب والسلوك ──
  var rankGrade = '', rankClass = '', absence = '0', tardiness = '0';
  var behaviorExcellent = '', behaviorPositive = '';
  var scanEnd = Math.min(allData.length, 68);
  
  for (var r = 50; r <= scanEnd; r++) {
    var rowLen = allData[r - 1] ? allData[r - 1].length : 0;
    for (var c = 0; c < rowLen; c++) {
      var cv = String(allData[r - 1][c] || '');
      if (!cv) continue;
      if (cv.indexOf('الترتيب على الصف') >= 0 || cv.indexOf('Sort By Grade') >= 0) rankGrade = cell(r, 7);
      if (cv.indexOf('الترتيب على الفصل') >= 0 || cv.indexOf('Sort By Class') >= 0) rankClass = cell(r, 7);
      if (cv.indexOf('غياب بدون عذر') >= 0) absence = cell(r, 9) || '0';
      if (cv.indexOf('تأخر بدون عذر') >= 0) tardiness = cell(r, 9) || '0';
      if (cv.indexOf('درجة السلوك المتميز') >= 0) behaviorExcellent = cell(r, 44);
      if (cv.indexOf('درجة السلوك الإيجابي') >= 0) behaviorPositive = cell(r, 44);
    }
  }
  
  return {
    name: nameAr, identity: identity, grade: gradeName, classNum: classNum,
    average: average, generalGrade: generalGrade,
    rankGrade: rankGrade, rankClass: rankClass,
    absence: absence, tardiness: tardiness,
    behaviorExcellent: behaviorExcellent, behaviorPositive: behaviorPositive,
    subjects: subjects,
    detectedSemester: detectedSemester
  };
}

// =================================================================
// 4. حذف بيانات فترة محددة
// =================================================================
function deletePeriodData_(sheet, semester, period) {
  if (!sheet || sheet.getLastRow() < 2) return;
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var semCol = headers.indexOf('الفصل_الدراسي');
  var perCol = headers.indexOf('الفترة');
  
  if (semCol < 0 || perCol < 0) return;
  
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][semCol]) === String(semester) && String(data[i][perCol]) === String(period)) {
      sheet.deleteRow(i + 1);
    }
  }
}

// حذف بيانات فترة محددة لصفوف معينة فقط (لا يمس باقي الصفوف)
function deletePeriodDataByGrades_(sheet, semester, period, grades) {
  if (!sheet || sheet.getLastRow() < 2 || !grades || grades.length === 0) return;
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var semCol = headers.indexOf('الفصل_الدراسي');
  var perCol = headers.indexOf('الفترة');
  var gradeCol = headers.indexOf('الصف');
  
  if (semCol < 0 || perCol < 0 || gradeCol < 0) return;
  
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][semCol]) === String(semester) 
        && String(data[i][perCol]) === String(period)
        && grades.indexOf(String(data[i][gradeCol])) >= 0) {
      sheet.deleteRow(i + 1);
    }
  }
}

function deleteAcademicPeriod(stage, semester, period) {
  try {
    var summarySheet = getAcademicSummarySheet_(stage);
    var gradesSheet  = getAcademicGradesSheet_(stage);
    deletePeriodData_(summarySheet, semester, period);
    deletePeriodData_(gradesSheet, semester, period);
    return { success: true, message: 'تم حذف بيانات ' + period + ' - ' + semester };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 5. جلب البيانات
// =================================================================
function getAcademicSummary(stage) {
  try {
    var sheet = getAcademicSummarySheet_(stage);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, records: [] };
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var records = [];
    
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var record = {};
      for (var j = 0; j < headers.length; j++) {
        record[String(headers[j]).trim()] = data[i][j];
      }
      record.rowIndex = i;
      records.push(record);
    }
    
    return { success: true, records: records };
  } catch (e) {
    return { success: false, error: e.toString(), records: [] };
  }
}

function getAcademicGrades(stage) {
  try {
    var sheet = getAcademicGradesSheet_(stage);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, records: [] };
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var records = [];
    
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var record = {};
      for (var j = 0; j < headers.length; j++) {
        record[String(headers[j]).trim()] = data[i][j];
      }
      records.push(record);
    }
    
    return { success: true, records: records };
  } catch (e) {
    return { success: false, error: e.toString(), records: [] };
  }
}

function getAcademicAllData(stage) {
  try {
    var summary = getAcademicSummary(stage);
    var grades  = getAcademicGrades(stage);
    
    var periods = [];
    var periodSet = {};
    if (summary.success && summary.records) {
      for (var i = 0; i < summary.records.length; i++) {
        var key = summary.records[i]['الفصل_الدراسي'] + '|' + summary.records[i]['الفترة'];
        if (!periodSet[key]) {
          periodSet[key] = true;
          periods.push({
            semester: summary.records[i]['الفصل_الدراسي'],
            period: summary.records[i]['الفترة']
          });
        }
      }
    }
    
    return { success: true, summary: summary.records || [], grades: grades.records || [], periods: periods };
  } catch (e) {
    return { success: false, error: e.toString(), summary: [], grades: [], periods: [] };
  }
}

// =================================================================
// 6. تقرير طالب فردي شامل
// =================================================================
function getStudentAcademicReport(stage, identityNo) {
  try {
    var allData = getAcademicAllData(stage);
    if (!allData.success) return allData;
    
    var studentSummary = allData.summary.filter(function(r) {
      return String(r['رقم_الهوية']) === String(identityNo);
    });
    
    var studentGrades = allData.grades.filter(function(r) {
      return String(r['رقم_الهوية']) === String(identityNo);
    });
    
    if (studentSummary.length === 0) {
      return { success: false, error: 'لم يتم العثور على الطالب' };
    }
    
    // ── تحليل نقاط القوة والضعف ──
    var latestPeriod = studentSummary[studentSummary.length - 1];
    var latestGrades = studentGrades.filter(function(r) {
      return r['الفصل_الدراسي'] === latestPeriod['الفصل_الدراسي'] &&
             r['الفترة'] === latestPeriod['الفترة'];
    });
    
    var strengths = [], weaknesses = [], academicGrades = [];
    latestGrades.forEach(function(g) {
      if (NON_ACADEMIC_SUBJECTS.indexOf(g['المادة']) >= 0) return;
      var total = parseFloat(g['المجموع']) || 0;
      academicGrades.push({ name: g['المادة'], total: total, grade: g['التقدير'] });
      if (total >= 90) strengths.push(g['المادة']);
      else if (total < 65) weaknesses.push(g['المادة']);
    });
    
    academicGrades.sort(function(a, b) { return b.total - a.total; });
    
    // ── تحليل نوع الضعف ──
    var weaknessPattern = 'لا يوجد';
    if (weaknesses.length > 0) {
      var scienceWeak = weaknesses.filter(function(s) {
        return s === 'الرياضيات' || s === 'العلوم';
      }).length;
      if (scienceWeak === weaknesses.length) weaknessPattern = 'ضعف علمي';
      else if (weaknesses.length >= 4) weaknessPattern = 'ضعف شامل';
      else weaknessPattern = 'ضعف جزئي';
    }
    
    // ── تحليل النهائي vs أعمال السنة ──
    var examVsWork = [];
    latestGrades.forEach(function(g) {
      if (NON_ACADEMIC_SUBJECTS.indexOf(g['المادة']) >= 0) return;
      var final_ = parseFloat(g['اختبار_نهائي']) || 0;
      var tools  = parseFloat(g['أدوات_تقييم']) || 0;
      var short_ = parseFloat(g['اختبارات_قصيرة']) || 0;
      var work   = tools + short_;
      if (final_ > 0 || work > 0) {
        examVsWork.push({ name: g['المادة'], finalExam: final_, classWork: work });
      }
    });
    
    return {
      success: true,
      student: {
        name: studentSummary[0]['اسم_الطالب'],
        identity: identityNo,
        grade: studentSummary[0]['الصف'],
        classNum: studentSummary[0]['الفصل']
      },
      summary: studentSummary,
      grades: studentGrades,
      analysis: {
        strengths: strengths,
        weaknesses: weaknesses,
        weaknessPattern: weaknessPattern,
        academicGrades: academicGrades,
        examVsWork: examVsWork,
        absence: parseInt(latestPeriod['الغياب']) || 0,
        tardiness: parseInt(latestPeriod['التأخر']) || 0,
        behaviorExcellent: latestPeriod['السلوك_متميز'],
        behaviorPositive: latestPeriod['السلوك_إيجابي']
      }
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 7. إحصائيات سريعة (لوحة المؤشرات)
// =================================================================
function getAcademicQuickStats(stage, semester, period) {
  try {
    var allData = getAcademicAllData(stage);
    if (!allData.success) return allData;
    
    var summary = allData.summary.filter(function(r) {
      if (semester && String(r['الفصل_الدراسي']) !== String(semester)) return false;
      if (period && String(r['الفترة']) !== String(period)) return false;
      return true;
    });
    
    var grades = allData.grades.filter(function(r) {
      if (semester && String(r['الفصل_الدراسي']) !== String(semester)) return false;
      if (period && String(r['الفترة']) !== String(period)) return false;
      return true;
    });
    
    var totalStudents = summary.length;
    var avgs = [];
    summary.forEach(function(r) {
      var a = parseFloat(r['المعدل']);
      if (!isNaN(a) && a > 0) avgs.push(a);
    });
    
    // توزيع التقديرات
    var gradeDist = {};
    summary.forEach(function(r) {
      var g = String(r['التقدير_العام'] || 'غير محدد');
      gradeDist[g] = (gradeDist[g] || 0) + 1;
    });
    
    // تصنيف الطلاب
    var categories = { excellent: 0, good: 0, average: 0, weak: 0, danger: 0 };
    avgs.forEach(function(a) {
      if (a >= 95) categories.excellent++;
      else if (a >= 80) categories.good++;
      else if (a >= 65) categories.average++;
      else if (a >= 50) categories.weak++;
      else categories.danger++;
    });
    
    // إحصائيات المواد
    var subjectStats = {};
    grades.forEach(function(r) {
      var subj = String(r['المادة'] || '');
      if (NON_ACADEMIC_SUBJECTS.indexOf(subj) >= 0) return;
      if (!subjectStats[subj]) subjectStats[subj] = { totals: [], name: subj };
      var t = parseFloat(r['المجموع']);
      if (!isNaN(t)) subjectStats[subj].totals.push(t);
    });
    
    var subjectSummary = [];
    for (var s in subjectStats) {
      var arr = subjectStats[s].totals;
      if (arr.length === 0) continue;
      var sum = arr.reduce(function(a, b) { return a + b; }, 0);
      subjectSummary.push({
        name: s,
        avg: Math.round(sum / arr.length * 10) / 10,
        max: Math.max.apply(null, arr),
        min: Math.min.apply(null, arr),
        count: arr.length,
        above90: arr.filter(function(v) { return v >= 90; }).length,
        below60: arr.filter(function(v) { return v < 60; }).length,
        below50: arr.filter(function(v) { return v < 50; }).length
      });
    }
    subjectSummary.sort(function(a, b) { return a.avg - b.avg; });
    
    // العشرة الأوائل
    var topTen = summary.slice().sort(function(a, b) {
      return (parseFloat(b['المعدل']) || 0) - (parseFloat(a['المعدل']) || 0);
    }).slice(0, 10);
    
    // أقل 10 طلاب
    var bottomTen = summary.slice().sort(function(a, b) {
      return (parseFloat(a['المعدل']) || 0) - (parseFloat(b['المعدل']) || 0);
    }).slice(0, 10);
    
    // إحصائيات الغياب
    var totalAbsence = 0, totalTardiness = 0, absenceStudents = 0;
    summary.forEach(function(r) {
      var ab = parseInt(r['الغياب']) || 0;
      var td = parseInt(r['التأخر']) || 0;
      totalAbsence += ab;
      totalTardiness += td;
      if (ab > 0) absenceStudents++;
    });
    
    // إحصائيات كل فصل
    var classStats = {};
    summary.forEach(function(r) {
      var key = r['الصف'] + ' - ' + r['الفصل'];
      if (!classStats[key]) classStats[key] = { avgs: [], count: 0, grade: r['الصف'], classNum: r['الفصل'] };
      classStats[key].count++;
      var a = parseFloat(r['المعدل']);
      if (!isNaN(a) && a > 0) classStats[key].avgs.push(a);
    });
    
    var classSummary = [];
    for (var k in classStats) {
      var cs = classStats[k];
      var csAvgs = cs.avgs;
      classSummary.push({
        label: k,
        grade: cs.grade,
        classNum: cs.classNum,
        count: cs.count,
        avg: csAvgs.length > 0 ? Math.round(csAvgs.reduce(function(a,b){return a+b;}, 0) / csAvgs.length * 100) / 100 : 0,
        max: csAvgs.length > 0 ? Math.max.apply(null, csAvgs) : 0,
        min: csAvgs.length > 0 ? Math.min.apply(null, csAvgs) : 0,
        excellent: csAvgs.filter(function(v) { return v >= 95; }).length,
        weak: csAvgs.filter(function(v) { return v < 65; }).length
      });
    }
    
    // طلاب منطقة الخطر (3+ مواد أقل من 60)
    var dangerStudents = [];
    var studentGradesMap = {};
    grades.forEach(function(r) {
      if (NON_ACADEMIC_SUBJECTS.indexOf(r['المادة']) >= 0) return;
      var id = String(r['رقم_الهوية']);
      if (!studentGradesMap[id]) studentGradesMap[id] = { name: r['اسم_الطالب'], weakSubjects: [] };
      var t = parseFloat(r['المجموع']);
      if (!isNaN(t) && t < 60) studentGradesMap[id].weakSubjects.push(r['المادة']);
    });
    for (var id in studentGradesMap) {
      if (studentGradesMap[id].weakSubjects.length >= 3) {
        dangerStudents.push({
          identity: id,
          name: studentGradesMap[id].name,
          weakSubjects: studentGradesMap[id].weakSubjects,
          weakCount: studentGradesMap[id].weakSubjects.length
        });
      }
    }
    
    return {
      success: true,
      stats: {
        totalStudents: totalStudents,
        avgAll: avgs.length > 0 ? Math.round(avgs.reduce(function(a,b){return a+b;}, 0) / avgs.length * 100) / 100 : 0,
        maxAvg: avgs.length > 0 ? Math.round(Math.max.apply(null, avgs) * 100) / 100 : 0,
        minAvg: avgs.length > 0 ? Math.round(Math.min.apply(null, avgs) * 100) / 100 : 0,
        gradeDist: gradeDist,
        categories: categories,
        subjects: subjectSummary,
        topTen: topTen,
        bottomTen: bottomTen,
        classSummary: classSummary,
        dangerStudents: dangerStudents,
        absence: { total: totalAbsence, tardiness: totalTardiness, studentsWithAbsence: absenceStudents },
        periods: allData.periods
      }
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 8. مقارنة الفصول
// =================================================================
function getClassComparison(stage, semester, period) {
  try {
    var allData = getAcademicAllData(stage);
    if (!allData.success) return allData;
    
    var grades = allData.grades.filter(function(r) {
      if (semester && String(r['الفصل_الدراسي']) !== String(semester)) return false;
      if (period && String(r['الفترة']) !== String(period)) return false;
      return true;
    });
    
    // بناء مقارنة: لكل مادة → لكل فصل → المتوسط
    var comparison = {};
    grades.forEach(function(r) {
      if (NON_ACADEMIC_SUBJECTS.indexOf(r['المادة']) >= 0) return;
      var subj = r['المادة'];
      var classKey = r['الصف'] + ' فصل ' + r['الفصل'];
      
      if (!comparison[subj]) comparison[subj] = {};
      if (!comparison[subj][classKey]) comparison[subj][classKey] = [];
      
      var t = parseFloat(r['المجموع']);
      if (!isNaN(t)) comparison[subj][classKey].push(t);
    });
    
    var result = [];
    for (var subj in comparison) {
      var classes = [];
      for (var cls in comparison[subj]) {
        var arr = comparison[subj][cls];
        var sum = arr.reduce(function(a, b) { return a + b; }, 0);
        classes.push({
          classLabel: cls,
          avg: Math.round(sum / arr.length * 10) / 10,
          count: arr.length,
          above90: arr.filter(function(v) { return v >= 90; }).length,
          below60: arr.filter(function(v) { return v < 60; }).length
        });
      }
      result.push({ subject: subj, classes: classes });
    }
    
    return { success: true, comparison: result };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 9. مقارنة الفترات (تطور طالب)
// =================================================================
function getStudentProgress(stage, identityNo) {
  try {
    var allData = getAcademicAllData(stage);
    if (!allData.success) return allData;
    
    var studentGrades = allData.grades.filter(function(r) {
      return String(r['رقم_الهوية']) === String(identityNo);
    });
    
    var studentSummary = allData.summary.filter(function(r) {
      return String(r['رقم_الهوية']) === String(identityNo);
    });
    
    if (studentSummary.length === 0) {
      return { success: false, error: 'لم يتم العثور على الطالب' };
    }
    
    // ترتيب حسب الفترة
    var periodOrder = ['الفترة الأولى', 'الفترة الثانية', 'نهاية الفصل'];
    
    // تطور المعدل
    var avgProgress = studentSummary.map(function(r) {
      return {
        semester: r['الفصل_الدراسي'],
        period: r['الفترة'],
        average: parseFloat(r['المعدل']) || 0,
        generalGrade: r['التقدير_العام']
      };
    });
    
    // تطور كل مادة
    var subjectProgress = {};
    studentGrades.forEach(function(r) {
      if (NON_ACADEMIC_SUBJECTS.indexOf(r['المادة']) >= 0) return;
      var subj = r['المادة'];
      if (!subjectProgress[subj]) subjectProgress[subj] = [];
      subjectProgress[subj].push({
        semester: r['الفصل_الدراسي'],
        period: r['الفترة'],
        total: parseFloat(r['المجموع']) || 0
      });
    });
    
    return {
      success: true,
      student: { name: studentSummary[0]['اسم_الطالب'], identity: identityNo },
      avgProgress: avgProgress,
      subjectProgress: subjectProgress
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 10. فلترة وبحث
// =================================================================
function searchAcademicStudents(stage, filters) {
  try {
    var allData = getAcademicAllData(stage);
    if (!allData.success) return allData;
    
    var results = allData.summary;
    
    // فلتر الفترة
    if (filters.semester) {
      results = results.filter(function(r) { return String(r['الفصل_الدراسي']) === String(filters.semester); });
    }
    if (filters.period) {
      results = results.filter(function(r) { return String(r['الفترة']) === String(filters.period); });
    }
    
    // فلتر الصف
    if (filters.grade) {
      results = results.filter(function(r) { return String(r['الصف']).indexOf(filters.grade) >= 0; });
    }
    
    // فلتر الفصل
    if (filters.classNum) {
      results = results.filter(function(r) { return String(r['الفصل']) === String(filters.classNum); });
    }
    
    // فلتر التقدير
    if (filters.generalGrade) {
      results = results.filter(function(r) { return String(r['التقدير_العام']) === filters.generalGrade; });
    }
    
    // فلتر المعدل (أقل من / أكثر من)
    if (filters.avgBelow) {
      var below = parseFloat(filters.avgBelow);
      results = results.filter(function(r) { return (parseFloat(r['المعدل']) || 0) < below; });
    }
    if (filters.avgAbove) {
      var above = parseFloat(filters.avgAbove);
      results = results.filter(function(r) { return (parseFloat(r['المعدل']) || 0) >= above; });
    }
    
    // بحث بالاسم
    if (filters.name) {
      var searchName = filters.name;
      results = results.filter(function(r) { return String(r['اسم_الطالب']).indexOf(searchName) >= 0; });
    }
    
    // الترتيب
    if (filters.sortBy === 'avg_asc') {
      results.sort(function(a, b) { return (parseFloat(a['المعدل']) || 0) - (parseFloat(b['المعدل']) || 0); });
    } else if (filters.sortBy === 'avg_desc') {
      results.sort(function(a, b) { return (parseFloat(b['المعدل']) || 0) - (parseFloat(a['المعدل']) || 0); });
    } else if (filters.sortBy === 'name') {
      results.sort(function(a, b) { return String(a['اسم_الطالب']).localeCompare(String(b['اسم_الطالب']), 'ar'); });
    }
    
    return { success: true, records: results, total: results.length };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 11. الفترات المتوفرة
// =================================================================
function getAcademicPeriods(stage) {
  try {
    var sheet = getAcademicSummarySheet_(stage);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, periods: [] };
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var semCol = headers.indexOf('الفصل_الدراسي');
    var perCol = headers.indexOf('الفترة');
    
    var periodMap = {};
    for (var i = 1; i < data.length; i++) {
      var key = String(data[i][semCol]) + '|' + String(data[i][perCol]);
      if (!periodMap[key]) periodMap[key] = { semester: String(data[i][semCol]), period: String(data[i][perCol]), count: 0 };
      periodMap[key].count++;
    }
    
    var periodsArr = [];
    for (var k in periodMap) periodsArr.push(periodMap[k]);
    return { success: true, periods: periodsArr };
  } catch (e) {
    return { success: false, periods: [], error: e.toString() };
  }
}

// =================================================================
// 12. دوال مساعدة
// =================================================================
function toNum_(val) {
  if (val === null || val === undefined || val === '' || val === ' ') return 0;
  var n = parseFloat(val);
  return isNaN(n) ? 0 : n;
}

// =================================================================
// 13. استيراد من Base64 (يستقبل من الواجهة)
// =================================================================
function importAcademicFromExcelBase64(base64Data, fileName, stage, semester, period) {
  try {
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileName);
    return importAcademicFromExcel(blob, stage, semester, period);
  } catch (e) {
    Logger.log('خطأ Base64: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 14. حفظ بيانات محللة من المتصفح (SheetJS) — سريع جداً
// =================================================================
function saveAcademicParsedData(jsonStr, stage, period) {
  try {
    var startTime = Date.now();
    var students = JSON.parse(jsonStr);
    
    if (!students || students.length === 0) {
      return { success: false, error: 'لا توجد بيانات طلاب' };
    }
    
    var summarySheet = getAcademicSummarySheet_(stage);
    var gradesSheet  = getAcademicGradesSheet_(stage);
    
    // استخراج الفصل الدراسي والصفوف الموجودة في الملف
    var detectedSemester = '';
    var gradesInFile = [];
    for (var i = 0; i < students.length; i++) {
      if (!detectedSemester && students[i].semester) detectedSemester = students[i].semester;
      if (students[i].grade && gradesInFile.indexOf(students[i].grade) < 0) {
        gradesInFile.push(students[i].grade);
      }
    }
    
    // حذف البيانات القديمة لنفس الفترة والصفوف فقط (لا يمس باقي الصفوف)
    if (detectedSemester && gradesInFile.length > 0) {
      deletePeriodDataByGrades_(summarySheet, detectedSemester, period, gradesInFile);
      deletePeriodDataByGrades_(gradesSheet, detectedSemester, period, gradesInFile);
    }
    
    // بناء صفوف الملخص والدرجات
    var summaryRows = [];
    var gradesRows  = [];
    
    for (var i = 0; i < students.length; i++) {
      var s = students[i];
      var sem = s.semester || detectedSemester || 'غير محدد';
      
      summaryRows.push([
        s.identity,
        s.name,
        s.grade,
        s.classNum,
        sem,
        period,
        s.average,
        s.generalGrade,
        s.rankGrade,
        s.rankClass,
        s.absence,
        s.tardiness,
        s.behaviorExcellent,
        s.behaviorPositive
      ]);
      
      if (s.subjects) {
        for (var j = 0; j < s.subjects.length; j++) {
          var subj = s.subjects[j];
          gradesRows.push([
            s.identity,
            s.name,
            s.grade,
            s.classNum,
            sem,
            period,
            subj.name,
            subj.total,
            subj.finalExam,
            subj.evalTools,
            subj.shortTests,
            subj.grade
          ]);
        }
      }
    }
    
    // كتابة دفعة واحدة
    if (summaryRows.length > 0) {
      summarySheet.getRange(summarySheet.getLastRow() + 1, 1, summaryRows.length, SUMMARY_HEADERS.length)
        .setValues(summaryRows);
    }
    if (gradesRows.length > 0) {
      gradesSheet.getRange(gradesSheet.getLastRow() + 1, 1, gradesRows.length, GRADES_HEADERS.length)
        .setValues(gradesRows);
    }
    
    var totalTime = ((Date.now() - startTime) / 1000).toFixed(1);
    Logger.log('حفظ ' + students.length + ' طالب في ' + totalTime + ' ثانية (سيرفر فقط)');
    
    return {
      success: true,
      imported: students.length,
      semester: detectedSemester,
      grades: gradesInFile,
      serverTime: totalTime,
      message: 'تم حفظ ' + students.length + ' طالب بنجاح'
        + ' | الصفوف: ' + gradesInFile.join('، ')
    };
    
  } catch (e) {
    Logger.log('خطأ في حفظ البيانات المحللة: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}
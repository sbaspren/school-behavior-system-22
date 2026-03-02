// =================================================================
// VIOLATION LOGIC - منطق المخالفات (النسخة النهائية المتوافقة)
// =================================================================

function calculateRepeatLevel(studentId, violationId) {
  try {
    const students = getStudents_();
    const student = students.find(s => s['رقم الطالب'] == studentId);
    if (!student) throw new Error("Student not found.");
    
    const logSheetName = getSheetName_('المخالفات', student['المرحلة']);
    const ss = getSpreadsheet_();
    const sheet = findSheet_(ss, logSheetName);

    if (!sheet || sheet.getLastRow() < 2) return { success: true, repeatLevel: 1, previousProcedures: [] };

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    
    // البحث الديناميكي عن الأعمدة
    const studentIdColIndex = headers.indexOf('رقم الطالب');
    const violationIdColIndex = headers.indexOf('رقم المخالفة');
    const proceduresColIndex = headers.indexOf('الإجراءات');

    const previousViolations = data.filter(row => row[studentIdColIndex] == studentId && row[violationIdColIndex] == violationId);
    
    let previousProcedures = [];
    if (previousViolations.length > 0) {
      const lastViolation = previousViolations[previousViolations.length - 1];
      previousProcedures = lastViolation[proceduresColIndex] ? lastViolation[proceduresColIndex].split('\n') : [];
    }

    return { success: true, repeatLevel: previousViolations.length + 1, previousProcedures };
  } catch (e) {
    console.log("Error in calculateRepeatLevel: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function getCachedViolationRecords(stage) {
  var cacheKey = 'violations_' + stage + '_' + new Date().toLocaleDateString('en-US');
  var cache = CacheService.getScriptCache();
  try {
    var cached = cache.get(cacheKey);
    if (cached != null) return JSON.parse(cached);
  } catch (e) {
    // تجاهل خطأ قراءة الكاش
  }

  var records = getViolationRecords(stage);
  // ★ حماية من تجاوز حد CacheService (100KB لكل مفتاح)
  try {
    var json = JSON.stringify(records);
    if (json.length < 90000) { // ~90KB أمان — الحد 100KB
      cache.put(cacheKey, json, 300);
    }
  } catch (e) {
    // تجاهل خطأ الكاش — البيانات تُرجع بدونه
  }
  return records;
}

function getViolationRecords(stage) {
  try {
    const logSheetName = getSheetName_('المخالفات', stage);
    const ss = getSpreadsheet_();
    const sheet = findSheet_(ss, logSheetName);

    if (!sheet || sheet.getLastRow() < 2) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    
    return data.map((row, i) => {
      let record = {};
      headers.forEach((header, index) => {
        var val = row[index];
        if (val && val instanceof Date) {
          // ★ التاريخ الهجري يُحول لنص هجري وليس ISO
          if (header === 'التاريخ الهجري') {
            record[header] = readHijriCellValue_(val);
          } else {
            record[header] = val.toISOString();
          }
        } else {
          record[header] = (val !== null && val !== undefined && val !== '') ? val : '';
        }
      });
      // ★ رقم صف الشيت (للتحديث لاحقاً، مثل حالة تم الإرسال)
      record.rowIndex = i + 2;
      return record;
    }).filter(record => record['رقم الطالب']);

  } catch (e) {
    console.error("❌ Error fetching records:", e.toString());
    return []; 
  }
}

/**
 * ★ جلب مخالفات اليوم فقط (أسرع وأخف من جلب كل السجلات)
 * يُستخدم في تبويب "اليومي" — يفلتر من السيرفر بدلاً من إرسال كل البيانات للعميل
 * @param {string} stage - المرحلة
 * @return {Object} { today: [...], allCount: N, criticalCount: N }
 */
function getTodayViolationRecords(stage) {
  try {
    var logSheetName = getSheetName_('المخالفات', stage);
    var ss = getSpreadsheet_();
    var sheet = findSheet_(ss, logSheetName);

    if (!sheet || sheet.getLastRow() < 2) return { today: [], allCount: 0, criticalCount: 0 };

    var data = sheet.getDataRange().getValues();
    var headers = data.shift();

    var tz = Session.getScriptTimeZone();
    var todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    // بناء خريطة الأعمدة
    var colMap = {};
    headers.forEach(function(h, i) { colMap[String(h).trim()] = i; });
    var dateColIdx = colMap['التاريخ الميلادي'];
    var degreeColIdx = colMap['الدرجة'];
    var studentColIdx = colMap['رقم الطالب'];

    var todayRecords = [];
    var allCount = 0;
    var criticalCount = 0;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[studentColIdx]) continue;
      allCount++;
      var deg = parseFloat(row[degreeColIdx] || 0);
      if (deg >= 4) criticalCount++;

      // ★ فلتر اليوم: مقارنة التاريخ الميلادي بتاريخ اليوم
      var cellDate = (dateColIdx !== undefined) ? row[dateColIdx] : null;
      if (!cellDate) continue;
      var isToday = false;
      if (cellDate instanceof Date) {
        isToday = Utilities.formatDate(cellDate, tz, 'yyyy-MM-dd') === todayStr;
      } else {
        // نص تاريخ أو serial
        try {
          var dt = new Date(cellDate);
          if (!isNaN(dt.getTime())) {
            isToday = Utilities.formatDate(dt, tz, 'yyyy-MM-dd') === todayStr;
          }
        } catch(e) {}
      }

      if (isToday) {
        var record = {};
        headers.forEach(function(header, idx) {
          var val = row[idx];
          if (val && val instanceof Date) {
            // ★ التاريخ الهجري يُحول لنص هجري وليس ISO
            if (header === 'التاريخ الهجري') {
              record[header] = readHijriCellValue_(val);
            } else {
              record[header] = val.toISOString();
            }
          } else {
            record[header] = (val !== null && val !== undefined && val !== '') ? val : '';
          }
        });
        record.rowIndex = i + 2;
        todayRecords.push(record);
      }
    }

    return { today: todayRecords, allCount: allCount, criticalCount: criticalCount };
  } catch (e) {
    console.error('❌ getTodayViolationRecords:', e.toString());
    return { today: [], allCount: 0, criticalCount: 0 };
  }
}

/**
 * تحديث حالة "تم الإرسال" لسجلات المخالفات بعد إرسال واتساب
 * @param {string} stage - المرحلة (متوسط / ثانوي)
 * @param {number[]} rowIndices - أرقام صفوف الشيت (1-based، الصف 1 = الترويسة)
 */
function updateViolationSentStatus(stage, rowIndices) {
  try {
    if (!rowIndices || rowIndices.length === 0) return;
    var logSheetName = getSheetName_('المخالفات', stage);
    var ss = getSpreadsheet_();
    var sheet = findSheet_(ss, logSheetName);
    if (!sheet) return;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var sentCol = headers.indexOf('تم الإرسال');
    if (sentCol === -1) return;
    var lastRow = sheet.getLastRow();
    for (var i = 0; i < rowIndices.length; i++) {
      var row = parseInt(rowIndices[i]);
      if (isNaN(row) || row < 2 || row > lastRow) continue;
      sheet.getRange(row, sentCol + 1).setValue('نعم');
    }
    var cacheKey = 'violations_' + stage + '_' + new Date().toLocaleDateString('en-US');
    CacheService.getScriptCache().remove(cacheKey);
  } catch (e) {
    console.error('updateViolationSentStatus: ' + e.toString());
  }
}

// =================================================================
// حذف سجل مخالفة
// =================================================================
function deleteViolationRecord(stage, rowIndex) {
  try {
    var logSheetName = getSheetName_('المخالفات', stage);
    var ss = getSpreadsheet_();
    var sheet = findSheet_(ss, logSheetName);
    if (!sheet) return { success: false, error: 'الشيت غير موجود' };

    var ri = parseInt(rowIndex);
    if (isNaN(ri) || ri < 2 || ri > sheet.getLastRow()) {
      return { success: false, error: 'رقم الصف غير صالح' };
    }

    sheet.deleteRow(ri);
    return { success: true, message: 'تم الحذف' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// SAVING DATA - حفظ المخالفة (هيكل موحد 18 عمود)
// =================================================================
function saveViolation(data) {
  console.log("🔍 بدء حفظ المخالفة:", data);
  
  try {
    if (!data || !data.studentId || !data.violationId) throw new Error("بيانات غير مكتملة");
    
    const students = getStudents_();
    const rules = getRulesData_();
    const violations = rules.violations;
    
    // 1. استدعاء بيانات الطالب (الموثوقة)
    const student = students.find(s => s['رقم الطالب'] == data.studentId);
    if (!student) throw new Error("الطالب غير موجود: " + data.studentId);
    
    // 2. استدعاء بيانات المخالفة
    const violation = violations.find(v => v.id == data.violationId);
    if (!violation) throw new Error("المخالفة غير موجودة: " + data.violationId);
    
    // 3. تحديد الشيت (بحث ذكي + إنشاء تلقائي)
    const logSheetName = getSheetName_('المخالفات', student['المرحلة']);
    const ss = getSpreadsheet_();
    var sheet = findSheet_(ss, logSheetName);
    
    // ★ الترويسة الموحدة 18 عمود (مطابقة Config.gs و Server_TeacherInput.gs)
    var violationHeaders = [
      'رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل',
      'رقم المخالفة', 'نص المخالفة', 'نوع المخالفة', 'الدرجة',
      'التاريخ الهجري', 'التاريخ الميلادي', 'مستوى التكرار', 'الإجراءات',
      'النقاط', 'اليوم', 'النماذج المحفوظة', 'المستخدم', 'وقت الإدخال', 'تم الإرسال'
    ];
    
    // إنشاء الشيت بالعناوين الصحيحة إذا لم يكن موجوداً
    if (!sheet) {
        sheet = ss.insertSheet(logSheetName);
        sheet.setRightToLeft(true);
        sheet.appendRow(violationHeaders);
        sheet.getRange(1, 1, 1, violationHeaders.length)
          .setBackground('#e74c3c')
          .setFontColor('#ffffff')
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
        sheet.setFrozenRows(1);
        // ★ تنسيق عمود التاريخ الهجري (9) كنص عادي لمنع التحويل التلقائي
        sheet.getRange(1, 9, sheet.getMaxRows(), 1).setNumberFormat('@');

        // لون التبويب من السجل المركزي
        var regEntry = SHEET_REGISTRY['المخالفات'];
        if (regEntry && regEntry.color) sheet.setTabColor(regEntry.color);
    } else if(sheet.getLastRow() < 1) {
        // إذا كان الشيت موجوداً ولكنه فارغ
        sheet.appendRow(violationHeaders);
    } else {
        // ★ ترقية الشيت القديم — إضافة العمود الناقص فقط (لا نستبدل الهيدرات كاملة لتجنب عدم تطابق البيانات)
        var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        if (currentHeaders.indexOf('تم الإرسال') === -1) {
            var newCol = sheet.getLastColumn() + 1;
            sheet.getRange(1, newCol).setValue('تم الإرسال').setBackground('#e74c3c').setFontColor('#ffffff').setFontWeight('bold');
        }
    }
    
    // ★ حساب النقاط تلقائياً من الدرجة ومستوى التكرار (إن لم يُمرَّر من الواجهة)
    var points = data.points;
    if (points === undefined || points === null || points === '') {
      points = getDeductionForViolation_(String(data.violationId), data.repeatLevel || 1, student['المرحلة']);
    }
    
    var now = new Date();
    var dayName = now.toLocaleDateString('ar-SA', { weekday: 'long' });
    var violTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    
    // 4. بناء الصف الجديد (18 عمود)
    const newRowData = [
      student['رقم الطالب'],       // 1
      student['اسم الطالب'],       // 2
      student['الصف'],            // 3
      student['الفصل'],           // 4
      violation.id,               // 5
      violation.text,             // 6
      violation.type,             // 7
      getEffectiveDegree_(violation.id, student['المرحلة']),  // 8: الدرجة الفعلية حسب المرحلة
      getHijriDate_(now),             // 9: التاريخ الهجري (★ موحد مع المعلم)
      now,                        // 10
      data.repeatLevel || 1,      // 11
      Array.isArray(data.procedures) ? data.procedures.map(function(p){ return sanitizeInput_(p); }).join('\n') : '', // 12
      points,                     // 13: النقاط (محسوبة أو مُمرَّرة)
      dayName,                    // 14: اليوم
      Array.isArray(data.forms) ? data.forms.map(function(f){ return sanitizeInput_(f); }).join('\n') : '', // 15
      'الوكيل',                   // 16
      violTime,                   // 17
      'لا'                        // 18: تم الإرسال
    ];
    
    // الحفظ في نفس الشيت
    sheet.appendRow(newRowData);
        
    // مسح الكاش
    const cacheKey = `violations_${student['المرحلة']}_${new Date().toLocaleDateString('en-US')}`;
    CacheService.getScriptCache().remove(cacheKey);
    
    return { 
      success: true, 
      message: "تم حفظ المخالفة بنجاح!",
      studentName: student['اسم الطالب'],
      proceduresCount: Array.isArray(data.procedures) ? data.procedures.length : 0,
      violationText: violation.text
    };

  } catch (e) {
    console.error("❌ خطأ في حفظ المخالفة:", e.toString());
    return { success: false, error: e.message };
  }
}

// =================================================================
// BATCH SAVE - حفظ مخالفة لعدة طلاب دفعة واحدة
// =================================================================
function saveViolationsBatch(data) {
  try {
    if (!data || !data.students || data.students.length === 0 || !data.violationId)
      throw new Error("بيانات غير مكتملة");

    var allStudents = getStudents_();
    var rules = getRulesData_();
    var violation = rules.violations.find(function(v) { return v.id == data.violationId; });
    if (!violation) throw new Error("المخالفة غير موجودة: " + data.violationId);

    var stage = data.stage;
    var logSheetName = getSheetName_('المخالفات', stage);
    var ss = getSpreadsheet_();
    var sheet = findSheet_(ss, logSheetName);

    // ★ الترويسة الموحدة 18 عمود
    var violationHeaders = [
      'رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل',
      'رقم المخالفة', 'نص المخالفة', 'نوع المخالفة', 'الدرجة',
      'التاريخ الهجري', 'التاريخ الميلادي', 'مستوى التكرار', 'الإجراءات',
      'النقاط', 'اليوم', 'النماذج المحفوظة', 'المستخدم', 'وقت الإدخال', 'تم الإرسال'
    ];

    // إنشاء الشيت إذا لم يوجد (نفس منطق saveViolation)
    if (!sheet) {
      sheet = ss.insertSheet(logSheetName);
      sheet.setRightToLeft(true);
      sheet.appendRow(violationHeaders);
      sheet.getRange(1, 1, 1, violationHeaders.length)
        .setBackground('#e74c3c').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
      sheet.getRange(1, 9, sheet.getMaxRows(), 1).setNumberFormat('@');
      var regEntry = SHEET_REGISTRY['المخالفات'];
      if (regEntry && regEntry.color) sheet.setTabColor(regEntry.color);
    } else if (sheet.getLastRow() < 1) {
      sheet.appendRow(violationHeaders);
    } else {
      var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (currentHeaders.indexOf('تم الإرسال') === -1) {
        var newCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCol).setValue('تم الإرسال').setBackground('#e74c3c').setFontColor('#ffffff').setFontWeight('bold');
      }
    }

    // ★ قراءة البيانات الحالية مرة واحدة لحساب التكرار في الذاكرة
    var existingData = [];
    var headers = [];
    if (sheet.getLastRow() >= 2) {
      var allData = sheet.getDataRange().getValues();
      headers = allData.shift();
      existingData = allData;
    }
    var sidCol = headers.indexOf('رقم الطالب');
    var vidCol = headers.indexOf('رقم المخالفة');

    var now = new Date();
    var dayName = now.toLocaleDateString('ar-SA', { weekday: 'long' });
    var violTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    var hijriDate = getHijriDate_(now);
    var proceduresText = Array.isArray(data.procedures) ? data.procedures.map(function(p) { return sanitizeInput_(p); }).join('\n') : '';
    var formsText = Array.isArray(data.forms) ? data.forms.map(function(f) { return sanitizeInput_(f); }).join('\n') : '';

    var rows = [];
    for (var i = 0; i < data.students.length; i++) {
      var inp = data.students[i];
      var student = allStudents.find(function(s) { return s['رقم الطالب'] == inp.studentId; });
      if (!student) continue;

      // حساب التكرار من البيانات الحالية + الصفوف المبنية في هذه الدفعة
      var repeatLevel = 1;
      if (sidCol !== -1 && vidCol !== -1) {
        for (var r = 0; r < existingData.length; r++) {
          if (existingData[r][sidCol] == inp.studentId && existingData[r][vidCol] == data.violationId) repeatLevel++;
        }
        for (var j = 0; j < rows.length; j++) {
          if (rows[j][0] == inp.studentId && rows[j][4] == data.violationId) repeatLevel++;
        }
      }

      var points = getDeductionForViolation_(String(data.violationId), repeatLevel, student['المرحلة']);

      rows.push([
        student['رقم الطالب'],
        student['اسم الطالب'],
        student['الصف'],
        student['الفصل'],
        violation.id,
        violation.text,
        violation.type,
        getEffectiveDegree_(violation.id, student['المرحلة']),
        hijriDate,
        now,
        repeatLevel,
        proceduresText,
        points,
        dayName,
        formsText,
        'الوكيل',
        violTime,
        'لا'
      ]);
    }

    // ★ كتابة دفعة واحدة
    if (rows.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 9, rows.length, 1).setNumberFormat('@');
      sheet.getRange(startRow, 1, rows.length, 18).setValues(rows);
    }

    // مسح الكاش
    var cacheKey = 'violations_' + stage + '_' + new Date().toLocaleDateString('en-US');
    CacheService.getScriptCache().remove(cacheKey);

    return { success: true, message: 'تم حفظ ' + rows.length + ' مخالفة بنجاح', count: rows.length };
  } catch (e) {
    console.error("❌ خطأ في saveViolationsBatch:", e.toString());
    return { success: false, error: e.message };
  }
}

// =================================================================
// COMPENSATION - درجات التعويض
// =================================================================

/**
 * جلب المخالفات المؤهلة للتعويض (النقاط > 0) مع حالة التعويض
 * @param {string} stage - المرحلة
 * @return {Object} { success, records[], stats }
 */
function getCompensationEligibleRecords(stage) {
  try {
    var ss = getSpreadsheet_();

    // 1. جلب المخالفات
    var violSheetName = getSheetName_('المخالفات', stage);
    var violSheet = findSheet_(ss, violSheetName);
    if (!violSheet || violSheet.getLastRow() < 2) {
      return { success: true, records: [], stats: { total: 0, compensated: 0, pending: 0, totalPoints: 0 } };
    }

    var violData = violSheet.getDataRange().getValues();
    var violHeaders = violData.shift();

    // أعمدة المخالفات
    var colMap = {};
    violHeaders.forEach(function(h, i) { colMap[String(h).trim()] = i; });

    // 2. جلب سجلات السلوك الإيجابي (للتحقق من التعويض)
    var posSheetName = getSheetName_('السلوك_الإيجابي', stage);
    var posSheet = findSheet_(ss, posSheetName);
    var compensationRecords = [];

    // ★ بناء مجموعة المخالفات المعوَّضة (بالربط المباشر + الاحتياطي بالعدد)
    var compensatedViolRows = {};    // ربط مباشر: { صف_المخالفة: true }
    var compensationCountsFallback = {};  // احتياطي: { studentId: count }

    if (posSheet && posSheet.getLastRow() >= 2) {
      var posData = posSheet.getDataRange().getValues();
      var posHeaders = posData.shift();
      var posColMap = {};
      posHeaders.forEach(function(h, i) { posColMap[String(h).trim()] = i; });

      var violRefCol = posColMap['صف_المخالفة'];

      for (var p = 0; p < posData.length; p++) {
        var behavior = String(posData[p][posColMap['السلوك المتمايز'] !== undefined ? posColMap['السلوك المتمايز'] : posColMap['السلوك_المتمايز']] || '');
        var degree = String(posData[p][posColMap['الدرجة']] || '');
        if (behavior.indexOf('فرص تعويض') >= 0 || degree === 'تعويض') {
          var sid = String(posData[p][posColMap['رقم الطالب']] || '');
          // ★ ربط مباشر: إذا يوجد عمود صف_المخالفة وله قيمة
          var violRowRef = (violRefCol !== undefined) ? String(posData[p][violRefCol] || '').trim() : '';
          if (violRowRef && violRowRef !== '') {
            compensatedViolRows[violRowRef] = true;
          } else {
            // ★ احتياطي: للسجلات القديمة التي ليس فيها ربط مباشر
            compensationCountsFallback[sid] = (compensationCountsFallback[sid] || 0) + 1;
          }
        }
      }
    }

    // 4. بناء قائمة المخالفات المؤهلة
    var records = [];
    var stats = { total: 0, compensated: 0, pending: 0, totalPoints: 0 };
    // ★ تتبع عدد المخالفات المعلّمة لكل طالب (احتياطي للسجلات القديمة بدون ربط مباشر)
    var markedPerStudent = {};

    for (var i = 0; i < violData.length; i++) {
      var row = violData[i];
      var studentId = String(row[colMap['رقم الطالب']] || '');
      if (!studentId) continue;

      var points = parseFloat(row[colMap['النقاط']] || 0);
      if (!points || points <= 0) continue;

      var violationId = String(row[colMap['رقم المخالفة']] || '');
      var violationText = String(row[colMap['نص المخالفة']] || '');
      var violRowNum = String(i + 2); // رقم الصف الفعلي في الشيت

      // ★ التحقق من التعويض بطريقتين:
      // 1. ربط مباشر: هل يوجد سجل تعويض يشير لهذا الصف تحديداً؟
      var isCompensated = false;
      if (compensatedViolRows[violRowNum]) {
        isCompensated = true;
      } else {
        // 2. احتياطي (للسجلات القديمة): مطابقة بالعدد
        var maxComp = compensationCountsFallback[studentId] || 0;
        var alreadyMarked = markedPerStudent[studentId] || 0;
        if (alreadyMarked < maxComp) {
          isCompensated = true;
          markedPerStudent[studentId] = alreadyMarked + 1;
        }
      }

      stats.total++;
      stats.totalPoints += points;
      if (isCompensated) stats.compensated++;
      else stats.pending++;

      records.push({
        rowIndex: i + 2,
        studentId: studentId,
        studentName: String(row[colMap['اسم الطالب']] || ''),
        grade: String(row[colMap['الصف']] || ''),
        section: String(row[colMap['الفصل']] || ''),
        violationId: violationId,
        violationText: violationText,
        violationType: String(row[colMap['نوع المخالفة']] || ''),
        degree: String(row[colMap['الدرجة']] || ''),
        points: points,
        dateHijri: String(row[colMap['التاريخ الهجري']] || ''),
        dateMiladi: row[colMap['التاريخ الميلادي']] instanceof Date
          ? row[colMap['التاريخ الميلادي']].toISOString()
          : String(row[colMap['التاريخ الميلادي']] || ''),
        repeatLevel: String(row[colMap['مستوى التكرار']] || '1'),
        compensated: isCompensated
      });
    }

    return { success: true, records: records, stats: stats };

  } catch (e) {
    console.error('getCompensationEligibleRecords: ' + e.toString());
    return { success: false, error: e.message, records: [], stats: { total: 0, compensated: 0, pending: 0, totalPoints: 0 } };
  }
}

/**
 * حفظ سجل تعويضي في سجل السلوك الإيجابي
 * @param {Object} data - { studentId, stage, behaviorText, noorValue }
 * @return {Object} { success, message }
 */
function saveCompensationRecord(data) {
  try {
    if (!data || !data.studentId || !data.behaviorText) {
      return { success: false, error: 'بيانات غير مكتملة' };
    }

    var students = getStudents_();
    var student = students.find(function(s) { return s['رقم الطالب'] == data.studentId; });
    if (!student) return { success: false, error: 'الطالب غير موجود' };
    var stage = data.stage || student['المرحلة'] || 'متوسط';

    var ss = getSpreadsheet_();
    var sheetName = getSheetName_('السلوك_الإيجابي', stage);
    var sheet = findSheet_(ss, sheetName);

    // إنشاء الشيت إذا لم يكن موجوداً
    // ★ إضافة عمود "صف_المخالفة" لربط التعويض بمخالفة محددة
    var positiveHeaders = ['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'رقم الجوال',
      'السلوك المتمايز', 'الدرجة', 'المعلم', 'اليوم', 'التاريخ الهجري',
      'التاريخ الميلادي', 'وقت الإدخال', 'تم الإرسال', 'صف_المخالفة'];

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.setRightToLeft(true);
      sheet.appendRow(positiveHeaders);
      sheet.getRange(1, 1, 1, positiveHeaders.length)
        .setBackground('#10b981').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
      // حماية عمود التاريخ الهجري من التحويل التلقائي
      sheet.getRange(1, 10, sheet.getMaxRows(), 1).setNumberFormat('@');
      var regEntry = SHEET_REGISTRY['السلوك_الإيجابي'];
      if (regEntry && regEntry.color) sheet.setTabColor(regEntry.color);
    }

    // ★ ترقية الشيت القديم: إضافة عمود "صف_المخالفة" إذا لم يكن موجوداً
    var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (existingHeaders.indexOf('صف_المخالفة') === -1) {
      var newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue('صف_المخالفة').setFontWeight('bold');
    }

    var now = new Date();
    var dayName = now.toLocaleDateString('ar-SA', { weekday: 'long' });
    var posTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');

    // 14 عمود — مطابق لـ buildRowData_ في Server_TeacherInput.gs + عمود ربط المخالفة
    var violRef = data.violationRowIndex ? String(data.violationRowIndex) : '';
    // ★ إضافة رقم المخالفة ونصها في سجل السلوك للتوثيق
    var behaviorNote = data.behaviorText + ' (فرص تعويض)';
    if (data.violationId) behaviorNote += ' [مخالفة:' + data.violationId + ']';

    var newRow = [
      student['رقم الطالب'],                              // 1
      student['اسم الطالب'],                              // 2
      student['الصف'],                                   // 3
      student['الفصل'],                                  // 4
      student['رقم الجوال'] || '',                        // 5
      behaviorNote,                                       // 6: السلوك المتمايز (مع رقم المخالفة)
      'تعويض',                                           // 7: الدرجة
      'الوكيل',                                          // 8: المعلم
      dayName,                                           // 9: اليوم
      getHijriDate_(now),                                // 10: التاريخ الهجري
      now,                                               // 11: التاريخ الميلادي
      posTime,                                           // 12: وقت الإدخال
      'لا',                                              // 13: تم الإرسال
      violRef                                            // 14: صف_المخالفة (★ ربط مباشر)
    ];

    sheet.appendRow(newRow);

    // مسح الكاش
    var cacheKey = 'positive_' + stage + '_' + new Date().toLocaleDateString('en-US');
    CacheService.getScriptCache().remove(cacheKey);

    return {
      success: true,
      message: 'تم حفظ التعويض بنجاح',
      studentName: student['اسم الطالب'],
      behavior: data.behaviorText
    };

  } catch (e) {
    console.error('saveCompensationRecord: ' + e.toString());
    return { success: false, error: e.message };
  }
}

// =================================================================
// POSITIVE BEHAVIOR - السلوك المتمايز
// =================================================================

/**
 * جلب سجلات السلوك المتمايز لمرحلة معينة (مع كاش)
 */
function getCachedPositiveBehaviorRecords(stage) {
  var cacheKey = 'positive_' + stage + '_' + new Date().toLocaleDateString('en-US');
  var cache = CacheService.getScriptCache();
  // ★ حماية قراءة الكاش (قد يكون تالفاً)
  try {
    var cached = cache.get(cacheKey);
    if (cached != null) return JSON.parse(cached);
  } catch (e) {
    // تجاهل خطأ قراءة الكاش
  }

  var records = getPositiveBehaviorRecords(stage);
  // ★ حماية من تجاوز حد CacheService (100KB لكل مفتاح)
  try {
    var json = JSON.stringify(records);
    if (json.length < 90000) {
      cache.put(cacheKey, json, 300);
    }
  } catch(e) {
    // تجاهل خطأ الكاش — البيانات تُرجع بدونه
  }
  return records;
}

/**
 * جلب سجلات السلوك المتمايز من الشيت مباشرة
 */
function getPositiveBehaviorRecords(stage) {
  try {
    var sheetName = getSheetName_('السلوك_الإيجابي', stage);
    var ss = getSpreadsheet_();
    var sheet = findSheet_(ss, sheetName);
    
    if (!sheet || sheet.getLastRow() < 2) return [];
    
    var data = sheet.getDataRange().getValues();
    var headers = data.shift();
    
    return data.map(function(row, i) {
      var record = {};
      headers.forEach(function(header, index) {
        var val = row[index];
        if (val && val instanceof Date) {
          // ★ التاريخ الهجري يُحول لنص هجري وليس ISO
          if (header === 'التاريخ الهجري') {
            record[header] = readHijriCellValue_(val);
          } else {
            record[header] = val.toISOString();
          }
        } else {
          record[header] = (val !== null && val !== undefined && val !== '') ? val : '';
        }
      });
      record.rowIndex = i + 2; // ★ رقم الصف للتحديث لاحقاً
      return record;
    }).filter(function(record) { return record['رقم الطالب']; });

  } catch (e) {
    console.error('❌ Error fetching positive behavior records:', e.toString());
    return [];
  }
}
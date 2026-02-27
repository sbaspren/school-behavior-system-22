// =================================================================
// Server_Extension.gs
// ★ خدمة إضافة "رصد الغياب" — تجهيز بيانات الغياب اليومي للإضافة
// الإضافة تسحب البيانات عبر: ?page=extension
// الخرج: { medium: [{name, grade, periods}], high: [{name, grade, periods}] }
// =================================================================

/**
 * الدالة الرئيسية — تُستدعى من doGet عند page=extension
 * تُرجع نفس شكل JSON اللي تتوقعه إضافة "رصد الغياب"
 */
function getExtensionAbsenceData() {
  try {
    ensureStudentsSheetsLoaded_();
    var stages = Object.keys(STUDENTS_SHEETS);
    var result = {};
    // ★ بناء النتائج ديناميكياً لكل مرحلة مفعّلة
    // مع الحفاظ على التوافق مع الإضافة القديمة (medium/high)
    var stageToKey = { 'متوسط': 'medium', 'ثانوي': 'high', 'ابتدائي': 'primary', 'طفولة مبكرة': 'kindergarten' };
    for (var i = 0; i < stages.length; i++) {
      var key = stageToKey[stages[i]] || stages[i];
      result[key] = getStageAbsenceForExtension_(stages[i]);
    }
    return result;
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * جلب غياب اليوم لمرحلة معينة وتجهيزه بالشكل المطلوب
 * المدخل: 'متوسط' أو 'ثانوي'
 * المخرج: [{name: "أحمد", grade: "الأول المتوسط", periods: 3}, ...]
 *
 * ★ أعمدة الشيت الفعلية:
 *   A: رقم_الطالب | B: اسم_الطالب | C: الصف | D: الفصل
 *   E: رقم_الجوال | F: نوع_الغياب | G: الحصة | H: التاريخ_هجري
 *   I: اليوم | J: المسجل | K: وقت_الإدخال | L: حالة_الاعتماد
 *   M: نوع_العذر | N: تم_الإرسال | O: حالة_التأخر | P: وقت_الحضور | Q: ملاحظات
 *
 * ★ قيم نوع_الغياب الفعلية: "غائب" (يوم كامل) | "غائب عن حصة" (حصة واحدة)
 * ★ قيم الصف الفعلية: "الأول المتوسط" | "الثاني الثانوي المسار العام" (كامل مع المرحلة)
 */
function getStageAbsenceForExtension_(stage) {
  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var sheetName = 'سجل_الغياب_اليومي_' + stage;
  var sheet = findSheet_(ss, sheetName);
  
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  // بناء خريطة الأعمدة ديناميكياً
  var colMap = {};
  for (var h = 0; h < headers.length; h++) {
    var colName = String(headers[h] || '').trim().replace(/\s+/g, '_');
    if (colName) colMap[colName] = h;
  }
  
  // تحديد الأعمدة المطلوبة
  var nameCol = colMap['اسم_الطالب'];     // B
  var gradeCol = colMap['الصف'];           // C
  var typeCol = colMap['نوع_الغياب'];      // F
  var dateCol = colMap['وقت_الإدخال'];     // K
  
  if (nameCol === undefined) return [];
  
  // تاريخ اليوم (بدون وقت)
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var todayTime = today.getTime();
  
  // ★ تجميع الطلاب بالمفتاح: "اسم_الطالب|الصف"
  var studentMap = {};
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // ① فلترة بالتاريخ — فقط سجلات اليوم
    if (dateCol !== undefined) {
      var dateValue = row[dateCol];
      var rowDate;
      
      if (dateValue instanceof Date) {
        rowDate = new Date(dateValue);
      } else if (dateValue) {
        // محاولة تحويل النص لتاريخ
        rowDate = new Date(dateValue);
      } else {
        continue; // قيمة فارغة — تخطي
      }
      
      if (isNaN(rowDate.getTime())) continue; // تاريخ غير صالح — تخطي
      
      rowDate.setHours(0, 0, 0, 0);
      if (rowDate.getTime() !== todayTime) continue;
    } else {
      continue; // عمود التاريخ غير موجود — تخطي
    }
    
    // ② استخراج البيانات
    var studentName = String(row[nameCol] || '').trim();
    if (!studentName) continue;
    
    var gradeName = (gradeCol !== undefined) ? String(row[gradeCol] || '').trim() : '';
    var absenceType = (typeCol !== undefined) ? String(row[typeCol] || '').trim() : '';
    
    // ③ تحديد نوع الغياب من القيم الفعلية في الشيت
    // "غائب" = يوم كامل = 6 حصص
    // "غائب عن حصة" = حصة واحدة
    var isFullDay = (absenceType === 'غائب');
    
    // ④ تجميع بالمفتاح
    var key = studentName + '|' + gradeName;
    
    if (studentMap[key]) {
      if (isFullDay) {
        studentMap[key].hasFullDay = true;
      } else {
        studentMap[key].periodCount++;
      }
    } else {
      studentMap[key] = {
        name: studentName,
        grade: gradeName,
        hasFullDay: isFullDay,
        periodCount: isFullDay ? 0 : 1
      };
    }
  }
  
  // ⑤ تحويل إلى المصفوفة النهائية بنفس شكل الإضافة
  // { name: "أحمد", grade: "الأول المتوسط", periods: 6 }
  var result = [];
  for (var k in studentMap) {
    var s = studentMap[k];
    var periods = s.hasFullDay ? 6 : s.periodCount;
    result.push({
      name: s.name,
      grade: s.grade,
      periods: periods
    });
  }
  
  return result;
}


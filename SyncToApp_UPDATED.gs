// =================================================================
// SyncToApp_UPDATED.gs — كود محدث لمزامنة الغياب (17 عمود)
// ★ استبدل دالة syncStageToApp_ في الملف الخارجي بهذه النسخة
// =================================================================

/**
 * مزامنة مرحلة واحدة — محدث: يكتب 17 عمود بدل 15
 */
function syncStageToApp_(sourceSheet, appSS, stage, today, studentData) {
  var sheetName = 'سجل_الغياب_اليومي_' + stage;
  var appSheet = appSS.getSheetByName(sheetName);

  // ★ الهيدرز الموحدة (17 عمود)
  var HEADERS = [
    'رقم_الطالب', 'اسم_الطالب', 'الصف', 'الفصل', 'رقم_الجوال',
    'نوع_الغياب', 'الحصة', 'التاريخ_هجري', 'اليوم', 'المسجل',
    'وقت_الإدخال', 'حالة_الاعتماد', 'نوع_العذر', 'تم_الإرسال',
    'حالة_التأخر', 'وقت_الحضور', 'ملاحظات'
  ];

  if (!appSheet) {
    appSheet = appSS.insertSheet(sheetName);
    appSheet.setRightToLeft(true);
    appSheet.appendRow(HEADERS);
    appSheet.getRange(1, 1, 1, HEADERS.length).setBackground('#f3f4f6').setFontWeight('bold');
    appSheet.setFrozenRows(1);
  }

  // فحص السجلات الموجودة لتجنب التكرار
  var existingKeys = syncGetExistingKeys_(appSheet, today);

  var sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length < 2) return 0;

  var now = new Date();
  var hijriDate = syncGetHijriDate_(now);
  // ★ توحيد: تحويل أرقام عربية إلى غربية وإزالة "هـ"
  hijriDate = syncNormalizeHijri_(hijriDate);
  var dayName = syncGetDayName_(now);

  var newRows = [];

  for (var i = 1; i < sourceData.length; i++) {
    var row = sourceData[i];
    var studentName = String(row[0] || '').trim();
    if (!studentName) continue;

    // تجنب التكرار
    var key = studentName + '_' + today;
    if (existingKeys[key]) continue;

    // مطابقة الطالب
    var matched = syncFindStudent_(studentName, studentData);
    var studentId = matched ? matched.id : '';
    var phone = matched ? matched.phone : '';
    var parsed = syncParseClassName_(matched ? matched.className : '');
    var teacher = String(row[1] || '').trim();

    // ★ بناء الصف الموحد (17 عمود)
    newRows.push([
      studentId,            // 1:  رقم_الطالب
      studentName,          // 2:  اسم_الطالب
      parsed.grade,         // 3:  الصف
      parsed.classNum,      // 4:  الفصل
      phone,                // 5:  رقم_الجوال
      'يوم كامل',          // 6:  نوع_الغياب (موحد)
      '',                   // 7:  الحصة
      hijriDate,            // 8:  التاريخ_هجري
      dayName,              // 9:  اليوم
      teacher,              // 10: المسجل
      now,                  // 11: وقت_الإدخال (Date object)
      'معلق',              // 12: حالة_الاعتماد
      '',                   // 13: نوع_العذر
      'لا',                // 14: تم_الإرسال
      'غائب',              // 15: حالة_التأخر ← (كان ناقص!)
      '',                   // 16: وقت_الحضور ← (كان ناقص!)
      'مزامنة تلقائية'     // 17: ملاحظات ← (كان في عمود 15!)
    ]);

    existingKeys[key] = true;
  }

  if (newRows.length > 0) {
    appSheet.getRange(appSheet.getLastRow() + 1, 1, newRows.length, 17).setValues(newRows);
  }

  return newRows.length;
}

/**
 * ★ تحويل أرقام عربية → غربية وإزالة "هـ" — نسخة محلية للملف الخارجي
 */
function syncNormalizeHijri_(str) {
  if (!str) return str;
  var m = {'٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9'};
  var r = String(str);
  for (var k in m) r = r.split(k).join(m[k]);
  return r.replace(/\s*هـ\s*/g, '').replace(/[\u200e\u200f\u200b\u200c\u200d\u2066\u2067\u2068\u2069\u061c]/g, '').trim();
}

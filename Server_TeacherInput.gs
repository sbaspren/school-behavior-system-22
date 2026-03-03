// =================================================================
// Server_TeacherInput.gs - إدارة إدخالات المعلمين (نظام الروابط الفريدة)
// الإصدار: 5.0 (محدث - توجيه المرحلة + إشعار واتساب)
// التاريخ: فبراير 2026
// =================================================================

// ★ SPREADSHEET_URL محذوف — يُستخدم المتغير المركزي من Config.gs

// =================================================================
// ★★★ بناء بيانات صفحة المعلم (حقن مباشر - صاروخي)
// =================================================================
function buildTeacherPageData_(token) {
  try {
    if (!token) return { success: false, error: 'الرمز غير موجود' };

    var ss = getSpreadsheet_();
    var teachersSheet = ss.getSheetByName("المعلمين");
    if (!teachersSheet) return { success: false, error: 'شيت المعلمين غير موجود' };

    var data = teachersSheet.getDataRange().getValues();
    var headers = data[0];

    var nameCol = findColumnIndex(headers, ['الاسم', 'اسم المعلم']);
    var subjectCol = findColumnIndex(headers, ['المواد', 'المادة', 'التخصص']);
    var classesCol = findColumnIndex(headers, ['الفصول المسندة', 'الفصول', 'فصوله']);
    var tokenCol = findColumnIndex(headers, ['رمز_الرابط', 'الرمز', 'Token']);
    
    if (tokenCol === -1) return { success: false, error: 'عمود الرمز غير موجود' };
    
    // البحث عن المعلم بالرمز
    var teacherRow = null;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][tokenCol]).trim() === String(token).trim()) {
        teacherRow = data[i];
        break;
      }
    }
    if (!teacherRow) return { success: false, error: 'رابط غير صالح أو منتهي الصلاحية' };
    
    var teacherName = String(teacherRow[nameCol] || '');
    var teacherSubject = subjectCol > -1 ? String(teacherRow[subjectCol] || '') : '';
    
    // ★ تحليل فصول المعلم مع تحويل الحرف لرقم + استخراج المادة لكل فصل
    var rawClassesRaw = teacherRow[classesCol] ? 
      String(teacherRow[classesCol]).split(',').map(function(c){return c.trim();}).filter(function(c){return c;}) : [];
    
    // ★ فصل classKey عن المادة (التنسيق الجديد: "classKey:مادة")
    var rawClasses = [];
    var classSubjectMap = {};
    for (var ci = 0; ci < rawClassesRaw.length; ci++) {
      var colonIdx = rawClassesRaw[ci].indexOf(':');
      if (colonIdx > -1) {
        var ck = rawClassesRaw[ci].substring(0, colonIdx);
        classSubjectMap[ck] = rawClassesRaw[ci].substring(colonIdx + 1);
        rawClasses.push(ck);
      } else {
        rawClasses.push(rawClassesRaw[ci]);
      }
    }
    
    var letterToNum = {'أ':'1','ب':'2','ج':'3','د':'4','ه':'5','هـ':'5','و':'6','ز':'7','ح':'8','ط':'9'};
    
    // ★ قراءة جميع الطلاب من كل الشيتات (مرة واحدة)
    var sheets = getAllStudentsSheets_();
    var allStudents = {}; // "الأول متوسط|1" → [{i,n,p}, ...]
    
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s].sheet;
      var sData = sheet.getDataRange().getValues();
      if (sData.length < 2) continue;
      var sHeaders = sData[0];
      
      var sIdCol = findColumnIndex(sHeaders, ['رقم الطالب', 'المعرف', 'الرقم', 'السجل المدني', 'رقم الهوية']);
      var sNameCol = findColumnIndex(sHeaders, ['اسم الطالب', 'الاسم']);
      var sGradeCol = findColumnIndex(sHeaders, ['الصف', 'رقم الصف', 'اسم الصف']);
      var sClassCol = findColumnIndex(sHeaders, ['الفصل', 'شعبة']);
      var sPhoneCol = findColumnIndex(sHeaders, ['رقم الجوال', 'الجوال', 'جوال ولي الأمر']);
      
      for (var r = 1; r < sData.length; r++) {
        var grade = sGradeCol > -1 ? cleanGradeName_(sData[r][sGradeCol]) : '';
        var cls = sClassCol > -1 ? String(sData[r][sClassCol] || '').trim() : '';
        if (!grade || !sData[r][sNameCol > -1 ? sNameCol : 0]) continue;
        
        var key = grade + '|' + cls;
        if (!allStudents[key]) allStudents[key] = [];
        allStudents[key].push({
          i: sIdCol > -1 ? String(sData[r][sIdCol] || '') : '',
          n: sNameCol > -1 ? String(sData[r][sNameCol] || '') : '',
          p: sPhoneCol > -1 ? String(sData[r][sPhoneCol] || '') : ''
        });
      }
    }
    
    // ★ بناء بيانات الفصول مع طلابها
    var classesInfo = [];
    var studentsByClass = {};
    
    for (var j = 0; j < rawClasses.length; j++) {
      var parsed = parseClassKey_(rawClasses[j]);
      var classNum = letterToNum[parsed.letter] || parsed.letter;
      var gradeWithStage = (parsed.grade + ' ' + parsed.stageName).trim();
      var displayName = gradeWithStage + ' ' + classNum;
      
      var stage = parsed.stageName || '';
      if (!stage) {
        stage = detectStage_(gradeWithStage);
      }
      
      // ★ استخراج المادة المرتبطة بالفصل
      var classSubject = classSubjectMap[rawClasses[j]] || teacherSubject || '';
      
      classesInfo.push({
        d: displayName,
        g: gradeWithStage,
        c: classNum,
        s: stage,
        sub: classSubject
      });
      
      // ★ مطابقة الطلاب: صف + رقم الفصل
      var key = gradeWithStage + '|' + classNum;
      studentsByClass[displayName] = allStudents[key] || [];
    }
    
    // ★ جلب اسم المدرسة من الكليشة
    var schoolName = '';
    try { schoolName = getSchoolNameForLinks_(); } catch(e) { schoolName = 'المدرسة'; }

    return {
      success: true,
      sn: schoolName,
      t: { n: teacherName, s: teacherSubject },
      cl: classesInfo,
      st: studentsByClass
    };
    
  } catch(e) {
    Logger.log('Error in buildTeacherPageData_: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// ★★★ Trigger يومي - يبني بيانات كل المعلمين ويحفظها في الكاش
// =================================================================
function bakeAllTeachersData() {
  try {
    var ss = getSpreadsheet_();
    var teachersSheet = ss.getSheetByName("المعلمين");
    if (!teachersSheet) { Logger.log('شيت المعلمين غير موجود'); return; }
    
    var data = teachersSheet.getDataRange().getValues();
    var headers = data[0];
    var tokenCol = findColumnIndex(headers, ['رمز_الرابط', 'الرمز', 'Token']);
    if (tokenCol === -1) { Logger.log('عمود الرمز غير موجود'); return; }
    
    var cache = CacheService.getScriptCache();
    var baked = 0;
    
    for (var i = 1; i < data.length; i++) {
      var token = String(data[i][tokenCol] || '').trim();
      if (!token) continue;
      
      var pageData = buildTeacherPageData_(token);
      var json = JSON.stringify(pageData);
      
      try {
        cache.put('tpd_' + token, json, 21600); // 6 ساعات
        baked++;
      } catch(e) {
        Logger.log('فشل تخزين بيانات المعلم (حجم كبير): ' + token);
      }
    }
    
    Logger.log('✅ تم بناء بيانات ' + baked + ' معلم بنجاح');
    
  } catch(e) {
    Logger.log('❌ خطأ في bakeAllTeachersData: ' + e.toString());
  }
}

// =================================================================
// ★★★ إعداد Trigger يومي الساعة 5 صباحاً
// =================================================================
function setupDailyBakeTrigger() {
  // حذف أي trigger قديم
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'bakeAllTeachersData') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // إنشاء trigger جديد الساعة 5 صباحاً
  ScriptApp.newTrigger('bakeAllTeachersData')
    .timeBased()
    .atHour(5)
    .everyDays(1)
    .inTimezone('Asia/Riyadh')
    .create();
  
  Logger.log('✅ تم إعداد Trigger يومي الساعة 5 صباحاً بتوقيت الرياض');
}

// =================================================================
// ★★★ جلب بيانات المعلم (كاش أولاً، ثم حي)
// =================================================================
function getTeacherPageData_(token) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('tpd_' + token);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }
  
  var data = buildTeacherPageData_(token);
  
  try {
    cache.put('tpd_' + token, JSON.stringify(data), 21600);
  } catch(e) {}
  
  return data;
}

// =================================================================
// 1. توليد رمز فريد للمعلم
// =================================================================
function generateTeacherToken() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var token = '';
  for (var i = 0; i < 8; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

// =================================================================
// 2. إنشاء/تجديد رابط المعلم (يلغي القديم)
// =================================================================
function createTeacherLink(teacherId) {
  try {
    var ss = getSpreadsheet_();
    var teachersSheet = ss.getSheetByName("المعلمين");
    
    if (!teachersSheet) {
      return { success: false, error: 'شيت المعلمين غير موجود' };
    }
    
    var data = teachersSheet.getDataRange().getValues();
    var headers = data[0];
    
    var idCol = findColumnIndex(headers, ['المعرف', 'رقم المعلم', 'الرقم']);
    var nameCol = findColumnIndex(headers, ['الاسم', 'اسم المعلم']);
    var phoneCol = findColumnIndex(headers, ['الجوال', 'رقم الجوال', 'الهاتف']);
    var tokenCol = findColumnIndex(headers, ['رمز_الرابط', 'الرمز', 'Token']);
    var linkCol = findColumnIndex(headers, ['الرابط', 'رابط_النموذج', 'Link']);
    var activeDateCol = findColumnIndex(headers, ['تاريخ_التفعيل', 'تاريخ التفعيل']);

    if (idCol === -1) return { success: false, error: 'لم يتم العثور على عمود المعرف' };

    var lastCol = headers.length;

    if (tokenCol === -1) {
      tokenCol = lastCol;
      teachersSheet.getRange(1, lastCol + 1).setValue('رمز_الرابط');
      lastCol++;
    }
    
    if (linkCol === -1) {
      linkCol = lastCol;
      teachersSheet.getRange(1, lastCol + 1).setValue('الرابط');
      lastCol++;
    }
    
    if (activeDateCol === -1) {
      activeDateCol = lastCol;
      teachersSheet.getRange(1, lastCol + 1).setValue('تاريخ_التفعيل');
    }
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(teacherId)) {
        // ★ مسح كاش الرابط القديم (لو كان موجود) حتى لا يفتح ببيانات قديمة
        var oldToken = tokenCol >= 0 ? String(data[i][tokenCol] || '').trim() : '';
        if (oldToken) {
          try { CacheService.getScriptCache().remove('tpd_' + oldToken); } catch(e) {}
        }
        
        var newToken = generateTeacherToken();
        var baseUrl = ScriptApp.getService().getUrl();
        // ★ التأكد من استخدام رابط النشر (exec) وليس التجريب (dev)
        baseUrl = baseUrl.replace(/\/dev$/, '/exec');
        var newLink = baseUrl + '?page=teacher&token=' + newToken;
        
        teachersSheet.getRange(i + 1, tokenCol + 1).setValue(newToken);
        teachersSheet.getRange(i + 1, linkCol + 1).setValue(newLink);
        teachersSheet.getRange(i + 1, activeDateCol + 1).setValue(new Date());
        
        return {
          success: true,
          teacherName: data[i][nameCol],
          teacherPhone: data[i][phoneCol] || '',
          token: newToken,
          link: newLink,
          message: 'تم إنشاء رابط جديد للمعلم'
        };
      }
    }
    
    return { success: false, error: 'المعلم غير موجود' };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 2.5 إنشاء رابط بالجوال (تبحث عن المعلم/الإداري بالجوال وتنشئ الرابط)
// =================================================================
function createTeacherLinkByPhone_(phone, name) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName("المعلمين");
    if (!sheet) return { success: false, error: 'شيت المعلمين غير موجود' };
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var idCol = findColumnIndex(headers, ['المعرف', 'رقم المعلم', 'الرقم']);
    var nameCol = findColumnIndex(headers, ['الاسم', 'اسم المعلم']);
    var phoneCol = findColumnIndex(headers, ['الجوال', 'رقم الجوال', 'الهاتف']);

    if (phoneCol === -1) return { success: false, error: 'لم يتم العثور على عمود الجوال' };
    if (nameCol === -1) return { success: false, error: 'لم يتم العثور على عمود الاسم' };

    // البحث بالجوال أولاً (الأدق)
    var targetRow = -1;
    var nameMatchRow = -1;
    var nameMatchCount = 0;

    for (var i = 1; i < data.length; i++) {
      var rowPhone = formatPhone(String(data[i][phoneCol] || ''));
      var rowName = String(data[i][nameCol] || '').trim();

      // ★ الجوال أولوية مطلقة
      if (rowPhone && rowPhone === phone) {
        targetRow = i;
        break;
      }
      // ★ البحث بالاسم كبديل — لكن فقط إذا كان فريداً
      if (name && rowName === name) {
        nameMatchRow = i;
        nameMatchCount++;
      }
    }

    // إذا ما لقينا بالجوال، نستخدم الاسم فقط إذا كان فريداً
    if (targetRow === -1 && nameMatchCount === 1) {
      targetRow = nameMatchRow;
    }

    if (targetRow === -1) {
      return { success: false, error: 'المعلم غير موجود: ' + name + (nameMatchCount > 1 ? ' (يوجد أكثر من معلم بنفس الاسم)' : '') };
    }

    // استخدام createTeacherLink بالمعرف
    var teacherId = String(data[targetRow][idCol] || data[targetRow][0] || '');
    return createTeacherLink(teacherId);
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function createUserLinkByPhone_(phone, name) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName("المستخدمين");
    if (!sheet) return { success: false, error: 'شيت المستخدمين غير موجود' };
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var idCol = findColumnIndex(headers, ['المعرف', 'ID']);
    var nameCol = findColumnIndex(headers, ['الاسم']);
    var phoneCol = findColumnIndex(headers, ['الجوال', 'رقم الجوال']);

    if (phoneCol === -1) return { success: false, error: 'لم يتم العثور على عمود الجوال' };
    if (nameCol === -1) return { success: false, error: 'لم يتم العثور على عمود الاسم' };

    // البحث بالجوال أولاً (الأدق)
    var targetRow = -1;
    var nameMatchRow = -1;
    var nameMatchCount = 0;

    for (var i = 1; i < data.length; i++) {
      var rowPhone = formatPhone(String(data[i][phoneCol] || ''));
      var rowName = String(data[i][nameCol] || '').trim();
      
      // ★ الجوال أولوية مطلقة
      if (rowPhone && rowPhone === phone) {
        targetRow = i;
        break;
      }
      // ★ البحث بالاسم كبديل — لكن فقط إذا كان فريداً
      if (name && rowName === name) {
        nameMatchRow = i;
        nameMatchCount++;
      }
    }
    
    // إذا ما لقينا بالجوال، نستخدم الاسم فقط إذا كان فريداً
    if (targetRow === -1 && nameMatchCount === 1) {
      targetRow = nameMatchRow;
    }
    
    if (targetRow === -1) {
      return { success: false, error: 'المستخدم غير موجود: ' + name + (nameMatchCount > 1 ? ' (يوجد أكثر من مستخدم بنفس الاسم)' : '') };
    }
    
    var userId = String(data[targetRow][idCol] || data[targetRow][0] || '');
    return createUserLink(userId);
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 3. تحويل مفتاح الفصل إلى اسم عربي قابل للعرض والمطابقة
// =================================================================
// المفتاح: الأول_intermediate_أ → العرض: الأول متوسط أ
function parseClassKey_(classKey) {
  var key = String(classKey || '').trim();
  if (!key) return { display: '', grade: '', stageName: '', letter: '' };
  
  var parts = key.split('_');
  if (parts.length < 3) {
    // تنسيق غير معروف - أعده كما هو
    return { display: key, grade: key, stageName: '', letter: '' };
  }
  
  var letter = parts[parts.length - 1];
  var stageId = parts[parts.length - 2];
  var gradeName = parts.slice(0, parts.length - 2).join(' ');
  
  // تحويل معرّف المرحلة الإنجليزي إلى اسم عربي
  var stageMap = {
    'kindergarten': 'طفولة مبكرة',
    'primary': 'ابتدائي',
    'intermediate': 'متوسط',
    'secondary': 'ثانوي'
  };
  var stageName = stageMap[stageId] || stageId;
  
  // الاسم المعروض: الأول متوسط أ
  var display = gradeName + ' ' + stageName + ' ' + letter;
  
  return {
    display: display,
    grade: gradeName,
    stageName: stageName,
    stageId: stageId,
    letter: letter
  };
}

// تطبيع اسم الصف لمقارنة مرنة (حذف ال التعريف)
function normalizeForMatch_(text) {
  return String(text || '')
    .replace(/^ال/g, '')
    .replace(/(\s)ال/g, '$1')
    .replace(/\s+/g, ' ')
    .trim();
}

// =================================================================
// 3.5 التحقق من الرمز وجلب بيانات المعلم (مع تحويل الفصول + الصلاحيات)
// =================================================================
function getTeacherByToken(token) {
  try {
    if (!token) {
      return { success: false, error: 'الرمز غير موجود' };
    }
    
    var ss = getSpreadsheet_();
    var teachersSheet = ss.getSheetByName("المعلمين");
    
    if (!teachersSheet) {
      return { success: false, error: 'شيت المعلمين غير موجود' };
    }
    
    var data = teachersSheet.getDataRange().getValues();
    var headers = data[0];
    
    var idCol = findColumnIndex(headers, ['المعرف', 'رقم المعلم', 'الرقم']);
    var nameCol = findColumnIndex(headers, ['الاسم', 'اسم المعلم']);
    var subjectCol = findColumnIndex(headers, ['المواد', 'المادة', 'التخصص']);
    var classesCol = findColumnIndex(headers, ['الفصول المسندة', 'الفصول', 'فصوله']);
    var tokenCol = findColumnIndex(headers, ['رمز_الرابط', 'الرمز', 'Token']);

    if (tokenCol === -1) {
      return { success: false, error: 'عمود الرمز غير موجود - يرجى إنشاء رابط أولاً' };
    }
    if (idCol === -1) return { success: false, error: 'لم يتم العثور على عمود المعرف' };
    if (nameCol === -1) return { success: false, error: 'لم يتم العثور على عمود الاسم' };

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][tokenCol]).trim() === String(token).trim()) {

        // ★ تحويل مفاتيح الفصول إلى أسماء عربية للعرض والمطابقة
        var rawClasses = classesCol > -1 && data[i][classesCol] ?
          String(data[i][classesCol]).split(',').map(function(c) { return c.trim(); }).filter(function(c) { return c; }) : [];

        var displayClasses = [];
        for (var j = 0; j < rawClasses.length; j++) {
          var parsed = parseClassKey_(rawClasses[j]);
          displayClasses.push(parsed.display);
        }

        // ★ المواد
        var subjects = subjectCol > -1 && data[i][subjectCol] ?
          String(data[i][subjectCol]).split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; }) : [];

        return {
          success: true,
          teacher: {
            id: String(data[i][idCol] || ''),
            name: String(data[i][nameCol] || ''),
            subject: subjects.join('، '),
            classes: displayClasses
          }
        };
      }
    }
    
    return { success: false, error: 'رابط غير صالح أو منتهي الصلاحية' };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 4. جلب بيانات المعلم للنموذج (بالرمز)
// =================================================================
function getTeacherFormData(token) {
  return getTeacherByToken(token);
}

// =================================================================
// ★ الدوال التالية حُذفت (كود قديم استُبدل بالقسم 10):
// - sendTeacherLinkWhatsApp → sendLinkToPersonWithStage
// - sendAllTeachersLinks → bulkSendLinks
// - getTeachersWithLinks → getLinksTabData

// =================================================================
// 8. جلب طلاب الفصل (من شيتات المراحل)
// =================================================================
function getClassStudents(className) {
  try {
    var sheets = getAllStudentsSheets_();
    var students = [];
    
    // ★ تحليل اسم الفصل المعروض (مثلاً: "الأول متوسط أ")
    var classNameStr = String(className || '').trim();
    var classNameParts = classNameStr.split(' ');
    var targetLetter = classNameParts.length > 1 ? classNameParts[classNameParts.length - 1] : '';
    var targetGrade = classNameParts.length > 1 ? classNameParts.slice(0, -1).join(' ') : classNameStr;
    var targetGradeNorm = normalizeForMatch_(targetGrade);
    
    // ★ استخراج اسم الصف بدون المرحلة (الأول متوسط → الأول)
    var stageWords = ['متوسط', 'المتوسط', 'ثانوي', 'الثانوي', 'ابتدائي', 'الابتدائي'];
    var targetGradeNoStage = targetGrade;
    for (var sw = 0; sw < stageWords.length; sw++) {
      targetGradeNoStage = targetGradeNoStage.replace(new RegExp('\\s*' + stageWords[sw] + '\\s*', 'g'), ' ').trim();
    }
    var targetGradeNoStageNorm = normalizeForMatch_(targetGradeNoStage);
    
    // ★ اكتشاف المرحلة المستهدفة من اسم الفصل
    var targetStage = '';
    if (classNameStr.indexOf('ثانوي') !== -1 || classNameStr.indexOf('ثانوى') !== -1) targetStage = 'ثانوي';
    else if (classNameStr.indexOf('متوسط') !== -1) targetStage = 'متوسط';
    else if (classNameStr.indexOf('ابتدائي') !== -1 || classNameStr.indexOf('ابتدائى') !== -1) targetStage = 'ابتدائي';
    
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s].sheet;
      var sheetStage = sheets[s].stage; // المرحلة من اسم الشيت
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;
      var headers = data[0];
      
      var idCol = findColumnIndex(headers, ['رقم الطالب', 'المعرف', 'الرقم', 'السجل المدني', 'رقم الهوية']);
      var nameCol = findColumnIndex(headers, ['اسم الطالب', 'الاسم']);
      var classCol = findColumnIndex(headers, ['الفصل', 'فصل', 'شعبة']);
      var gradeCol = findColumnIndex(headers, ['الصف', 'رقم الصف', 'المرحلة', 'اسم الصف']);
      var phoneCol = findColumnIndex(headers, ['جوال ولي الأمر', 'الجوال', 'رقم الجوال']);
      
      for (var i = 1; i < data.length; i++) {
        var grade = gradeCol > -1 ? cleanGradeName_(data[i][gradeCol]) : '';
        var cls = classCol > -1 ? String(data[i][classCol] || '').trim() : '';
        
        // ★ مطابقة حرف الفصل
        var letterMatch = (cls === targetLetter);
        if (!letterMatch) continue; // تخطي مبكر لتسريع
        
        var gradeNorm = normalizeForMatch_(grade);
        var gradeMatch = false;
        
        // ★ طريقة 1: مطابقة كاملة (الصف يشمل المرحلة مثل "الأول متوسط")
        if (gradeNorm === targetGradeNorm || 
            gradeNorm.indexOf(targetGradeNorm) !== -1 || 
            targetGradeNorm.indexOf(gradeNorm) !== -1) {
          gradeMatch = true;
        }
        
        // ★ طريقة 2: الصف بدون مرحلة + التأكد من الشيت الصحيح
        // مثال: الصف = "الأول" في شيت "طلاب_متوسط" والمطلوب "الأول متوسط"
        if (!gradeMatch && targetStage && sheetStage === targetStage) {
          var gradeNoStage = grade;
          for (var sw2 = 0; sw2 < stageWords.length; sw2++) {
            gradeNoStage = gradeNoStage.replace(new RegExp('\\s*' + stageWords[sw2] + '\\s*', 'g'), ' ').trim();
          }
          var gradeNoStageNorm = normalizeForMatch_(gradeNoStage);
          
          if (gradeNoStageNorm === targetGradeNoStageNorm ||
              gradeNoStageNorm.indexOf(targetGradeNoStageNorm) !== -1 ||
              targetGradeNoStageNorm.indexOf(gradeNoStageNorm) !== -1) {
            gradeMatch = true;
          }
        }
        
        // ★ طريقة 3: مطابقة مرنة - اسم الفصل الكامل
        if (!gradeMatch) {
          var fullClass = (grade + ' ' + cls).trim();
          var fullClassNorm = normalizeForMatch_(fullClass);
          var targetFullNorm = normalizeForMatch_(classNameStr);
          if (fullClassNorm === targetFullNorm || 
              fullClassNorm.indexOf(targetFullNorm) !== -1 || 
              targetFullNorm.indexOf(fullClassNorm) !== -1) {
            gradeMatch = true;
          }
        }
        
        if (gradeMatch) {
          students.push({
            id: idCol > -1 ? String(data[i][idCol] || '') : '',
            name: nameCol > -1 ? String(data[i][nameCol] || '') : '',
            grade: grade,
            class: cls,
            phone: phoneCol > -1 ? String(data[i][phoneCol] || '') : ''
          });
        }
      }
    }
    
    return { success: true, students: students };
    
  } catch (e) {
    return { success: false, error: e.toString(), students: [] };
  }
}

// =================================================================
// 9. حفظ إدخال المعلم (مع توجيه المرحلة + إشعار واتساب)
// =================================================================
function submitTeacherForm(formData) {
  try {
    var ss = getSpreadsheet_();
    var now = new Date();
    var hijriDate = getHijriDate_(now);
    var dayName = getDayNameAr_(now);
    
    // ★ حالة "لا يوجد غائب" — تسجيل تأكيد حضور جميع الطلاب
    if (formData.noAbsence === true && formData.inputType === 'absence') {
      var nStage = formData.stage || detectStage_(formData.className);
      if (!nStage) return { success: false, error: 'لم يتم تحديد المرحلة' };
      var nSheet = ensureDailyAbsenceSheet_(nStage);
      // تسجيل صف خاص يدل على "لا يوجد غائب"
      var classParts = String(formData.className || '').trim().split(/\s+/);
      var gradeName = classParts.length > 1 ? classParts.slice(0, -1).join(' ') : formData.className;
      var sectionNum = classParts.length > 1 ? classParts[classParts.length - 1] : '';
      // ★ استخدام writeDailyAbsenceRows_ لضمان تنسيق '@' على عمود التاريخ الهجري
      writeDailyAbsenceRows_(nSheet, [[
        'NO_ABSENCE',                                      // 1: رقم_الطالب (علامة خاصة)
        'لا يوجد غائب',                                   // 2: اسم_الطالب
        sanitizeInput_(gradeName),                          // 3: الصف
        sanitizeInput_(sectionNum),                         // 4: الفصل
        '',                                                 // 5: رقم_الجوال
        'يوم كامل',                                        // 6: نوع_الغياب
        '',                                                 // 7: الحصة
        hijriDate,                                          // 8: التاريخ_هجري
        dayName,                                            // 9: اليوم
        sanitizeInput_(formData.teacherName || ''),         // 10: المسجل
        now,                                                // 11: وقت_الإدخال (Date object)
        'مؤكد',                                             // 12: حالة_الاعتماد
        '',                                                 // 13: نوع_العذر
        'نعم',                                              // 14: تم_الإرسال
        'حاضر',                                             // 15: حالة_التأخر
        '',                                                 // 16: وقت_الحضور
        '',                                                 // 17: ملاحظات
        ''                                                  // 18: حالة_نور
      ]]);
      logActivity_(formData, 0);
      return {
        success: true,
        message: 'تم تأكيد حضور جميع طلاب ' + formData.className + ' ✅',
        count: 0,
        records: [],
        stage: nStage
      };
    }
    
    // ★ تحليل اسم الفصل إلى اسم الصف ورقم الفصل
    // مثلاً: "الأول ثانوي 1" → الصف: "الأول ثانوي"، الفصل: "1"
    var classNameParts = String(formData.className || '').trim().split(/\s+/);
    if (classNameParts.length > 1) {
      formData._grade = classNameParts.slice(0, -1).join(' ');
      formData._section = classNameParts[classNameParts.length - 1];
    } else {
      formData._grade = formData.className || '';
      formData._section = '';
    }
    
    // اكتشاف المرحلة: من الواجهة أولاً، ثم من اسم الفصل
    var stage = formData.stage || detectStage_(formData.className);
    var sheetName = getTargetSheetName_(formData.inputType, stage);
    
    if (!sheetName) {
      return { success: false, error: 'نوع إدخال غير صالح' };
    }
    
    var sheet = findSheet_(ss, sheetName);
    
    if (!sheet) {
      sheet = createSheet_(ss, sheetName, formData.inputType);
    }
    
    var savedCount = 0;
    var savedRecords = [];

    // ★ الغياب اليومي: استخدام writeDailyAbsenceRows_ لضمان تنسيق '@' على عمود التاريخ الهجري
    // appendRow لا يضمن حفظ التاريخ الهجري كنص مما يسبب عدم ظهوره في التوثيق
    if (formData.inputType === 'absence') {
      var absRows = [];
      for (var i = 0; i < formData.students.length; i++) {
        var student = formData.students[i];
        absRows.push(buildRowData_(formData, student, now, hijriDate, dayName));
        savedRecords.push({
          studentName: student.name,
          studentId: student.id
        });
      }
      savedCount = writeDailyAbsenceRows_(sheet, absRows);
    } else {
      for (var i = 0; i < formData.students.length; i++) {
        var student = formData.students[i];
        var row = buildRowData_(formData, student, now, hijriDate, dayName);
        sheet.appendRow(row);
        savedCount++;

        savedRecords.push({
          studentName: student.name,
          studentId: student.id
        });
      }
    }
    
    // تسجيل النشاط
    logActivity_(formData, savedCount);
    
    // إشعار الوكيل عبر واتساب (إذا مطلوب)
    if (formData.notifyDeputy) {
      notifyDeputyWhatsApp_(formData, savedRecords, stage);
    }
    
    return { 
      success: true, 
      message: 'تم إرسال ' + savedCount + ' سجل بنجاح إلى وكيل ' + stage,
      count: savedCount,
      records: savedRecords,
      stage: stage
    };
    
  } catch (e) {
    Logger.log('خطأ في submitTeacherForm: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 10. إشعار الوكيل عبر واتساب (حسب المرحلة)
// =================================================================
function notifyDeputyWhatsApp_(formData, records, stage) {
  try {
    var typeLabels = {
      'absence': '⚠️ غياب',
      'violation': '🚫 مخالفة سلوكية',
      'note': '📝 ملاحظة تربوية',
      'positive': '⭐ سلوك متمايز',
      'custom': '📋 ملاحظة خاصة'
    };
    
    var message = '📋 *إشعار من نموذج المعلم*\n';
    message += '━━━━━━━━━━━━━━━\n\n';
    message += '👤 المعلم: ' + formData.teacherName + '\n';
    message += '📚 الفصل: ' + formData.className + '\n';
    message += '📝 النوع: ' + (typeLabels[formData.inputType] || formData.inputType) + '\n';
    message += '👥 العدد: ' + records.length + ' طالب\n';
    
    if (formData.itemText) {
      message += '📄 التفاصيل: ' + formData.itemText + '\n';
    }
    
    message += '\n━━━━━━━━━━━━━━━\n';
    
    // أسماء الطلاب (أول 5 فقط)
    var maxShow = Math.min(records.length, 5);
    for (var i = 0; i < maxShow; i++) {
        message += (i + 1) + '. ' + records[i].studentName + '\n';
    }
    if (records.length > 5) {
        message += '... و ' + (records.length - 5) + ' آخرين\n';
    }
    
    message += '\n⏰ ' + new Date().toLocaleTimeString('ar-SA');
    
    // ★ التوجيه الذكي: البحث عن الوكيل المسؤول عن هذا الفصل تحديداً
    var specificDeputyPhone = getResponsibleDeputyPhone_(formData.className, stage);
    
    // إرسال عبر واتساب إلى الوكيل المحدد (أو المرحلة كاحتياط)
    if (typeof sendWhatsAppMessage === 'function') {
        if (specificDeputyPhone) {
            sendWhatsAppMessage(specificDeputyPhone, message, stage);
            Logger.log('تم إرسال إشعار الوكيل المتخصص للفصل: ' + formData.className + ' على الرقم: ' + specificDeputyPhone);
        } else {
             sendWhatsAppMessage(null, message, stage);
             Logger.log('تم إرسال إشعار الوكيل العام لمرحلة ' + stage);
        }
    }
    
  } catch (e) {
    Logger.log('خطأ في إرسال إشعار الوكيل: ' + e.toString());
  }
}

// ★ دالة مساعدة لجلب رقم جوال الوكيل المسؤول عن فصل معين
function getResponsibleDeputyPhone_(className, stage) {
  try {
    var ss = getSpreadsheet_();
    var usersSheet = ss.getSheetByName("المستخدمين");
    if (!usersSheet) return null;
    
    var data = usersSheet.getDataRange().getValues();
    if (data.length < 2) return null;
    
    var headers = data[0];
    var roleCol = findColumnIndex(headers, ['الدور', 'المسمى', 'الوظيفة']);
    var phoneCol = findColumnIndex(headers, ['الجوال', 'رقم الجوال']);
    var scopeCol = findColumnIndex(headers, ['قيمة النطاق', 'النطاق', 'classes']);
    var typeCol = findColumnIndex(headers, ['نوع النطاق']);
    
    var targetClassName = String(className || '').trim();
    var stageDeputyPhone = null; // الاحتياط (وكيل المرحلة كاملة)
    
    for (var i = 1; i < data.length; i++) {
      var role = String(data[i][roleCol] || '').trim();
      if (role.indexOf('وكيل') === -1 && role.indexOf('مدير') === -1) continue; // نبحث عن وكيل أو مدير
      
      var phone = String(data[i][phoneCol] || '').trim();
      if (!phone) continue;
      
      var scopeType = typeCol > -1 ? String(data[i][typeCol] || '').trim() : '';
      var scopeValue = String(data[i][scopeCol] || '').trim();
      
      // 1. هل هو مسؤول عن هذا الفصل تحديداً؟
      if (scopeType === 'classes' || (!scopeType && scopeValue && scopeValue !== 'متوسط' && scopeValue !== 'ثانوي' && scopeValue !== 'ابتدائي')) {
          var assignedClasses = scopeValue.split(',').map(function(c) { return c.trim(); });
          if (assignedClasses.indexOf(targetClassName) !== -1 || assignedClasses.indexOf(targetClassName.replace(/\s+/g,' ')) !== -1) {
              return formatPhone(phone); // وجدنا المطابقة الدقيقة
          }
      }
      
      // 2. هل هو وكيل للمرحلة كاملة؟ (نحفظه كاحتياط اذا لم نجد مخصص)
      if (scopeType === 'stage' || scopeValue === stage || (!scopeValue && stage)) {
          stageDeputyPhone = formatPhone(phone);
      }
    }
    
    return stageDeputyPhone; // قد يكون null إذا لم نجد أحداً
  } catch(e) {
    Logger.log('خطأ في استخراج وكيل الفصل: ' + e);
    return null;
  }
}

// =================================================================
// الدوال المساعدة
// =================================================================

function findColumnIndex(headers, possibleNames) {
  // 1. مطابقة تامة أولاً
  for (var i = 0; i < possibleNames.length; i++) {
    var idx = headers.indexOf(possibleNames[i]);
    if (idx !== -1) return idx;
  }
  // 2. مطابقة جزئية آمنة: الرأس يحتوي على الاسم ككلمة كاملة (وليس كجزء من كلمة أخرى)
  for (var h = 0; h < headers.length; h++) {
    var headerStr = String(headers[h]).trim();
    for (var n = 0; n < possibleNames.length; n++) {
      var name = possibleNames[n];
      // تطابق فقط إذا كان الاسم يظهر ككلمة كاملة داخل الرأس
      var regex = new RegExp('(^|[\\s_\\-/])' + name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '([\\s_\\-/]|$)');
      if (headerStr === name || regex.test(headerStr)) {
        return h;
      }
    }
  }
  return -1; // ★ إصلاح: كانت 0 — تسبب في الكتابة على العمود الأول بالخطأ
}

function formatPhone(phone) {
  phone = String(phone).replace(/\D/g, '');
  if (phone.startsWith('05')) {
    phone = '966' + phone.substring(1);
  } else if (phone.startsWith('5')) {
    phone = '966' + phone;
  }
  return phone;
}

function detectStage_(className) {
  if (!className) return '';
  var cn = String(className);

  if (cn.indexOf('ثانوي') !== -1 || cn.indexOf('ثانوى') !== -1) return 'ثانوي';
  if (cn.indexOf('متوسط') !== -1) return 'متوسط';
  if (cn.indexOf('ابتدائي') !== -1 || cn.indexOf('ابتدائى') !== -1) return 'ابتدائي';
  if (cn.indexOf('طفولة') !== -1 || cn.indexOf('روضة') !== -1) return 'طفولة مبكرة';
  // لا افتراضي — أرجع فارغ لتجنب حفظ في مرحلة خاطئة
  return '';
}

function getTargetSheetName_(inputType, stage) {
  if (!stage) return null;
  // ★ تحويل نوع الإدخال الإنجليزي إلى مفتاح SHEET_REGISTRY العربي
  var typeToRegistryKey = {
    'absence': 'الغياب_اليومي',
    'violation': 'المخالفات',
    'note': 'الملاحظات_التربوية',
    'positive': 'السلوك_الإيجابي',
    'custom': 'الملاحظات_التربوية'
  };
  var registryKey = typeToRegistryKey[inputType];
  if (!registryKey) return null;
  return getSheetName_(registryKey, stage);
}

// ★ findSheet_ محذوفة من هنا - تستخدم النسخة الشاملة في Config.gs
// التي تدعم SHEET_ALIASES للتوافق مع كل الأسماء القديمة والجديدة

function createSheet_(ss, sheetName, inputType) {
  var sheet = ss.insertSheet(sheetName);
  sheet.setRightToLeft(true);
  
  var headers = getSheetHeaders_(inputType);
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1e3a5f')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);

  // حماية عمود التاريخ الهجري من التحويل التلقائي
  var hijriColMap = { 'violation': 9, 'positive': 10, 'note': 9, 'custom': 9, 'absence': 8 };
  var hijriCol = hijriColMap[inputType];
  if (hijriCol) {
    sheet.getRange(1, hijriCol, sheet.getMaxRows(), 1).setNumberFormat('@');
  }

  return sheet;
}

function getSheetHeaders_(inputType) {
  // ★ يجب أن تتطابق مع أعمدة الواجهة الرئيسية (الوكيل)
  switch (inputType) {
    case 'absence':
      // ★ الغياب اليومي - 18 عمود (مطابق لـ getDailyAbsenceHeaders_ في Server_Absence_Daily)
      return ['رقم_الطالب', 'اسم_الطالب', 'الصف', 'الفصل', 'رقم_الجوال',
        'نوع_الغياب', 'الحصة', 'التاريخ_هجري', 'اليوم', 'المسجل',
        'وقت_الإدخال', 'حالة_الاعتماد', 'نوع_العذر', 'تم_الإرسال',
        'حالة_التأخر', 'وقت_الحضور', 'ملاحظات', 'حالة_نور'];
    case 'violation':
      // ★ المخالفات - 18 عمود (مطابق لـ Server_Actions)
      return ['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل',
        'رقم المخالفة', 'نص المخالفة', 'نوع المخالفة', 'الدرجة',
        'التاريخ الهجري', 'التاريخ الميلادي', 'مستوى التكرار', 'الإجراءات',
        'النقاط', 'اليوم', 'النماذج المحفوظة', 'المستخدم', 'وقت الإدخال', 'تم الإرسال'];
    case 'note':
    case 'custom':
      // ★ الملاحظات التربوية - 11 عمود (مطابق لـ Server_EducationalNotes)
      return ['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'رقم الجوال',
        'نوع الملاحظة', 'التفاصيل', 'المعلم/المسجل',
        'التاريخ', 'وقت الإدخال', 'تم الإرسال'];
    case 'positive':
      // ★ السلوك المتمايز - 13 عمود
      return ['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'رقم الجوال',
        'السلوك المتمايز', 'الدرجة', 'المعلم',
        'اليوم', 'التاريخ الهجري', 'التاريخ الميلادي', 'وقت الإدخال', 'تم الإرسال'];
    default:
      return [];
  }
}

// ★ حماية من حقن الصيغ (Formula Injection) — تزيل HTML ثم تمنع بداية الخلية بـ = + - @
function sanitizeInput_(str) {
  var s = stripHtml_(str);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function buildRowData_(formData, student, now, hijriDate, dayName) {

  // ★ استخدام الصف والفصل المحللين من اسم الفصل المحدد
  // formData._grade = اسم الصف (مثل "الأول ثانوي")
  // formData._section = رقم الفصل (مثل "1")
  var gradeName = sanitizeInput_(formData._grade || student.grade || '');
  var sectionNum = sanitizeInput_(formData._section || student.class || '');
  var sTeacherName = sanitizeInput_(formData.teacherName || '');
  var sItemText = sanitizeInput_(formData.itemText || '');
  var sItemDegree = sanitizeInput_(formData.itemDegree || '');
  var sStudentName = sanitizeInput_(student.name || '');
  var sStudentId = sanitizeInput_(student.id || '');
  var sStudentPhone = sanitizeInput_(student.phone || '');
  
  switch (formData.inputType) {
    case 'absence':
      // ★ الغياب اليومي — يستخدم الدالة المركزية الموحدة
      return buildDailyAbsenceRow_({
        studentId:   student.id,
        studentName: student.name,
        grade:       formData._grade || student.grade,
        section:     formData._section || student.class,
        phone:       student.phone,
        absenceType: formData.absenceType,
        period:      formData.teacherSubject,
        recorder:    formData.teacherName,
        notes:       '',
        dateOverride: now
      });
      
    case 'violation':
      // ★ المخالفات - 18 عمود
      var violTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
      return [
        sStudentId,                                // 1: رقم الطالب
        sStudentName,                              // 2: اسم الطالب
        gradeName,                                 // 3: الصف
        sectionNum,                                // 4: الفصل
        sanitizeInput_(formData.itemId || ''),         // 5: رقم المخالفة
        sItemText,                                 // 6: نص المخالفة
        sanitizeInput_(formData.violationType || ''), // 7: نوع المخالفة (حضوري/رقمي/هيئة)
        sItemDegree,                               // 8: الدرجة
        hijriDate,                                 // 9: التاريخ الهجري
        now,                                       // 10: التاريخ الميلادي
        1,                                         // 11: مستوى التكرار (افتراضي 1)
        '',                                        // 12: الإجراءات
        0,                                         // 13: النقاط (الدرجة المخصومة - يحددها الوكيل)
        dayName,                                   // 14: اليوم (السبت، الأحد، ...)
        '',                                        // 15: النماذج المحفوظة
        sTeacherName,                              // 16: المستخدم
        violTime,                                  // 17: وقت الإدخال (الوقت فقط)
        'لا'                                       // 18: تم الإرسال
      ];
      
    case 'note':
      // ★ الملاحظات التربوية - 11 عمود (مطابق للوكيل والشيت)
      var noteClass = sanitizeInput_(formData.noteClassification || 'سلبي');
      // دمج التصنيف مع التفاصيل
      var detailsWithClass = sanitizeInput_(formData.details || '');
      if (noteClass) detailsWithClass = '[' + noteClass + '] ' + detailsWithClass;
      return [
        sStudentId,                                // 1: رقم الطالب
        sStudentName,                              // 2: اسم الطالب
        gradeName,                                 // 3: الصف
        sectionNum,                                // 4: الفصل
        sStudentPhone,                             // 5: رقم الجوال
        sItemText,                                 // 6: نوع الملاحظة
        detailsWithClass,                          // 7: التفاصيل (مع التصنيف)
        sTeacherName,                              // 8: المعلم/المسجل
        hijriDate,                                 // 9: التاريخ
        now,                                       // 10: وقت الإدخال (Date object - للفلتر)
        'لا'                                       // 11: تم الإرسال
      ];
      
    case 'positive':
      // ★ السلوك المتمايز - 13 عمود
      var posTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
      return [
        sStudentId,                                // 1: رقم الطالب
        sStudentName,                              // 2: اسم الطالب
        gradeName,                                 // 3: الصف
        sectionNum,                                // 4: الفصل
        sStudentPhone,                             // 5: رقم الجوال
        sItemText,                                 // 6: السلوك المتمايز
        sItemDegree,                               // 7: الدرجة
        sTeacherName,                              // 8: المعلم
        dayName,                                   // 9: اليوم
        hijriDate,                                 // 10: التاريخ الهجري
        now,                                       // 11: التاريخ الميلادي
        posTime,                                   // 12: وقت الإدخال (الوقت فقط)
        'لا'                                       // 13: تم الإرسال
      ];
      
    case 'custom':
      // ★ ملاحظة خاصة → تدخل في شيت الملاحظات التربوية - 11 عمود (مطابق للوكيل)
      return [
        sStudentId,                                // 1: رقم الطالب
        sStudentName,                              // 2: اسم الطالب
        gradeName,                                 // 3: الصف
        sectionNum,                                // 4: الفصل
        sStudentPhone,                             // 5: رقم الجوال
        'ملاحظة خاصة',                            // 6: نوع الملاحظة
        sanitizeInput_(formData.customNote || ''),     // 7: التفاصيل
        sTeacherName,                              // 8: المعلم/المسجل
        hijriDate,                                 // 9: التاريخ
        now,                                       // 10: وقت الإدخال (Date object - للفلتر)
        'لا'                                       // 11: تم الإرسال
      ];

    default:
      return [sStudentId, sStudentName, gradeName, sectionNum, sStudentPhone];
  }
}

// getHijriDate_() و getDayNameAr_() → مركزية في Config.gs

function logActivity_(formData, count) {
  try {
    var ss = getSpreadsheet_();
    var logSheet = ss.getSheetByName('سجل_النشاطات');
    
    if (!logSheet) {
      logSheet = ss.insertSheet('سجل_النشاطات');
      logSheet.setRightToLeft(true);
      logSheet.appendRow(['التاريخ', 'الوقت', 'المستخدم', 'النوع', 'التفاصيل', 'العدد', 'المرحلة']);
      logSheet.getRange(1, 1, 1, 7).setBackground('#1e3a5f').setFontColor('#ffffff');
    }
    
    var typeLabels = {
      'absence': 'غياب', 'violation': 'مخالفة سلوكية',
      'note': 'ملاحظة تربوية', 'positive': 'سلوك متمايز', 'custom': 'ملاحظة خاصة'
    };
    
    var stage = formData.stage || detectStage_(formData.className);
    
    logSheet.appendRow([
      new Date().toLocaleDateString('ar-SA'),
      new Date().toLocaleTimeString('ar-SA'),
      sanitizeInput_(formData.teacherName || ''),
      typeLabels[formData.inputType] || sanitizeInput_(formData.inputType || ''),
      'الفصل: ' + sanitizeInput_(formData.className || '') + (formData.itemText ? ' - ' + sanitizeInput_(formData.itemText) : ''),
      count,
      stage
    ]);
  } catch (e) {
    Logger.log('خطأ في تسجيل النشاط: ' + e.toString());
  }
}

// =================================================================
// دوال الروابط - جلب البيانات
// =================================================================
// ★ getLinksData() — حُذفت (كود قديم غير مستخدم)
// استُبدلت بـ getLinksTabData() الأحدث والأشمل

// إنشاء روابط لجميع المعلمين
function createAllTeachersLinks() {
  try {
    var ss = getSpreadsheet_();
    var teachersSheet = ss.getSheetByName("المعلمين");
    
    if (!teachersSheet) {
      return { success: false, error: 'شيت المعلمين غير موجود' };
    }
    
    var data = teachersSheet.getDataRange().getValues();
    var headers = data[0];
    
    var idCol = findColumnIndex(headers, ['المعرف', 'رقم المعلم', 'الرقم']);

    if (idCol === -1) return { success: false, error: 'لم يتم العثور على عمود المعرف' };

    var created = 0;

    for (var i = 1; i < data.length; i++) {
      var teacherId = data[i][idCol];
      if (teacherId) {
        var result = createTeacherLink(String(teacherId));
        if (result.success) {
          created++;
        }
      }
    }

    return { success: true, created: created };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// الحصول على رابط الخادم الأساسي
function getAppBaseUrl() {
  return ScriptApp.getService().getUrl().replace(/\/dev$/, '/exec');
}

// =================================================================
// دوال إنشاء روابط المستخدمين (الإداريين والحراس)
// =================================================================

function createUserLink(userId) {
  try {
    var ss = getSpreadsheet_();
    var usersSheet = ss.getSheetByName("المستخدمين");
    
    if (!usersSheet) {
      return { success: false, error: 'شيت المستخدمين غير موجود' };
    }
    
    var data = usersSheet.getDataRange().getValues();
    var headers = data[0];
    
    var idCol = findColumnIndex(headers, ['المعرف', 'رقم المستخدم', 'الرقم', 'ID']);
    var nameCol = findColumnIndex(headers, ['الاسم', 'اسم المستخدم']);
    var roleCol = findColumnIndex(headers, ['الدور', 'المسمى', 'الوظيفة']);
    var phoneCol = findColumnIndex(headers, ['الجوال', 'رقم الجوال', 'الهاتف']);
    var tokenCol = findColumnIndex(headers, ['رمز_الرابط', 'الرمز', 'Token']);
    var linkCol = findColumnIndex(headers, ['الرابط', 'رابط_النموذج', 'Link']);

    if (idCol === -1) return { success: false, error: 'لم يتم العثور على عمود المعرف' };

    if (tokenCol < 0) {
      tokenCol = headers.length;
      usersSheet.getRange(1, tokenCol + 1).setValue('رمز_الرابط');
    }
    if (linkCol < 0) {
      linkCol = headers.length + (tokenCol === headers.length ? 1 : 0);
      usersSheet.getRange(1, linkCol + 1).setValue('الرابط');
    }
    
    var userRow = -1;
    var userName = '';
    var userPhone = '';
    var userRole = '';
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(userId)) {
        userRow = i + 1;
        userName = String(data[i][nameCol] || '');
        userPhone = String(data[i][phoneCol] || '');
        userRole = String(data[i][roleCol] || '');
        break;
      }
    }
    
    if (userRow < 0) {
      return { success: false, error: 'المستخدم غير موجود' };
    }
    
    var token = generateTeacherToken();
    
    // ★ خريطة الدور → نوع الصفحة (5 واجهات مستقلة حسب الدور)
    var ROLE_PAGE_MAP = {
      'مدير المدرسة': 'wakeel',
      'وكيل شؤون الطلاب': 'wakeel',
      'وكيل الشؤون التعليمية': 'wakeel',
      'وكيل الشؤون المدرسية': 'wakeel',
      'موجه طلابي': 'counselor',
      'إداري': 'admin',
      'حارس': 'guard'
    };
    var pageType = ROLE_PAGE_MAP[userRole] || 'admin';
    
    var baseUrl = ScriptApp.getService().getUrl().replace(/\/dev$/, '/exec');
    var link = baseUrl + '?page=' + pageType + '&token=' + token;
    
    usersSheet.getRange(userRow, tokenCol + 1).setValue(token);
    usersSheet.getRange(userRow, linkCol + 1).setValue(link);
    
    return {
      success: true,
      link: link,
      token: token,
      userName: userName,
      userPhone: userPhone,
      role: userRole,
      pageType: pageType
    };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function createAllUsersLinks() {
  try {
    var ss = getSpreadsheet_();
    var usersSheet = ss.getSheetByName("المستخدمين");
    
    if (!usersSheet) {
      return { success: false, error: 'شيت المستخدمين غير موجود' };
    }
    
    var data = usersSheet.getDataRange().getValues();
    var headers = data[0];
    
    var idCol = findColumnIndex(headers, ['المعرف', 'رقم المستخدم', 'الرقم', 'ID']);

    if (idCol === -1) return { success: false, error: 'لم يتم العثور على عمود المعرف' };

    var created = 0;

    for (var i = 1; i < data.length; i++) {
      var userId = data[i][idCol];
      if (userId) {
        var result = createUserLink(String(userId));
        if (result.success) {
          created++;
        }
      }
    }

    return { success: true, created: created };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ★ getStaffFormData — حُذفت (مكررة في Server_StaffInput.gs)
// ★ getGuardFormData — حُذفت (كود ميت — Main.gs يستخدم getStaffByToken)

// جلب الفصول المتاحة (من شيتات المراحل)
function getAvailableClasses() {
  try {
    var sheets = getAllStudentsSheets_();
    var classes = [];
    
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s].sheet;
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;
      var headers = data[0];
      
      var gradeCol = findColumnIndex(headers, ['الصف', 'رقم الصف']);
      var classCol = findColumnIndex(headers, ['الفصل', 'فصل']);
      
      for (var i = 1; i < data.length; i++) {
        var grade = gradeCol > -1 ? cleanGradeName_(data[i][gradeCol]) : '';
        var cls = classCol > -1 ? String(data[i][classCol] || '').trim() : '';
        var fullClass = (grade + ' ' + cls).trim();
        if (fullClass && classes.indexOf(fullClass) === -1) {
          classes.push(fullClass);
        }
      }
    }
    
    return classes.sort();
    
  } catch (e) {
    return [];
  }
}

// =================================================================
// القسم 10: دوال تبويب الروابط المتقدمة (دعم المراحل والمعلم المشترك)
// =================================================================

/**
 * جلب بيانات تبويب الروابط (نسخة متقدمة مع دعم المراحل)
 * @returns {Object} بيانات المعلمين والإداريين والمربوطين
 */
function getLinksTabData() {
  try {
    var ss = getSpreadsheet_();
    
    // جلب المعلمين
    var teachersSheet = ss.getSheetByName('المعلمين');
    var teachers = [];
    if (teachersSheet) {
      var teachersData = teachersSheet.getDataRange().getValues();
      var teachersHeaders = teachersData[0];
      
      for (var i = 1; i < teachersData.length; i++) {
        var row = teachersData[i];
        if (row[0]) {
          var actualId = getColumnValue_(row, teachersHeaders, ['المعرف', 'رقم المعلم', 'الرقم']);
          var token = getColumnValue_(row, teachersHeaders, ['رمز_الرابط', 'الرمز', 'Token']);
          var link = getColumnValue_(row, teachersHeaders, ['الرابط', 'رابط_النموذج', 'Link']);
          
          teachers.push({
            id: String(actualId || row[0] || i),
            rowIndex: i,
            name: String(getColumnValue_(row, teachersHeaders, ['الاسم', 'اسم المعلم', 'name']) || ''),
            phone: String(getColumnValue_(row, teachersHeaders, ['الجوال', 'رقم الجوال', 'phone']) || ''),
            subject: String(getColumnValue_(row, teachersHeaders, ['المادة', 'subject']) || ''),
            classes: String(getColumnValue_(row, teachersHeaders, ['الفصول', 'الفصول المسندة', 'classes']) || ''),
            hasToken: !!token,
            hasLink: !!link,
            link: String(link || '')
          });
        }
      }
    }
    
    // جلب الإداريين (الكل بدون استثناء)
    var adminsSheet = ss.getSheetByName('المستخدمين') || ss.getSheetByName('الهيئة_الإدارية');
    var admins = [];
    if (adminsSheet) {
      var adminsData = adminsSheet.getDataRange().getValues();
      var adminsHeaders = adminsData[0];
      
      for (var i = 1; i < adminsData.length; i++) {
        var row = adminsData[i];
        if (row[0]) {
          var role = getColumnValue_(row, adminsHeaders, ['الدور', 'المسمى', 'role']);
          var name = getColumnValue_(row, adminsHeaders, ['الاسم', 'name']);
          
          // ★ إظهار الجميع: وكيل، موجه، مدير، إداري، حارس
          if (name) {
            var adminToken = getColumnValue_(row, adminsHeaders, ['رمز_الرابط', 'الرمز', 'Token']);
            var adminLink = getColumnValue_(row, adminsHeaders, ['الرابط', 'رابط_النموذج', 'Link']);
            var adminId = getColumnValue_(row, adminsHeaders, ['المعرف', 'ID']);
            
            admins.push({
              id: String(adminId || row[0] || i),
              name: String(name),
              phone: String(getColumnValue_(row, adminsHeaders, ['الجوال', 'رقم الجوال', 'phone']) || ''),
              role: String(role || ''),
              scope: String(getColumnValue_(row, adminsHeaders, ['النطاق', 'قيمة النطاق', 'نوع النطاق', 'scope']) || ''),
              classes: String(getColumnValue_(row, adminsHeaders, ['الفصول', 'classes', 'قيمة النطاق']) || ''),
              hasToken: !!adminToken,
              hasLink: !!adminLink,
              link: String(adminLink || '')
            });
          }
        }
      }
    }
    
    // جلب الأشخاص المربوطين من عدة مصادر
    var linkedPersons = [];
    
    // ★ المصدر 1: من التوكنات في شيت المعلمين نفسه
    for (var t = 0; t < teachers.length; t++) {
      if (teachers[t].hasToken) {
        linkedPersons.push({
          identifier: teachers[t].phone || teachers[t].id,
          type: 'teacher',
          linkedBy: '',
          stage: '',
          linkedDate: ''
        });
      }
    }
    
    // ★ المصدر 2: من التوكنات في شيت المستخدمين نفسه
    for (var a = 0; a < admins.length; a++) {
      if (admins[a].hasToken) {
        linkedPersons.push({
          identifier: admins[a].phone || String(admins[a].id),
          type: 'admin',
          linkedBy: '',
          stage: '',
          linkedDate: ''
        });
      }
    }
    
    // ★ المصدر 3: من شيت روابط_المعلمين (إن وجد)
    var linkedSheet = ss.getSheetByName('روابط_المعلمين') || ss.getSheetByName('الروابط');
    if (linkedSheet) {
      var linkedData = linkedSheet.getDataRange().getValues();
      var linkedHeaders = linkedData[0];
      
      for (var i = 1; i < linkedData.length; i++) {
        var row = linkedData[i];
        if (row[0]) {
          var identifier = String(getColumnValue_(row, linkedHeaders, ['الجوال', 'المعرف', 'phone', 'identifier']) || '');
          // تجنب التكرار
          var alreadyExists = linkedPersons.some(function(lp) { return lp.identifier === identifier; });
          if (!alreadyExists) {
            linkedPersons.push({
              identifier: identifier,
              type: String(getColumnValue_(row, linkedHeaders, ['النوع', 'type']) || 'teacher'),
              linkedBy: String(getColumnValue_(row, linkedHeaders, ['تم الربط بواسطة', 'linkedBy']) || ''),
              stage: String(getColumnValue_(row, linkedHeaders, ['المرحلة', 'stage']) || ''),
              linkedDate: String(getColumnValue_(row, linkedHeaders, ['تاريخ الربط', 'linkedDate']) || '')
            });
          }
        }
      }
    }
    
    // ★ المراحل المتاحة (من Config.gs — المرجع الوحيد)
    ensureStudentsSheetsLoaded_();
    var availableStages = Object.keys(STUDENTS_SHEETS);

    return {
      success: true,
      teachers: teachers,
      admins: admins,
      linkedPersons: linkedPersons,
      availableStages: availableStages
    };
    
  } catch (error) {
    Logger.log('Error in getLinksTabData: ' + error.message);
    return {
      success: false,
      message: error.message,
      teachers: [],
      admins: [],
      linkedPersons: []
    };
  }
}

/**
 * إرسال رابط لشخص مع تحديد المرحلة
 * @param {Object} data - بيانات الشخص والمرحلة
 * @returns {Object} نتيجة الإرسال
 */
function sendLinkToPersonWithStage(data) {
  try {
    var name = data.name;
    var phone = formatPhone(data.phone);
    var type = data.type;
    var stage = data.stage;
    var classes = data.classes;
    
    if (!phone) {
      return { success: false, message: 'رقم الجوال غير صحيح' };
    }
    
    // ★ الخطوة 1: إنشاء الرابط وحفظ التوكن في الشيت (الأهم!)
    var linkResult;
    if (type === 'teacher') {
      // البحث عن المعلم بالجوال أو المعرف
      linkResult = createTeacherLinkByPhone_(phone, name);
    } else {
      // البحث عن الإداري بالجوال أو المعرف
      linkResult = createUserLinkByPhone_(phone, name);
    }
    
    if (!linkResult || !linkResult.success) {
      return { success: false, message: (linkResult && linkResult.error) || 'فشل إنشاء الرابط' };
    }
    
    var link = linkResult.link;
    
    // ★ الخطوة 2: محاولة إرسال عبر واتساب (اختياري - لا يمنع النجاح)
    var whatsappSent = false;
    var whatsappMessage = '';
    
    try {
      var wakilPhone = getWakilPhoneByStage_(stage);
      if (wakilPhone) {
        var message = buildLinkMessage_(name, type, link);
        var sendResult = sendWhatsAppFromStage_(phone, message, stage, wakilPhone);
        whatsappSent = sendResult.success;
        if (!whatsappSent) {
          whatsappMessage = 'تم إنشاء الرابط لكن فشل إرسال الواتساب. يمكنك نسخ الرابط وإرساله يدوياً.';
        }
      } else {
        whatsappMessage = 'تم إنشاء الرابط بنجاح. لا يوجد واتساب متصل لمرحلة ' + stage + ' - يمكنك نسخ الرابط وإرساله يدوياً.';
      }
    } catch(whatsErr) {
      Logger.log('WhatsApp send error (non-blocking): ' + whatsErr.message);
      whatsappMessage = 'تم إنشاء الرابط لكن فشل إرسال الواتساب.';
    }
    
    // ★ الخطوة 3: تسجيل الربط
    saveLinkedPerson_({
      phone: phone,
      name: name,
      type: type,
      stage: stage,
      classes: classes,
      linkedBy: stage
    });
    
    // ★ الخطوة 4: إشعار الوكلاء الآخرين (للمعلم المشترك)
    var otherStages = getOtherStagesForPerson_(classes, stage);
    var notifyResult = { notifyOtherWakil: false };
    
    if (otherStages.length > 0) {
      try {
        notifyOtherWakils_(name, phone, type, stage, otherStages);
      } catch(e) {}
      notifyResult = {
        notifyOtherWakil: true,
        otherStage: otherStages.join(' - ')
      };
    }
    
    return {
      success: true,
      message: whatsappSent ? 'تم إرسال الرابط عبر واتساب بنجاح ✅' : whatsappMessage,
      link: link,
      whatsappSent: whatsappSent,
      notifyOtherWakil: notifyResult.notifyOtherWakil,
      otherStage: notifyResult.otherStage
    };
    
  } catch (error) {
    Logger.log('Error in sendLinkToPersonWithStage: ' + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * إرسال جماعي للروابط
 * @param {Array} persons - قائمة الأشخاص للإرسال
 * @returns {Object} نتيجة الإرسال
 */
function bulkSendLinks(persons) {
  try {
    var sentCount = 0;
    var failedCount = 0;
    var errors = [];
    
    for (var i = 0; i < persons.length; i++) {
      var person = persons[i];
      
      if (i > 0) {
        Utilities.sleep(2000);
      }
      
      var result = sendLinkToPersonWithStage(person);
      
      if (result.success) {
        sentCount++;
      } else {
        failedCount++;
        errors.push(person.name + ': ' + result.message);
      }
    }
    
    return {
      success: true,
      sentCount: sentCount,
      failedCount: failedCount,
      errors: errors
    };
    
  } catch (error) {
    Logger.log('Error in bulkSendLinks: ' + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * إلغاء ربط شخص
 * @param {Object} personData - بيانات الشخص
 * @returns {Object} نتيجة العملية
 */
function removeLinkForPerson(personData) {
  try {
    var phone = formatPhone(personData.phone);
    var name = personData.name || '';
    var type = personData.type;
    
    var ss = getSpreadsheet_();
    var removed = false;
    var oldToken = ''; // ★ نحفظ التوكن القديم لمسح الكاش
    
    // ★ 1. حذف من شيت الروابط
    var linkedSheet = ss.getSheetByName('روابط_المعلمين') || ss.getSheetByName('الروابط');
    if (linkedSheet) {
      var sheetData = linkedSheet.getDataRange().getValues();
      var headers = sheetData[0];
      var phoneCol = findColumnIndex(headers, ['الجوال', 'المعرف', 'phone', 'identifier']);
      var typeCol = findColumnIndex(headers, ['النوع', 'type']);
      
      for (var i = sheetData.length - 1; i >= 1; i--) {
        var rowPhone = formatPhone(sheetData[i][phoneCol]);
        var rowType = sheetData[i][typeCol] || 'teacher';
        
        if (rowPhone === phone && rowType === type) {
          linkedSheet.deleteRow(i + 1);
          removed = true;
        }
      }
    }
    
    // ★ 2. مسح رمز_الرابط والرابط من شيت المصدر (المعلمين أو المستخدمين)
    var sourceSheetName = (type === 'teacher') ? 'المعلمين' : 'المستخدمين';
    var sourceSheet = ss.getSheetByName(sourceSheetName);
    
    if (sourceSheet && sourceSheet.getLastRow() > 1) {
      var srcData = sourceSheet.getDataRange().getValues();
      var srcHeaders = srcData[0];
      var srcPhoneCol = findColumnIndex(srcHeaders, ['الجوال', 'رقم الجوال', 'phone']);
      var srcNameCol = findColumnIndex(srcHeaders, ['الاسم', 'name']);
      var srcTokenCol = findColumnIndex(srcHeaders, ['رمز_الرابط', 'الرمز', 'Token']);
      var srcLinkCol = findColumnIndex(srcHeaders, ['الرابط', 'رابط_النموذج', 'Link']);
      
      // ★ بحث آمن: الجوال أولاً، الاسم فقط إذا فريد
      var targetRow = -1;
      var nameMatchRow = -1;
      var nameMatchCount = 0;
      
      for (var j = 1; j < srcData.length; j++) {
        var rowPhone2 = formatPhone(srcData[j][srcPhoneCol]);
        var rowName2 = String(srcData[j][srcNameCol] || '').trim();
        
        if (rowPhone2 && rowPhone2 === phone) {
          targetRow = j;
          break;
        }
        if (name && rowName2 === name) {
          nameMatchRow = j;
          nameMatchCount++;
        }
      }
      
      if (targetRow === -1 && nameMatchCount === 1) {
        targetRow = nameMatchRow;
      }
      
      if (targetRow > 0) {
        // ★ قراءة التوكن القديم قبل المسح (لمسح الكاش)
        if (srcTokenCol !== -1) {
          oldToken = String(srcData[targetRow][srcTokenCol] || '').trim();
          sourceSheet.getRange(targetRow + 1, srcTokenCol + 1).setValue('');
        }
        if (srcLinkCol !== -1) {
          sourceSheet.getRange(targetRow + 1, srcLinkCol + 1).setValue('');
        }
        removed = true;
      }
    }
    
    // ★ 3. مسح الكاش — الرابط يتوقف فوراً بدل الانتظار 6 ساعات
    if (oldToken) {
      try {
        CacheService.getScriptCache().remove('tpd_' + oldToken);
      } catch(cacheErr) {
        Logger.log('Cache clear warning: ' + cacheErr.message);
      }
    }
    
    SpreadsheetApp.flush();
    
    if (removed) {
      return { success: true, message: 'تم إلغاء الربط بنجاح' };
    } else {
      return { success: false, message: 'لم يتم العثور على الشخص' };
    }
    
  } catch (error) {
    Logger.log('Error in removeLinkForPerson: ' + error.message);
    return { success: false, message: error.message };
  }
}

// =================================================================
// القسم 11: الدوال المساعدة للروابط المتقدمة
// =================================================================

/**
 * الحصول على رقم وكيل المرحلة
 */
function getWakilPhoneByStage_(stage) {
  try {
    // ★ استخدام الرقم الرئيسي مباشرة
    if (typeof getPrimaryPhoneForStage === 'function') {
      var primary = getPrimaryPhoneForStage(stage);
      if (primary.success) {
        return formatPhone(primary.phone);
      }
    }
    
    // ★ بديل: أي رقم متصل
    if (typeof getConnectedSessionsByStage === 'function') {
      var result = getConnectedSessionsByStage(stage);
      if (result.success && result.sessions && result.sessions.length > 0) {
        return formatPhone(result.sessions[0].phone);
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log('Error in getWakilPhoneByStage_: ' + error.message);
    return null;
  }
}

// ★ generateLinkUrl_() — حُذفت (كود قديم غير مستخدم)
// الروابط تُنشأ فعلياً في createTeacherLink() و createUserLink()

/**
 * بناء رسالة الرابط
 */
function buildLinkMessage_(name, type, link) {
  var typeText = type === 'teacher' ? 'المعلم' : 'الإداري';
  var schoolName = getSchoolNameForLinks_();
  
  var message = '🎓 *' + schoolName + '*\n\n';
  message += 'السلام عليكم ورحمة الله وبركاته\n\n';
  message += 'الأستاذ الفاضل/ ' + name + '\n\n';
  message += 'تم إعداد رابط خاص بك للدخول إلى نظام المتابعة:\n\n';
  message += '🔗 ' + link + '\n\n';
  message += '📱 يمكنك استخدام هذا الرابط من جوالك أو جهاز الكمبيوتر.\n\n';
  message += 'شكراً لتعاونكم 🙏';
  
  return message;
}

/**
 * إرسال واتساب من رقم المرحلة المحددة
 */
function sendWhatsAppFromStage_(to, message, stage, fromPhone) {
  try {
    // ★ استخدام الدالة الصحيحة من Server_WhatsApp.gs
    if (typeof sendWhatsAppMessageFrom === 'function') {
      return sendWhatsAppMessageFrom(fromPhone, to, message);
    }
    
    // بديل: استدعاء مباشر بالـ API الصحيح
    var serverUrl = getWhatsAppServerUrl_();
    var cleanSender = formatPhone(fromPhone);
    var cleanRecipient = formatPhone(to);
    
    var payload = {
      phone: cleanRecipient,
      message: message
    };
    
    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(serverUrl + '/send/' + cleanSender, options);
    var result = JSON.parse(response.getContentText());
    
    return {
      success: result.success || false,
      message: result.message || ''
    };
    
  } catch (error) {
    Logger.log('Error in sendWhatsAppFromStage_: ' + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * حفظ بيانات الشخص المربوط
 */
function saveLinkedPerson_(data) {
  try {
    var ss = getSpreadsheet_();
    var linkedSheet = ss.getSheetByName('روابط_المعلمين');
    
    if (!linkedSheet) {
      linkedSheet = ss.insertSheet('روابط_المعلمين');
      linkedSheet.appendRow([
        'الجوال', 'الاسم', 'النوع', 'المرحلة', 'الفصول', 
        'تم الربط بواسطة', 'تاريخ الربط'
      ]);
    }
    
    // ★ فحص التكرار: إذا نفس الجوال ونفس النوع موجود — حدّثه بدل ما تكرره
    var sheetData = linkedSheet.getDataRange().getValues();
    var phoneToFind = formatPhone(data.phone);
    var existingRow = -1;
    
    for (var i = 1; i < sheetData.length; i++) {
      var rowPhone = formatPhone(String(sheetData[i][0] || ''));
      var rowType = String(sheetData[i][2] || '');
      if (rowPhone === phoneToFind && rowType === data.type) {
        existingRow = i + 1;
        break;
      }
    }
    
    var rowData = [
      data.phone,
      data.name,
      data.type,
      data.stage,
      data.classes,
      data.linkedBy,
      new Date().toISOString()
    ];
    
    if (existingRow > 0) {
      // ★ تحديث الصف الموجود
      linkedSheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      // إضافة صف جديد
      linkedSheet.appendRow(rowData);
    }
    
  } catch (error) {
    Logger.log('Error in saveLinkedPerson_: ' + error.message);
  }
}

/**
 * الحصول على المراحل الأخرى للشخص
 */
function getOtherStagesForPerson_(classes, currentStage) {
  var allStages = getStagesFromClasses_(classes);
  return allStages.filter(function(stage) {
    return stage !== currentStage;
  });
}

/**
 * استخراج المراحل من الفصول
 */
function getStagesFromClasses_(classes) {
  var stages = [];
  
  if (!classes) return stages;
  
  var classesArray = classes.toString().split(',');
  
  classesArray.forEach(function(cls) {
    cls = cls.trim();
    var stage = null;
    
    // ★ دعم الصيغ العربية والإنجليزية (classKey = الأول_intermediate_أ)
    if (cls.match(/ابتدائي|primary/i) || cls.match(/[1-6]\s*-?\s*ب/)) {
      stage = 'ابتدائي';
    } else if (cls.match(/متوسط|intermediate/i) || cls.match(/[1-3]\s*-?\s*م/)) {
      stage = 'متوسط';
    } else if (cls.match(/ثانوي|secondary/i) || cls.match(/[1-3]\s*-?\s*ث/)) {
      stage = 'ثانوي';
    }
    
    if (stage && stages.indexOf(stage) === -1) {
      stages.push(stage);
    }
  });
  
  return stages;
}

/**
 * إشعار الوكلاء الآخرين بالربط
 */
function notifyOtherWakils_(name, phone, type, linkedStage, otherStages) {
  try {
    otherStages.forEach(function(stage) {
      var wakilPhone = getWakilPhoneByStage_(stage);
      if (wakilPhone) {
        var typeText = type === 'teacher' ? 'المعلم' : 'الإداري';
        var message = '📢 *إشعار ربط*\n\n';
        message += 'تم ربط ' + typeText + ' *' + name + '* بالنظام\n';
        message += 'من قبل وكيل مرحلة *' + linkedStage + '*\n\n';
        message += 'ملاحظة: هذا الشخص مشترك في مرحلتك أيضاً، والردود ستصلك بحسب الفصل المختار.';
        
        sendWhatsAppFromStage_(wakilPhone, message, stage, wakilPhone);
      }
    });
  } catch (error) {
    Logger.log('Error in notifyOtherWakils_: ' + error.message);
  }
}

/**
 * الحصول على رابط سيرفر الواتساب
 */
function getWhatsAppServerUrl_() {
  try {
    var ss = getSpreadsheet_();
    var settingsSheet = ss.getSheetByName('إعدادات_المدرسة');
    if (settingsSheet) {
      var data = settingsSheet.getDataRange().getValues();
      for (var i = 0; i < data.length; i++) {
        if (data[i][0] === 'whatsapp_server_url' || data[i][0] === 'رابط_سيرفر_الواتساب') {
          return data[i][1];
        }
      }
    }
  } catch (e) {}
  
  return 'http://194.163.133.252:3000';
}

/**
 * الحصول على قيمة عمود معين من صف
 */
function getColumnValue_(row, headers, possibleNames) {
  for (var i = 0; i < possibleNames.length; i++) {
    var index = headers.indexOf(possibleNames[i]);
    if (index !== -1) {
      return row[index];
    }
  }
  return '';
}

/**
 * الحصول على إعدادات المدرسة
 */
// ★ دالة مخصصة لجلب اسم المدرسة فقط (للروابط والرسائل)
// getSchoolSettings_() الرئيسية في Server_Data.gs — لا تكررها هنا
function getSchoolNameForLinks_() {
  try {
    var ss = getSpreadsheet_();
    var settingsSheet = ss.getSheetByName('إعدادات_المدرسة');
    
    if (!settingsSheet) {
      return 'المدرسة';
    }
    
    var data = settingsSheet.getDataRange().getValues();
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === 'school_name' || data[i][0] === 'اسم_المدرسة' || data[i][0] === 'schoolName') {
        return String(data[i][1] || 'المدرسة');
      }
    }
    
    return 'المدرسة';
    
  } catch (error) {
    return 'المدرسة';
  }
}

// ★ دوال الاختبار حُذفت (كانت testTeacherToken و testGetTeacher — كود ميت)
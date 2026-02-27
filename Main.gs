// =================================================================
// HTML SERVICE - خدمة عرض الصفحات
// =================================================================
function doGet(e) {
  var page = e.parameter.page || 'main';
  var token = e.parameter.token || '';
  
  // نموذج المعلم (غياب، مخالفات، ملاحظات)
  if (page === 'teacher' && token) {
    var template = HtmlService.createTemplateFromFile('TeacherInputForm');
    try {
      var data = getTeacherPageData_(token);
      template.pageData = JSON.stringify(data || {success:false, error:'الدالة أرجعت null'});
    } catch(err) {
      template.pageData = JSON.stringify({success:false, error:'خطأ في تحميل البيانات: ' + err.toString()});
    }
    return template.evaluate()
      .setTitle('نموذج إدخال المعلم')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  }
  
  // نموذج الهيئة الإدارية (تأخر، استئذان)
  if (page === 'staff' && token) {
    var template = HtmlService.createTemplateFromFile('StaffInputForm');
    try {
      var data = getStaffByToken(token);
      data.token = token;
      template.pageData = JSON.stringify(data || {success:false, error:'الدالة أرجعت null'});
      // ★ حقن بيانات المراحل والصفوف من هيكل_المدرسة
      var stagesConfig = getEnabledStagesConfig_();
      template.gradeMap = JSON.stringify(stagesConfig.stages || {});
    } catch(err) {
      template.pageData = JSON.stringify({success:false, error:'خطأ في تحميل البيانات: ' + err.toString()});
      template.gradeMap = JSON.stringify({});
    }
    return template.evaluate()
      .setTitle('نموذج التأخر والاستئذان')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  }
  
  // واجهة ولي الأمر (تقديم عذر غياب)
  if (page === 'parent' && token) {
    var template = HtmlService.createTemplateFromFile('ParentExcuseForm');
    try {
      var data = getParentExcusePageData_(token);
      if (data && data.success) data.token = token;
      template.pageData = JSON.stringify(data || {success:false, error:'الدالة أرجعت null'});
    } catch(err) {
      template.pageData = JSON.stringify({success:false, error:'خطأ في تحميل البيانات: ' + err.toString()});
    }
    return template.evaluate()
      .setTitle('تقديم عذر غياب - ولي الأمر')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  }

  // واجهة الحارس (عرض المستأذنين)
  if (page === 'guard' && token) {
    var template = HtmlService.createTemplateFromFile('GuardDisplay');
    try {
      var data = getStaffByToken(token);
      // ★ إصلاح: التحقق من أن المستخدم فعلاً حارس
      if (data && data.success && data.staff && !data.staff.isGuard) {
        data = { success: false, error: 'هذا الرابط مخصص للحراس فقط. دورك: ' + (data.staff.role || 'غير محدد') };
      }
      if (data && data.success) {
        data.token = token;
        // ★ حقن المراحل المفعّلة لبناء التبويبات ديناميكياً
        ensureStudentsSheetsLoaded_();
        data.enabledStages = Object.keys(STUDENTS_SHEETS);
      }
      template.pageData = JSON.stringify(data || {success:false, error:'الدالة أرجعت null'});
    } catch(err) {
      template.pageData = JSON.stringify({success:false, error:'خطأ في تحميل البيانات: ' + err.toString()});
    }
    return template.evaluate()
      .setTitle('سجل المستأذنين - الحارس')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  }
  
  // ★ واجهة الوكيل (جميع الميزات: مخالفات + غياب + سلوك + ملاحظات + استئذان + تأخر)
  if (page === 'wakeel' && token) {
    var template = HtmlService.createTemplateFromFile('WakeelForm');
    try {
      var data = getWakeelPageData_(token);
      data.token = token;
      template.pageData = JSON.stringify(data || {success:false, error:'الدالة أرجعت null'});
      var stagesConfig = getEnabledStagesConfig_();
      template.gradeMap = JSON.stringify(stagesConfig.stages || {});
    } catch(err) {
      template.pageData = JSON.stringify({success:false, error:'خطأ في تحميل البيانات: ' + err.toString()});
      template.gradeMap = JSON.stringify({});
    }
    return template.evaluate()
      .setTitle('نموذج الوكيل - النظام الشامل')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  }

  // ★ واجهة الموجه الطلابي (استئذان + ملاحظات تربوية + سلوك متمايز)
  if (page === 'counselor' && token) {
    var template = HtmlService.createTemplateFromFile('CounselorForm');
    try {
      var data = getCounselorPageData_(token);
      data.token = token;
      template.pageData = JSON.stringify(data || {success:false, error:'الدالة أرجعت null'});
      var stagesConfig = getEnabledStagesConfig_();
      template.gradeMap = JSON.stringify(stagesConfig.stages || {});
    } catch(err) {
      template.pageData = JSON.stringify({success:false, error:'خطأ في تحميل البيانات: ' + err.toString()});
      template.gradeMap = JSON.stringify({});
    }
    return template.evaluate()
      .setTitle('نموذج الموجه الطلابي')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  }

  // ★ واجهة الإداري (تسجيل التأخر الصباحي فقط)
  if (page === 'admin' && token) {
    var template = HtmlService.createTemplateFromFile('AdminTardinessForm');
    try {
      var data = getAdminPageData_(token);
      data.token = token;
      template.pageData = JSON.stringify(data || {success:false, error:'الدالة أرجعت null'});
      var stagesConfig = getEnabledStagesConfig_();
      template.gradeMap = JSON.stringify(stagesConfig.stages || {});
    } catch(err) {
      template.pageData = JSON.stringify({success:false, error:'خطأ في تحميل البيانات: ' + err.toString()});
      template.gradeMap = JSON.stringify({});
    }
    return template.evaluate()
      .setTitle('تسجيل التأخر الصباحي')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  }

  // ★ إضافة "رصد الغياب" — ترجع JSON بيانات الغياب اليومي
  if (page === 'extension') {
    var extensionData = getExtensionAbsenceData();
    return ContentService.createTextOutput(JSON.stringify(extensionData))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // الصفحة الرئيسية (لوحة التحكم)
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('نظام المخالفات السلوكية الشامل')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// دالة الربط
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}
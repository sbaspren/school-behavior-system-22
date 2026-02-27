// =================================================================
// نظام واتساب المتكامل - Server_WhatsApp.gs
// النسخة المحدثة: دعم المرحلة + نوع المستخدم
// =================================================================

// ★ رابط السيرفر يُقرأ من الإعدادات أو Script Properties — لا يُخزّن هنا مباشرة
function getWhatsAppServerUrl_() {
  // أولاً: حاول من إعدادات الواتساب في الشيت
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('إعدادات_واتساب');
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === 'رابط_السيرفر') {
          var url = String(data[i][1] || '').trim();
          if (url) return url;
        }
      }
    }
  } catch (e) {}
  // ثانياً: من Script Properties
  var prop = PropertiesService.getScriptProperties().getProperty('whatsapp_server_url');
  if (prop) return prop;
  // افتراضي (للتوافقية فقط — يجب تعيينه في الإعدادات)
  return '';
}
var WHATSAPP_SERVER_URL = getWhatsAppServerUrl_();

// أنواع المستخدمين المتاحة
var USER_TYPES = ['وكيل', 'مدير', 'موجه'];

// =================================================================
// 1. إنشاء الأوراق تلقائياً (الهيكل الجديد)
// =================================================================

/**
 * الحصول على أو إنشاء ورقة إعدادات الواتساب
 */
function getWhatsAppSettingsSheet() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('إعدادات_واتساب');
  
  if (!sheet) {
    sheet = ss.insertSheet('إعدادات_واتساب');
    
    // إنشاء الهيكل
    var data = [
      ['الإعداد', 'القيمة', 'الوصف'],
      ['رمز_الأمان', '', 'رمز من 6 خانات للحماية'],
      ['جوال_الاسترجاع_1', '', 'رقم الجوال الأول لاسترجاع الرمز (إجباري)'],
      ['جوال_الاسترجاع_2', '', 'رقم الجوال الثاني لاسترجاع الرمز (اختياري)'],
      ['رابط_السيرفر', WHATSAPP_SERVER_URL, 'رابط سيرفر الواتساب'],
      ['حالة_الخدمة', 'مفعل', 'مفعل أو معطل'],
      ['رمز_الاسترجاع_المؤقت', '', 'يُستخدم داخلياً - لا تعدله'],
      ['وقت_انتهاء_الاسترجاع', '', 'يُستخدم داخلياً - لا تعدله']
    ];
    
    sheet.getRange(1, 1, data.length, 3).setValues(data);
    
    // تنسيق الترويسة
    sheet.getRange(1, 1, 1, 3)
      .setBackground('#25D366')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    
    // ضبط عرض الأعمدة
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 300);
    
    // إخفاء الصفوف الداخلية
    sheet.hideRows(7, 2);
  }
  
  return sheet;
}

/**
 * الحصول على أو إنشاء ورقة جلسات الواتساب (الهيكل الجديد)
 */
function getWhatsAppSessionsSheet() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('جلسات_واتساب');
  
  if (!sheet) {
    sheet = ss.insertSheet('جلسات_واتساب');
    
    // الهيكل الجديد مع المرحلة ونوع المستخدم
    var headers = [
      'رقم_الواتساب',
      'المرحلة',
      'نوع_المستخدم',
      'حالة_الاتصال',
      'تاريخ_الربط',
      'آخر_استخدام',
      'عدد_الرسائل',
      'الرقم_الرئيسي'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#25D366')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    sheet.setFrozenRows(1);
    
    // ضبط عرض الأعمدة
    sheet.setColumnWidth(1, 140); // رقم_الواتساب
    sheet.setColumnWidth(2, 80);  // المرحلة
    sheet.setColumnWidth(3, 100); // نوع_المستخدم
    sheet.setColumnWidth(4, 100); // حالة_الاتصال
    sheet.setColumnWidth(5, 120); // تاريخ_الربط
    sheet.setColumnWidth(6, 120); // آخر_استخدام
    sheet.setColumnWidth(7, 100); // عدد_الرسائل
  }
  
  return sheet;
}

/**
 * إعادة بناء شيت الجلسات (حذف القديم وإنشاء جديد)
 */
function rebuildSessionsSheet() {
  try {
    var ss = getSpreadsheet_();
    var oldSheet = ss.getSheetByName('جلسات_واتساب');
    
    // حذف الشيت القديم إن وجد
    if (oldSheet) {
      ss.deleteSheet(oldSheet);
    }
    
    // إنشاء شيت جديد
    var newSheet = getWhatsAppSessionsSheet();
    
    return { success: true, message: 'تم إعادة بناء شيت الجلسات بنجاح' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 2. دوال الإعدادات
// =================================================================

/**
 * جلب إعداد معين
 */
function getWhatsAppSetting(settingName) {
  try {
    var sheet = getWhatsAppSettingsSheet();
    var data = sheet.getDataRange().getValues();
    
    var cleanSettingName = String(settingName).trim();
    
    for (var i = 1; i < data.length; i++) {
      var cellValue = String(data[i][0]).trim();
      
      if (cellValue === cleanSettingName) {
        return data[i][1];
      }
    }
    return null;
  } catch (e) {
    console.error('خطأ في getWhatsAppSetting:', e);
    return null;
  }
}

/**
 * حفظ إعداد معين
 */
function setWhatsAppSetting(settingName, value) {
  var sheet = getWhatsAppSettingsSheet();
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === settingName) {
      sheet.getRange(i + 1, 2).setValue(value);
      return true;
    }
  }
  return false;
}

/**
 * التحقق من وجود رمز أمان مُعد
 */
function isSecurityCodeConfigured() {
  var code = getWhatsAppSetting('رمز_الأمان');
  var phone1 = getWhatsAppSetting('جوال_الاسترجاع_1');
  
  var codeStr = code ? String(code).trim() : '';
  var phoneStr = phone1 ? String(phone1).trim() : '';
  
  return codeStr.length >= 6 && phoneStr.length >= 10;
}

// =================================================================
// 3. دوال رمز الأمان
// =================================================================

/**
 * إعداد رمز الأمان لأول مرة
 */
function setupSecurityCode(code, phone1, phone2) {
  try {
    if (!code || String(code).length < 6) {
      return { success: false, error: 'رمز الأمان يجب أن يكون 6 خانات على الأقل' };
    }
    
    if (!phone1 || String(phone1).length < 10) {
      return { success: false, error: 'رقم الجوال الأول إجباري' };
    }
    
    var cleanedPhone1 = cleanPhoneNumber(phone1);
    var cleanedPhone2 = phone2 ? cleanPhoneNumber(phone2) : '';
    
    setWhatsAppSetting('رمز_الأمان', String(code).trim());
    setWhatsAppSetting('جوال_الاسترجاع_1', cleanedPhone1);
    setWhatsAppSetting('جوال_الاسترجاع_2', cleanedPhone2);
    
    return { success: true, message: 'تم إعداد رمز الأمان بنجاح' };
  } catch (e) {
    return { success: false, error: 'خطأ: ' + e.toString() };
  }
}

/**
 * التحقق من رمز الأمان
 */
function verifySecurityCode(inputCode) {
  try {
    var savedCode = getWhatsAppSetting('رمز_الأمان');
    
    if (!savedCode || savedCode === '') {
      return { success: false, error: 'لم يتم إعداد رمز الأمان بعد', needSetup: true };
    }
    
    var inputStr = String(inputCode).trim();
    var savedStr = String(savedCode).trim();
    
    if (inputStr === savedStr) {
      return { success: true };
    }
    
    return { success: false, error: 'رمز الأمان غير صحيح' };
  } catch (e) {
    return { success: false, error: 'خطأ في التحقق: ' + e.toString() };
  }
}

/**
 * إرسال رمز استرجاع للجوال
 */
function sendRecoveryCode(phoneChoice) {
  try {
    var phone1 = getWhatsAppSetting('جوال_الاسترجاع_1');
    var phone2 = getWhatsAppSetting('جوال_الاسترجاع_2');
    
    var targetPhone = '';
    if (phoneChoice === 1 || !phone2) {
      targetPhone = phone1;
    } else if (phoneChoice === 2 && phone2) {
      targetPhone = phone2;
    } else {
      return { success: false, error: 'رقم غير صحيح' };
    }
    
    if (!targetPhone) {
      return { success: false, error: 'لا يوجد رقم جوال للاسترجاع' };
    }
    
    var recoveryCode = Math.floor(1000 + Math.random() * 9000).toString();
    
    var expiryTime = new Date(Date.now() + 5 * 60 * 1000).toISOString();
    setWhatsAppSetting('رمز_الاسترجاع_المؤقت', recoveryCode);
    setWhatsAppSetting('وقت_انتهاء_الاسترجاع', expiryTime);
    
    var message = '🔐 رمز استرجاع رمز الأمان الخاص بنظام التوجيه الطلابي:\n\n' + recoveryCode + '\n\nصالح لمدة 5 دقائق فقط.';
    
    // استخدام أي رقم متصل للإرسال
    var sessions = getAllConnectedSessions();
    if (!sessions.success || sessions.sessions.length === 0) {
      return { success: false, error: 'لا يوجد رقم واتساب متصل لإرسال رمز الاسترجاع' };
    }
    
    var senderPhone = sessions.sessions[0].phone;
    var sendResult = sendWhatsAppMessageFrom(senderPhone, targetPhone, message);
    
    if (sendResult.success) {
      var maskedPhone = targetPhone.substring(0, 6) + '****' + targetPhone.substring(targetPhone.length - 2);
      return { success: true, message: 'تم إرسال رمز الاسترجاع إلى ' + maskedPhone };
    } else {
      return { success: false, error: 'فشل في إرسال الرمز: ' + sendResult.error };
    }
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * التحقق من رمز الاسترجاع
 */
function verifyRecoveryCode(inputCode) {
  try {
    var savedCode = getWhatsAppSetting('رمز_الاسترجاع_المؤقت');
    var expiryTime = getWhatsAppSetting('وقت_انتهاء_الاسترجاع');
    
    if (!savedCode || savedCode === '' || !expiryTime) {
      return { success: false, error: 'لم يتم طلب رمز استرجاع' };
    }
    
    if (new Date() > new Date(expiryTime)) {
      setWhatsAppSetting('رمز_الاسترجاع_المؤقت', '');
      setWhatsAppSetting('وقت_انتهاء_الاسترجاع', '');
      return { success: false, error: 'انتهت صلاحية رمز الاسترجاع' };
    }
    
    var inputStr = String(inputCode).trim();
    var savedStr = String(savedCode).trim();
    
    if (inputStr === savedStr) {
      setWhatsAppSetting('رمز_الاسترجاع_المؤقت', '');
      setWhatsAppSetting('وقت_انتهاء_الاسترجاع', '');
      return { success: true };
    }
    
    return { success: false, error: 'رمز الاسترجاع غير صحيح' };
  } catch (e) {
    return { success: false, error: 'خطأ: ' + e.toString() };
  }
}

/**
 * تغيير رمز الأمان
 */
function changeSecurityCode(newCode) {
  try {
    if (!newCode || newCode.length < 6) {
      return { success: false, error: 'رمز الأمان يجب أن يكون 6 خانات على الأقل' };
    }
    
    setWhatsAppSetting('رمز_الأمان', newCode);
    return { success: true, message: 'تم تغيير رمز الأمان بنجاح' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * جلب أرقام الاسترجاع (مخفية جزئياً)
 */
function getRecoveryPhones() {
  try {
    var phone1 = getWhatsAppSetting('جوال_الاسترجاع_1');
    var phone2 = getWhatsAppSetting('جوال_الاسترجاع_2');
    
    var maskPhone = function(phone) {
      if (!phone || String(phone).length < 8) return null;
      var phoneStr = String(phone);
      return phoneStr.substring(0, 6) + '****' + phoneStr.substring(phoneStr.length - 2);
    };
    
    return {
      success: true,
      phone1: maskPhone(phone1),
      phone2: maskPhone(phone2),
      hasPhone1: !!phone1 && String(phone1).length >= 10,
      hasPhone2: !!phone2 && String(phone2).length >= 10
    };
  } catch (e) {
    return { 
      success: false, 
      error: e.toString(),
      hasPhone1: false, 
      hasPhone2: false 
    };
  }
}

// =================================================================
// 4. دوال الجلسات (مع المرحلة ونوع المستخدم)
// =================================================================

/**
 * جلب جميع الأرقام المحفوظة في الشيت
 */
function getSavedPhoneSessions() {
  try {
    var sheet = getWhatsAppSessionsSheet();
    var data = sheet.getDataRange().getValues();
    
    var sessions = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        // تحويل التاريخ بشكل آمن
        var linkedDate = data[i][4];
        var lastUsed = data[i][5];
        
        if (linkedDate instanceof Date) {
          linkedDate = Utilities.formatDate(linkedDate, 'Asia/Riyadh', 'yyyy/MM/dd HH:mm');
        }
        if (lastUsed instanceof Date) {
          lastUsed = Utilities.formatDate(lastUsed, 'Asia/Riyadh', 'yyyy/MM/dd HH:mm');
        }
        
        sessions.push({
          phone: String(data[i][0] || ''),
          stage: String(data[i][1] || ''),
          userType: String(data[i][2] || ''),
          status: String(data[i][3] || 'غير معروف'),
          linkedDate: String(linkedDate || ''),
          lastUsed: String(lastUsed || ''),
          messageCount: parseInt(data[i][6]) || 0,
          isPrimary: String(data[i][7] || '') === 'نعم'
        });
      }
    }
    
    var result = { success: true, sessions: sessions };
    return result;
  } catch (e) {
    var errorResult = { success: false, error: String(e.toString()), sessions: [] };
    return errorResult;
  }
}

/**
 * جلب الأرقام المحفوظة حسب المرحلة
 */
function getSavedPhonesByStage(stage) {
  try {
    var allSessions = getSavedPhoneSessions();
    if (!allSessions.success) {
      var errorResult = { success: false, error: String(allSessions.error || 'خطأ'), sessions: [] };
      return errorResult;
    }
    
    var filtered = [];
    for (var i = 0; i < allSessions.sessions.length; i++) {
      if (allSessions.sessions[i].stage === stage) {
        filtered.push(allSessions.sessions[i]);
      }
    }
    
    var result = { success: true, sessions: filtered };
    return result;
  } catch (e) {
    var errorResult = { success: false, error: String(e.toString()), sessions: [] };
    return errorResult;
  }
}

/**
 * جلب الأرقام المحفوظة حسب المرحلة ونوع المستخدم
 */
function getSavedPhonesByStageAndType(stage, userType) {
  try {
    var allSessions = getSavedPhoneSessions();
    if (!allSessions.success) {
      return allSessions;
    }
    
    var filtered = allSessions.sessions.filter(function(s) {
      return s.stage === stage && s.userType === userType;
    });
    
    return { success: true, sessions: filtered };
  } catch (e) {
    return { success: false, error: e.toString(), sessions: [] };
  }
}

/**
 * التحقق من وجود رقم في الشيت للمرحلة ونوع المستخدم
 */
function isPhoneSavedForStage(phone, stage, userType) {
  var cleanPhone = cleanPhoneNumber(phone);
  var saved = getSavedPhonesByStageAndType(stage, userType);
  
  if (!saved.success) return false;
  
  return saved.sessions.some(function(s) {
    return cleanPhoneNumber(s.phone) === cleanPhone;
  });
}

/**
 * حفظ رقم واتساب جديد مع المرحلة ونوع المستخدم
 */
function saveWhatsAppPhone(phone, stage, userType) {
  try {
    var cleanPhone = cleanPhoneNumber(phone);
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('جلسات_واتساب');
    
    if (!sheet) {
      sheet = getWhatsAppSessionsSheet();
    }
    
    var data = sheet.getDataRange().getValues();
    var now = Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy/MM/dd HH:mm');
    
    // ★ التحقق من نمط الواتساب
    var whatsappMode = getWhatsAppMode_();
    
    // ★ في النمط الموحد: المرحلة تكون "الكل"
    var effectiveStage = (whatsappMode === 'unified') ? 'الكل' : stage;
    
    // التحقق من عدم وجود الرقم مسبقاً
    for (var i = 1; i < data.length; i++) {
      if (cleanPhoneNumber(String(data[i][0])) === cleanPhone && 
          String(data[i][1]) === effectiveStage && 
          String(data[i][2]) === userType) {
        // تحديث الرقم الموجود
        sheet.getRange(i + 1, 4).setValue('متصل');
        sheet.getRange(i + 1, 6).setValue(now);
        // ★ جعله رئيسي تلقائياً
        clearPrimaryForStage_(sheet, data, effectiveStage);
        sheet.getRange(i + 1, 8).setValue('نعم');
        SpreadsheetApp.flush();
        var result = { success: true, message: String('تم تحديث الرقم وتعيينه كرئيسي'), phone: String(cleanPhone), isPrimary: true };
        return result;
      }
    }
    
    // ★ إزالة الرئيسي القديم للمرحلة
    clearPrimaryForStage_(sheet, data, effectiveStage);
    
    // إضافة رقم جديد كرئيسي
    sheet.appendRow([cleanPhone, effectiveStage, userType, 'متصل', now, now, 0, 'نعم']);
    SpreadsheetApp.flush();
    
    var result = { success: true, message: String('تم حفظ الرقم كرقم رئيسي'), phone: String(cleanPhone), isPrimary: true };
    return result;
    
  } catch (e) {
    var errorResult = { success: false, error: String(e.toString()) };
    return errorResult;
  }
}

/**
 * ★ إزالة علامة "رئيسي" من جميع أرقام المرحلة
 */
function clearPrimaryForStage_(sheet, data, stage) {
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === stage && String(data[i][7]) === 'نعم') {
      sheet.getRange(i + 1, 8).setValue('');
    }
  }
}

/**
 * ★ جلب نمط الواتساب من الإعدادات
 */
function getWhatsAppMode_() {
  try {
    var result = getSchoolSettings();
    if (result.success && result.data) {
      return result.data.whatsapp_mode || 'per_stage';
    }
    return 'per_stage';
  } catch (e) {
    return 'per_stage';
  }
}

/**
 * ★ جلب الرقم الرئيسي للمرحلة
 */
function getPrimaryPhoneForStage(stage) {
  try {
    var whatsappMode = getWhatsAppMode_();
    var effectiveStage = (whatsappMode === 'unified') ? 'الكل' : stage;
    
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('جلسات_واتساب');
    if (!sheet) return { success: false, error: 'شيت الجلسات غير موجود' };
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === effectiveStage && String(data[i][7]) === 'نعم') {
        return {
          success: true,
          phone: String(data[i][0]),
          stage: effectiveStage,
          userType: String(data[i][2]),
          status: String(data[i][3]),
          isPrimary: true
        };
      }
    }
    
    return { success: false, error: 'لا يوجد رقم رئيسي لمرحلة ' + stage };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * ★ تعيين رقم كرئيسي يدوياً
 */
function setPrimaryPhone(phone, stage) {
  try {
    var cleanPhone = cleanPhoneNumber(phone);
    var whatsappMode = getWhatsAppMode_();
    var effectiveStage = (whatsappMode === 'unified') ? 'الكل' : stage;
    
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('جلسات_واتساب');
    if (!sheet) return { success: false, error: 'شيت الجلسات غير موجود' };
    
    var data = sheet.getDataRange().getValues();
    
    // إزالة الرئيسي القديم
    clearPrimaryForStage_(sheet, data, effectiveStage);
    
    // تعيين الجديد
    for (var i = 1; i < data.length; i++) {
      if (cleanPhoneNumber(String(data[i][0])) === cleanPhone && String(data[i][1]) === effectiveStage) {
        sheet.getRange(i + 1, 8).setValue('نعم');
        SpreadsheetApp.flush();
        return { success: true, message: 'تم تعيين ' + cleanPhone + ' كرقم رئيسي' };
      }
    }
    
    return { success: false, error: 'الرقم غير موجود' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * حذف رقم واتساب من الشيت
 */
function deleteWhatsAppPhone(phone, stage, userType) {
  try {
    var cleanPhone = cleanPhoneNumber(phone);
    var sheet = getWhatsAppSessionsSheet();
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (cleanPhoneNumber(data[i][0]) === cleanPhone && 
          data[i][1] === stage && 
          data[i][2] === userType) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'تم حذف الرقم' };
      }
    }
    
    return { success: false, error: 'الرقم غير موجود' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * تحديث حالة رقم واتساب
 */
function updatePhoneStatus(phone, stage, userType, status) {
  try {
    var cleanPhone = cleanPhoneNumber(phone);
    var sheet = getWhatsAppSessionsSheet();
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (cleanPhoneNumber(data[i][0]) === cleanPhone && 
          data[i][1] === stage && 
          data[i][2] === userType) {
        sheet.getRange(i + 1, 4).setValue(status);
        return { success: true };
      }
    }
    
    return { success: false, error: 'الرقم غير موجود' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * زيادة عداد الرسائل لرقم معين
 */
function incrementPhoneMessageCount(phone, stage, userType) {
  try {
    var cleanPhone = cleanPhoneNumber(phone);
    var sheet = getWhatsAppSessionsSheet();
    var data = sheet.getDataRange().getValues();
    var now = Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy/MM/dd HH:mm');
    
    for (var i = 1; i < data.length; i++) {
      if (cleanPhoneNumber(data[i][0]) === cleanPhone && 
          data[i][1] === stage && 
          data[i][2] === userType) {
        var currentCount = parseInt(data[i][6]) || 0;
        sheet.getRange(i + 1, 7).setValue(currentCount + 1);
        sheet.getRange(i + 1, 6).setValue(now);
        return { success: true };
      }
    }
    
    return { success: false };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 5. دوال السيرفر
// =================================================================

/**
 * جلب جميع الأرقام المتصلة من السيرفر
 */
function getAllConnectedSessions() {
  try {
    if (!WHATSAPP_SERVER_URL) {
      return { success: false, error: 'رابط سيرفر الواتساب غير مُعيّن — عيّنه في إعدادات الواتساب', sessions: [] };
    }
    var response = UrlFetchApp.fetch(WHATSAPP_SERVER_URL + "/sessions", {
      muteHttpExceptions: true
    });
    
    var text = response.getContentText();
    
    try {
      var jsonResult = JSON.parse(text);
      if (Array.isArray(jsonResult)) {
        var sessions = [];
        for (var i = 0; i < jsonResult.length; i++) {
          var s = jsonResult[i];
          sessions.push({
            phone: String(s.phone || s.id || s),
            status: String(s.status || 'متصل')
          });
        }
        var result = { success: true, sessions: sessions };
        return result;
      }
    } catch (jsonError) {
      // ليس JSON، نحاول استخراج الأرقام من HTML
    }
    
    var phoneMatches = text.match(/966\d{9}/g);
    
    if (phoneMatches && phoneMatches.length > 0) {
      var uniquePhones = [];
      for (var j = 0; j < phoneMatches.length; j++) {
        if (uniquePhones.indexOf(phoneMatches[j]) === -1) {
          uniquePhones.push(phoneMatches[j]);
        }
      }
      
      var sessions = [];
      for (var k = 0; k < uniquePhones.length; k++) {
        sessions.push({
          phone: String(uniquePhones[k]),
          status: 'متصل'
        });
      }
      
      var result = { success: true, sessions: sessions };
      return result;
    }
    
    var emptyResult = { success: true, sessions: [] };
    return emptyResult;
  } catch (e) {
    var errorResult = { success: false, error: String(e.toString()), sessions: [] };
    return errorResult;
  }
}

/**
 * جلب الأرقام المتصلة حسب المرحلة (من الشيت المحلي + تحقق من السيرفر)
 */
function getConnectedSessionsByStage(stage) {
  try {
    // 1. جلب الأرقام من السيرفر
    var serverSessions = getAllConnectedSessions();
    var connectedPhones = [];
    
    if (serverSessions.success && serverSessions.sessions) {
      connectedPhones = serverSessions.sessions.map(function(s) {
        return String(s.phone);
      });
    }
    
    // 2. جلب الأرقام المحفوظة للمرحلة
    var savedSessions = getSavedPhonesByStage(stage);
    
    if (!savedSessions.success) {
      var errorResult = { success: false, error: String(savedSessions.error || 'خطأ غير معروف'), sessions: [], allSessions: [] };
      return errorResult;
    }
    
    // 3. تحديث حالة الاتصال للأرقام المحفوظة
    var sessions = [];
    for (var i = 0; i < savedSessions.sessions.length; i++) {
      var s = savedSessions.sessions[i];
      var isConnected = connectedPhones.indexOf(String(s.phone)) !== -1;
      sessions.push({
        phone: String(s.phone || ''),
        stage: String(s.stage || ''),
        userType: String(s.userType || ''),
        status: isConnected ? 'متصل' : 'غير متصل',
        linkedDate: String(s.linkedDate || ''),
        lastUsed: String(s.lastUsed || ''),
        messageCount: parseInt(s.messageCount) || 0,
        isPrimary: s.isPrimary || false
      });
    }
    
    // فلترة المتصلين فقط
    var connectedSessions = [];
    for (var j = 0; j < sessions.length; j++) {
      if (sessions[j].status === 'متصل') {
        connectedSessions.push(sessions[j]);
      }
    }
    
    var result = { success: true, sessions: connectedSessions, allSessions: sessions };
    return result;
  } catch (e) {
    var errorResult = { success: false, error: String(e.toString()), sessions: [], allSessions: [] };
    return errorResult;
  }
}

/**
 * جلب QR Code لربط جديد
 */
function getWhatsAppQR() {
  try {
    if (!WHATSAPP_SERVER_URL) {
      return { success: false, error: 'رابط سيرفر الواتساب غير مُعيّن' };
    }
    var response = UrlFetchApp.fetch(WHATSAPP_SERVER_URL + "/qr", {
      muteHttpExceptions: true,
      followRedirects: true
    });
    
    var finalUrl = response.getAllHeaders()['Location'] || WHATSAPP_SERVER_URL + "/qr";
    
    return { 
      success: true, 
      qrUrl: finalUrl,
      message: 'افتح الرابط في المتصفح لمسح الباركود'
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * فحص حالة رقم معين في السيرفر
 */
function checkPhoneStatusInServer(phone) {
  try {
    var sessions = getAllConnectedSessions();
    if (!sessions.success) {
      return { connected: false, error: sessions.error };
    }
    
    var cleanPhone = cleanPhoneNumber(phone);
    var found = sessions.sessions.find(function(s) {
      return s.phone === cleanPhone || s.phone.indexOf(cleanPhone) !== -1;
    });
    
    return { connected: !!found, phone: found ? found.phone : null };
  } catch (e) {
    return { connected: false, error: e.toString() };
  }
}

/**
 * فحص حالة الاتصال حسب المرحلة
 */
function getWhatsAppStatus(stage) {
  try {
    // 1. التحقق من إعداد رمز الأمان
    if (!isSecurityCodeConfigured()) {
      return {
        connected: false,
        phone: null,
        needSetup: true,
        sessions: [],
        stage: String(stage || ''),
        whatsappMode: getWhatsAppMode_()
      };
    }
    
    // ★ 2. التحقق من نمط الواتساب
    var whatsappMode = getWhatsAppMode_();
    var effectiveStage = (whatsappMode === 'unified') ? 'الكل' : stage;
    
    // 3. جلب الأرقام المتصلة
    var stageSessions = getConnectedSessionsByStage(effectiveStage);
    
    // ★ 4. جلب الرقم الرئيسي
    var primaryResult = getPrimaryPhoneForStage(stage);
    var primaryPhone = primaryResult.success ? primaryResult.phone : null;
    
    if (stageSessions.success && stageSessions.sessions && stageSessions.sessions.length > 0) {
      return {
        connected: true,
        phone: primaryPhone || String(stageSessions.sessions[0].phone),
        primaryPhone: primaryPhone,
        hasPrimary: !!primaryPhone,
        needSetup: false,
        sessions: stageSessions.sessions,
        allSessions: stageSessions.allSessions || [],
        stage: String(stage || ''),
        effectiveStage: String(effectiveStage),
        whatsappMode: whatsappMode
      };
    }
    
    // 5. لا توجد أرقام متصلة
    return {
      connected: false,
      phone: null,
      primaryPhone: primaryPhone,
      hasPrimary: !!primaryPhone,
      needSetup: false,
      sessions: [],
      allSessions: stageSessions.allSessions || [],
      stage: String(stage || ''),
      effectiveStage: String(effectiveStage),
      whatsappMode: whatsappMode
    };
  } catch (e) {
    return { 
      connected: false, 
      error: String(e.toString()), 
      sessions: [], 
      stage: String(stage || ''),
      whatsappMode: 'per_stage'
    };
  }
}

/**
 * إيقاظ السيرفر
 */
function pingWhatsAppServer() {
  try {
    if (!WHATSAPP_SERVER_URL) return { success: false };
    UrlFetchApp.fetch(WHATSAPP_SERVER_URL + "/", { muteHttpExceptions: true });
    return { success: true };
  } catch (e) {
    return { success: false };
  }
}

// =================================================================
// 6. دوال الإرسال
// =================================================================

/**
 * تنظيف رقم الجوال
 */
function cleanPhoneNumber(phone) {
  var clean = phone.toString().replace(/\D/g, '');
  if (clean.substring(0, 2) === '05') {
    clean = '966' + clean.substring(1);
  } else if (clean.substring(0, 1) === '5' && clean.length === 9) {
    clean = '966' + clean;
  } else if (clean.substring(0, 3) !== '966' && clean.length === 9) {
    clean = '966' + clean;
  }
  return clean;
}

/**
 * إرسال رسالة من رقم محدد
 */
function sendWhatsAppMessageFrom(senderPhone, recipientPhone, message) {
  try {
    if (!WHATSAPP_SERVER_URL) {
      return { success: false, error: 'رابط سيرفر الواتساب غير مُعيّن' };
    }
    var cleanSender = cleanPhoneNumber(senderPhone);
    var cleanRecipient = cleanPhoneNumber(recipientPhone);
    
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

    var response = UrlFetchApp.fetch(WHATSAPP_SERVER_URL + "/send/" + cleanSender, options);
    return JSON.parse(response.getContentText());

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * إرسال رسالة واتساب حسب المرحلة (يستخدم رقم المرحلة فقط)
 */
function sendWhatsAppMessage(recipientPhone, message, stage) {
  try {
    // ★ جلب الرقم الرئيسي للمرحلة
    var primaryResult = getPrimaryPhoneForStage(stage);
    
    if (!primaryResult.success) {
      // بديل: جلب أي رقم متصل
      var sessions = getConnectedSessionsByStage(stage);
      if (!sessions.success || !sessions.sessions || sessions.sessions.length === 0) {
        return { 
          success: false, 
          error: 'لا يوجد رقم رئيسي متصل لمرحلة ' + stage + '. يرجى ربط رقم رئيسي من أدوات واتساب.' 
        };
      }
      // استخدام أول رقم متصل كبديل
      primaryResult = { phone: sessions.sessions[0].phone, userType: sessions.sessions[0].userType };
    }
    
    var senderPhone = primaryResult.phone;
    var result = sendWhatsAppMessageFrom(senderPhone, recipientPhone, message);
    
    // زيادة عداد الرسائل
    if (result.success) {
      try {
        incrementPhoneMessageCount(senderPhone, stage, primaryResult.userType || 'وكيل');
      } catch(e) {}
    }
    
    return result;
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * إرسال رسالة مع التوثيق حسب المرحلة
 */
function sendWhatsAppWithLog(data, stage) {
  try {
    // ★ 1. جلب الرقم الرئيسي للمرحلة
    var primaryResult = getPrimaryPhoneForStage(stage);
    var senderPhone, senderUserType;
    
    if (primaryResult.success) {
      senderPhone = primaryResult.phone;
      senderUserType = primaryResult.userType;
    } else {
      // بديل: أي رقم متصل
      var sessions = getConnectedSessionsByStage(stage);
      if (!sessions.success || !sessions.sessions || sessions.sessions.length === 0) {
        return { 
          success: false, 
          error: 'لا يوجد رقم رئيسي متصل لمرحلة ' + stage + '. يرجى ربط رقم رئيسي من أدوات واتساب.'
        };
      }
      senderPhone = sessions.sessions[0].phone;
      senderUserType = sessions.sessions[0].userType;
    }
    
    // 2. تسجيل الرسالة في سجل التواصل
    var logResult = logCommunication({
      studentId: data.studentId,
      studentName: data.studentName,
      grade: data.grade,
      class: data.class,
      phone: data.phone,
      messageType: data.messageType,
      messageTitle: data.messageTitle,
      messageContent: data.message,
      sender: data.sender || 'الوكيل',
      status: 'جاري الإرسال'
    }, stage);
    
    if (!logResult.success) {
      return { success: false, error: 'فشل في تسجيل الرسالة: ' + logResult.error };
    }
    
    // 3. إرسال الرسالة
    var sendResult = sendWhatsAppMessageFrom(senderPhone, data.phone, data.message);
    
    // 4. تحديث حالة الإرسال
    if (sendResult.success) {
      updateCommunicationStatus(logResult.logId, '✅ تم الإرسال', '', stage);
      incrementPhoneMessageCount(senderPhone, stage, senderUserType);
      return { success: true, logId: logResult.logId };
    } else {
      updateCommunicationStatus(logResult.logId, '❌ فشل', sendResult.error || 'خطأ غير معروف', stage);
      return { success: false, error: sendResult.error, logId: logResult.logId };
    }
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 7. دوال الإحصائيات
// =================================================================

/**
 * جلب إحصائيات الأرقام حسب المرحلة
 */
function getWhatsAppStats(stage) {
  try {
    var stageSessions = getConnectedSessionsByStage(stage);
    var savedSessions = getSavedPhonesByStage(stage);
    
    var totalMessages = 0;
    if (savedSessions.success && savedSessions.sessions) {
      for (var i = 0; i < savedSessions.sessions.length; i++) {
        totalMessages += parseInt(savedSessions.sessions[i].messageCount) || 0;
      }
    }
    
    var result = {
      success: true,
      stats: {
        connectedPhones: stageSessions.sessions ? stageSessions.sessions.length : 0,
        savedPhones: savedSessions.sessions ? savedSessions.sessions.length : 0,
        totalMessages: totalMessages,
        sessions: stageSessions.sessions || [],
        allSessions: stageSessions.allSessions || []
      },
      stage: String(stage || '')
    };
    return result;
  } catch (e) {
    var errorResult = { success: false, error: String(e.toString()) };
    return errorResult;
  }
}

/**
 * جلب إحصائيات جميع الأرقام (لجميع المراحل)
 */
function getAllPhonesStats() {
  try {
    var sheet = getWhatsAppSessionsSheet();
    var data = sheet.getDataRange().getValues();
    
    var phones = [];
    var totalMessages = 0;
    var connectedCount = 0;
    
    // جلب الأرقام المتصلة من السيرفر
    var serverSessions = getAllConnectedSessions();
    var connectedPhones = [];
    if (serverSessions.success && serverSessions.sessions) {
      for (var j = 0; j < serverSessions.sessions.length; j++) {
        connectedPhones.push(String(serverSessions.sessions[j].phone));
      }
    }
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var isConnected = connectedPhones.indexOf(String(data[i][0])) !== -1;
        
        // تحويل التاريخ بشكل آمن
        var linkedDate = data[i][4];
        var lastUsed = data[i][5];
        
        if (linkedDate instanceof Date) {
          linkedDate = Utilities.formatDate(linkedDate, 'Asia/Riyadh', 'yyyy/MM/dd HH:mm');
        }
        if (lastUsed instanceof Date) {
          lastUsed = Utilities.formatDate(lastUsed, 'Asia/Riyadh', 'yyyy/MM/dd HH:mm');
        }
        
        var phone = {
          phone: String(data[i][0] || ''),
          stage: String(data[i][1] || ''),
          userType: String(data[i][2] || ''),
          status: isConnected ? 'متصل' : 'غير متصل',
          linkedDate: String(linkedDate || ''),
          lastUsed: String(lastUsed || ''),
          messageCount: parseInt(data[i][6]) || 0
        };
        
        phones.push(phone);
        totalMessages += phone.messageCount;
        if (isConnected) connectedCount++;
      }
    }
    
    var result = {
      success: true,
      stats: {
        totalPhones: phones.length,
        connectedPhones: connectedCount,
        totalMessages: totalMessages,
        phones: phones
      }
    };
    return result;
  } catch (e) {
    var errorResult = { success: false, error: String(e.toString()) };
    return errorResult;
  }
}

// =================================================================
// 8. دوال مزامنة الأرقام من السيرفر
// =================================================================

/**
 * مزامنة رقم من السيرفر وحفظه للمرحلة ونوع المستخدم
 */
function syncAndSavePhone(phone, stage, userType) {
  try {
    // التحقق من أن الرقم متصل في السيرفر
    var serverStatus = checkPhoneStatusInServer(phone);
    
    if (!serverStatus.connected) {
      return { success: false, error: 'الرقم غير متصل في السيرفر' };
    }
    
    // حفظ الرقم مع المرحلة ونوع المستخدم
    var saveResult = saveWhatsAppPhone(serverStatus.phone, stage, userType);
    
    return saveResult;
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * جلب أنواع المستخدمين المتاحة
 */
function getUserTypes() {
  return { success: true, types: USER_TYPES };
}

// =================================================================
// 9. دوال الاختبار
// =================================================================

function TEST_WhatsApp() {
  Logger.log("=== اختبار نظام الواتساب ===");
  
  Logger.log("1. فحص الاتصال بالسيرفر...");
  var sessions = getAllConnectedSessions();
  Logger.log("   النتيجة: " + JSON.stringify(sessions));
  
  Logger.log("2. فحص حالة الواتساب للمتوسط...");
  var status = getWhatsAppStatus('متوسط');
  Logger.log("   الحالة: " + JSON.stringify(status));
  
  Logger.log("3. الإحصائيات للمتوسط...");
  var stats = getWhatsAppStats('متوسط');
  Logger.log("   الإحصائيات: " + JSON.stringify(stats));
  
  Logger.log("=== انتهى الاختبار ===");
}

function TEST_RebuildSheet() {
  Logger.log("=== إعادة بناء شيت الجلسات ===");
  var result = rebuildSessionsSheet();
  Logger.log("النتيجة: " + JSON.stringify(result));
}

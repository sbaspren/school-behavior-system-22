// =================================================================
// Server_SMS.gs - إرسال الرسائل النصية SMS عبر Madar API
// =================================================================

// =================================================================
// إعدادات Madar API — تُقرأ من PropertiesService (لا تُخزّن هنا مباشرة)
// لتعيين القيم: افتح Script Editor → File → Project properties → Script properties
//   sms_api_token = الرمز
//   sms_sender_name = اسم المرسل
// =================================================================
function getMadarConfig_() {
  var props = PropertiesService.getScriptProperties();
  return {
    apiToken: props.getProperty('sms_api_token') || '',
    senderName: props.getProperty('sms_sender_name') || 'School',
    apiSendEndpoint: 'https://app.mobile.net.sa/api/v1/send'
  };
}

// =================================================================
// قوالب الرسائل — اسم المدرسة يُجلب ديناميكياً
// =================================================================
function getSmsTemplates_() {
  var schoolName = '';
  try { schoolName = getSchoolNameForLinks_(); } catch(e) { schoolName = 'المدرسة'; }
  return {
    'تأخر': 'المكرم ولي الامر نود إبلاغكم بتأخر ابنكم {student_name} عن الحضور إلى المدرسة لهذا اليوم بتاريخ {date}\n' + schoolName,
    'استئذان': 'المكرم ولي الامر نود إبلاغكم باستئذان ابنكم {student_name} للخروج من المدرسة لهذا اليوم بتاريخ {date}\n' + schoolName
  };
}

// =================================================================
// 1. إرسال رسالة SMS واحدة
// =================================================================
function sendSingleSMS(phoneNumber, message) {
  try {
    var config = getMadarConfig_();
    if (!config.apiToken) {
      return { success: false, error: 'رمز SMS API غير مُعيّن — عيّنه في Script Properties (sms_api_token)' };
    }

    // تنظيف رقم الجوال
    var cleanPhone = String(phoneNumber).replace(/\D/g, '');

    // التحقق من صحة الرقم
    if (!cleanPhone || cleanPhone.length < 9) {
      return { success: false, error: 'رقم الجوال غير صالح' };
    }

    // إضافة مفتاح الدولة إذا لزم الأمر
    if (cleanPhone.startsWith('05')) {
      cleanPhone = '966' + cleanPhone.substring(1);
    } else if (cleanPhone.startsWith('5')) {
      cleanPhone = '966' + cleanPhone;
    } else if (!cleanPhone.startsWith('966')) {
      cleanPhone = '966' + cleanPhone;
    }

    var payload = {
      number: cleanPhone,
      senderName: config.senderName,
      sendAtOption: 'NOW',
      messageBody: message,
      allow_duplicate: false
    };

    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + config.apiToken,
        'Accept': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(config.apiSendEndpoint, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    if (responseCode === 200 || responseCode === 201) {
      return { success: true, message: 'تم الإرسال بنجاح' };
    } else {
      return { success: false, error: 'خطأ API: ' + responseCode + ' - ' + responseText };
    }
    
  } catch (error) {
    return { success: false, error: error.message || error.toString() };
  }
}

// =================================================================
// 2. إرسال رسائل SMS متعددة
// =================================================================
function sendBulkSMS(recipients, messageTemplate, type) {
  var results = {
    success: 0,
    failed: 0,
    errors: []
  };
  
  var date = new Date().toLocaleDateString('ar-SA');
  
  for (var i = 0; i < recipients.length; i++) {
    var recipient = recipients[i];
    var message = messageTemplate
      .replace('{student_name}', recipient.name || '')
      .replace('{اسم_الطالب}', recipient.name || '')
      .replace('{date}', date);
    
    var result = sendSingleSMS(recipient.phone, message);
    
    if (result.success) {
      results.success++;
    } else {
      results.failed++;
      results.errors.push({
        name: recipient.name,
        phone: recipient.phone,
        error: result.error
      });
    }
    
    // تأخير بسيط بين الرسائل لتجنب الحظر
    if (i < recipients.length - 1) {
      Utilities.sleep(100);
    }
  }
  
  return {
    success: true,
    results: results,
    message: 'تم إرسال ' + results.success + ' رسالة بنجاح، فشل ' + results.failed
  };
}

// =================================================================
// 3. إنشاء رسالة من القالب
// =================================================================
function createSMSMessage(type, studentName) {
  var templates = getSmsTemplates_();
  var template = templates[type] || templates['تأخر'];
  var date = new Date().toLocaleDateString('ar-SA');

  return template
    .replace('{student_name}', studentName)
    .replace('{date}', date);
}

// =================================================================
// 4. التحقق من رصيد SMS (اختياري)
// =================================================================
function checkSMSBalance() {
  try {
    var config = getMadarConfig_();
    if (!config.apiToken) return { success: false, error: 'رمز SMS API غير مُعيّن' };

    var options = {
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + config.apiToken,
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch('https://app.mobile.net.sa/api/v1/account/balance', options);
    var responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      var data = JSON.parse(response.getContentText());
      return { success: true, balance: data };
    } else {
      return { success: false, error: 'فشل جلب الرصيد' };
    }
    
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// =================================================================
// 5. دالة اختبار SMS
// =================================================================
function testSMSConnection() {
  Logger.log('جاري اختبار اتصال SMS...');
  var config = getMadarConfig_();

  if (!config.apiToken) {
    Logger.log('⚠️ رمز SMS API غير مُعيّن في Script Properties');
    return { success: false, message: 'رمز SMS API غير مُعيّن — عيّنه في Script Properties (sms_api_token)' };
  }

  Logger.log('Token: ' + config.apiToken.substring(0, 10) + '...');
  Logger.log('Sender: ' + config.senderName);

  return { success: true, message: 'الإعدادات صحيحة' };
}
// =================================================================
// Server_Templates.gs - إدارة قوالب الرسائل
// =================================================================

/**
 * حفظ قالب رسالة كافتراضي لنوع معين
 * @param {string} type - نوع الرسالة (مخالفة، تأخر، غياب، استئذان، ملاحظة)
 * @param {string} message - نص القالب
 */
function saveMessageTemplate(type, message) {
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('msg_template_' + type, message);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * جلب قالب الرسالة المحفوظ
 * @param {string} type - نوع الرسالة
 * @returns {object} {success, template}
 */
function getMessageTemplate(type) {
  try {
    var props = PropertiesService.getScriptProperties();
    var template = props.getProperty('msg_template_' + type);
    return { success: true, template: template || '' };
  } catch (e) {
    return { success: false, template: '', error: e.message };
  }
}

/**
 * حذف قالب محفوظ (استعادة الافتراضي)
 * @param {string} type - نوع الرسالة
 */
function deleteMessageTemplate(type) {
  try {
    var props = PropertiesService.getScriptProperties();
    props.deleteProperty('msg_template_' + type);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * جلب جميع القوالب المحفوظة
 */
function getAllMessageTemplates() {
  try {
    var props = PropertiesService.getScriptProperties();
    var all = props.getProperties();
    var templates = {};
    for (var key in all) {
      if (key.indexOf('msg_template_') === 0) {
        var type = key.replace('msg_template_', '');
        templates[type] = all[key];
      }
    }
    return { success: true, templates: templates };
  } catch (e) {
    return { success: false, templates: {}, error: e.message };
  }
}

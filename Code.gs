/**
 * DALEELAK SUDANI — MVP + ADMIN DASHBOARD
 * Google Apps Script backend
 *
 * Keep the existing Providers / Categories / Settings sheet structure.
 * Admin password is read once from Script Properties (ADMIN_PASSWORD),
 * then migrated to a salted SHA-256 hash and the plain property is deleted.
 */

const APP = {
  title: 'دليلك السوداني',
  version: '2.1.0',
  defaultCategories: [
    'صيانة المنازل', 'السيارات والنقل', 'العقارات والسكن', 'التعليم والتدريب',
    'التقنية والتصميم', 'المطاعم والطبخ المنزلي', 'المناسبات والخدمات الشخصية',
    'خدمات أخرى', 'تاكسي ومشاوير', 'توصيل طلبات', 'توصيل طرود وأغراض',
    'خدمات التوصيل والترحيب', 'نقل عفش وأثاث'
  ],
  sheets: {
    providers: 'Providers',
    categories: 'Categories',
    settings: 'Settings'
  },
  adminSessionSeconds: 21600,
  listingFee: 20,
  paymentColumns: ['PaymentStatus', 'PaymentAmount', 'PaymentReference', 'PaymentSubmittedAt', 'PaymentVerifiedAt']
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP.title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Cross-origin bridge for the GitHub Pages frontend.
 * Requests are submitted as ordinary HTML forms to avoid browser CORS
 * restrictions. The response is delivered to the exact requesting iframe
 * with postMessage; admin passwords and tokens are never placed in URLs.
 */
function doPost(e) {
  var requestId = String((e && e.parameter && e.parameter.requestId) || '');
  if (!/^[A-Za-z0-9_-]{8,100}$/.test(requestId)) requestId = 'invalid_request';

  var result, ok = false, error = '';
  try {
    var action = String((e && e.parameter && e.parameter.action) || '');
    var payloadText = String((e && e.parameter && e.parameter.payload) || '{}');
    if (payloadText.length > 30000) throw new Error('حجم الطلب أكبر من المسموح.');
    var payload = JSON.parse(payloadText || '{}');
    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
      throw new Error('صيغة الطلب غير صحيحة.');
    }

    switch (action) {
      case 'getPublicData': result = getPublicData(); break;
      case 'submitProviderListing': result = submitProviderListing(payload); break;
      case 'adminLogin': result = adminLogin(payload.password); break;
      case 'adminLogout': result = adminLogout(payload.token); break;
      case 'adminGetDashboard': result = adminGetDashboard(payload.token); break;
      case 'adminSetProviderStatus': result = adminSetProviderStatus(payload.token, payload.id, payload.status); break;
      case 'adminConfirmPaymentAndApprove': result = adminConfirmPaymentAndApprove(payload.token, payload.id); break;
      case 'adminSetProviderFeatured': result = adminSetProviderFeatured(payload.token, payload.id, payload.featured); break;
      default: throw new Error('العملية المطلوبة غير معروفة.');
    }
    ok = true;
  } catch (err) {
    error = err && err.message ? String(err.message) : 'حدث خطأ غير متوقع في الخادم.';
  }

  var message = { type: 'daleelak-api-response', requestId: requestId, ok: ok };
  if (ok) message.data = result;
  else message.error = error;
  return HtmlService.createHtmlOutput(buildApiResponseHtml_(message))
    .setTitle('Daleelak API')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function buildApiResponseHtml_(message) {
  // Fixed GitHub Pages origin; never accept a target origin from the request.
  var targetOrigin = 'https://jomujahid007-boop.github.io';
  var json = JSON.stringify(message)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');

  return '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
    '<script>(function(){var m=' + json + ';' +
    'try{window.top.postMessage(m,' + JSON.stringify(targetOrigin) + ');}' +
    'catch(e){try{window.parent.postMessage(m,' + JSON.stringify(targetOrigin) + ');}catch(e2){}}' +
    '})();</script></body></html>';
}


function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('افتح Google Sheet المرتبط بهذا المشروع أولاً.');

  const providers = ensureSheet_(ss, APP.sheets.providers, [
    'ID', 'CreatedAt', 'Status', 'Name', 'Category', 'City',
    'Description', 'Phone', 'WhatsApp', 'Featured'
  ]);
  ensureProviderPaymentColumns_(providers);
  ensureSheet_(ss, APP.sheets.categories, ['Name', 'Active']);
  ensureSheet_(ss, APP.sheets.settings, ['Key', 'Value']);

  const cats = ss.getSheetByName(APP.sheets.categories);
  const defaultCategories = APP.defaultCategories.map(name => [name, true]);
  const existingCategoryNames = cats.getLastRow() >= 2
    ? cats.getRange(2, 1, cats.getLastRow() - 1, 1).getDisplayValues().flat().map(x => String(x).trim())
    : [];
  const missingCategories = defaultCategories.filter(row => existingCategoryNames.indexOf(row[0]) === -1);
  if (missingCategories.length) {
    cats.getRange(cats.getLastRow() + 1, 1, missingCategories.length, 2).setValues(missingCategories);
  }

  const settings = ss.getSheetByName(APP.sheets.settings);
  if (settings.getLastRow() < 2) {
    settings.getRange(2, 1, 2, 2).setValues([
      ['AppName', APP.title],
      ['Version', 'MVP-2']
    ]);
  }

  return 'تم تجهيز قاعدة بيانات دليلك السوداني.';
}

/* -------------------- PUBLIC API -------------------- */

function getPublicData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('تعذر الوصول إلى قاعدة البيانات المرتبطة بالتطبيق.');
  const providersSheet = ss.getSheetByName(APP.sheets.providers);
  const categoriesSheet = ss.getSheetByName(APP.sheets.categories);
  if (!providersSheet || !categoriesSheet) throw new Error('أوراق قاعدة البيانات غير مكتملة. تأكد من وجود Providers و Categories.');

  const providerHeaders = getHeaders_(providersSheet);
  const requiredProviderHeaders = ['ID', 'Status', 'Name', 'Category', 'City', 'Description', 'Phone', 'WhatsApp', 'Featured'];
  const missingHeaders = requiredProviderHeaders.filter(name => providerHeaders.indexOf(name) === -1);
  if (missingHeaders.length) throw new Error('أعمدة ورقة Providers ناقصة: ' + missingHeaders.join(', ') + '. لم يتم تعديل أي بيانات.');
  const categoryHeaders = getHeaders_(categoriesSheet);
  if (categoryHeaders.indexOf('Name') === -1 || categoryHeaders.indexOf('Active') === -1) throw new Error('عناوين ورقة Categories غير صحيحة. المطلوب Name و Active.');

  let categories = readObjects_(categoriesSheet)
    .filter(x => x.Active === true || String(x.Active).toLowerCase() === 'true')
    .map(x => String(x.Name || '').trim()).filter(Boolean);
  if (!categories.length) categories = APP.defaultCategories.slice();

  const approved = readObjects_(providersSheet)
    .filter(x => String(x.Status).toLowerCase() === 'approved')
    .map(x => ({
      id: String(x.ID || ''), name: String(x.Name || ''), category: String(x.Category || ''),
      city: String(x.City || ''), description: String(x.Description || ''), phone: String(x.Phone || ''),
      whatsapp: cleanPhone_(x.WhatsApp || x.Phone || ''),
      featured: x.Featured === true || String(x.Featured).toLowerCase() === 'true',
      createdAt: x.CreatedAt instanceof Date ? x.CreatedAt.getTime() : 0
    }))
    .sort((a, b) => Number(b.featured) - Number(a.featured) || b.createdAt - a.createdAt);

  return {
    ok: true, version: APP.version, title: APP.title, categories: categories,
    providers: approved, totalProviders: approved.length, listingFee: APP.listingFee,
    stcPayNumber: PropertiesService.getScriptProperties().getProperty('STCPAY_NUMBER') || '',
    cities: ['كل المدن', 'جدة', 'الرياض', 'مكة المكرمة', 'المدينة المنورة', 'الدمام', 'الخبر', 'الطائف', 'أبها', 'مدينة أخرى']
  };
}

function submitProviderListing(form) {
  form = form || {};
  const name = clean_(form.name, 80);
  const category = clean_(form.category, 80);
  const city = clean_(form.city, 50);
  const description = clean_(form.description, 500);
  const phone = clean_(form.phone, 25);
  const whatsapp = cleanPhone_(form.whatsapp || phone);
  const paymentReference = clean_(form.paymentReference, 80);

  if (!name || !category || !city || !phone || !paymentReference) {
    throw new Error('يرجى تعبئة البيانات المطلوبة وإدخال رقم عملية التحويل.');
  }
  if (!/^[+\d\s()-]{7,25}$/.test(phone)) {
    throw new Error('رقم التواصل غير صحيح. أدخل الرقم مع رمز الدولة إن أمكن.');
  }
  if (!/^[\w\-\u0600-\u06FF ]{4,80}$/.test(paymentReference)) {
    throw new Error('رقم العملية غير صحيح. راجع الرقم كما يظهر في إيصال التحويل.');
  }
  const paymentNumber = PropertiesService.getScriptProperties().getProperty('STCPAY_NUMBER');
  if (!paymentNumber) throw new Error('بيانات استقبال STC Pay لم تُجهّز بعد. يرجى المحاولة لاحقاً.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getProvidersSheet_();
    ensureProviderPaymentColumns_(sh);
    const headers = getHeaders_(sh);
    const refCol = headers.indexOf('PaymentReference') + 1;
    if (refCol && sh.getLastRow() >= 2) {
      const refs = sh.getRange(2, refCol, sh.getLastRow() - 1, 1).getDisplayValues().flat();
      if (refs.some(x => String(x).trim() === paymentReference)) {
        throw new Error('رقم العملية مستخدم من قبل. تأكد من رقم العملية الصحيح.');
      }
    }

    const record = {
      ID: Utilities.getUuid(), CreatedAt: new Date(), Status: 'Pending',
      Name: name, Category: category, City: city, Description: description,
      Phone: phone, WhatsApp: whatsapp, Featured: false,
      PaymentStatus: 'Submitted', PaymentAmount: APP.listingFee,
      PaymentReference: paymentReference, PaymentSubmittedAt: new Date(), PaymentVerifiedAt: ''
    };
    sh.appendRow(headers.map(h => Object.prototype.hasOwnProperty.call(record, h) ? record[h] : ''));
    return { ok: true, message: 'وصل إعلانك ورقم العملية بنجاح. الإعلان الآن بانتظار تحقق الإدارة من وصول مبلغ 20 ريال.' };
  } finally {
    lock.releaseLock();
  }
}

/* -------------------- ADMIN AUTHENTICATION -------------------- */

/**
 * First login migrates ADMIN_PASSWORD into a salted hash.
 * Set ADMIN_PASSWORD under Project Settings > Script Properties before use.
 */
function adminLogin(password) {
  const supplied = String(password || '');
  if (!supplied) throw new Error('أدخل كلمة مرور الإدارة.');

  const props = PropertiesService.getScriptProperties();
  let salt = props.getProperty('ADMIN_PASSWORD_SALT');
  let savedHash = props.getProperty('ADMIN_PASSWORD_HASH');

  if (!savedHash) {
    const initialPassword = props.getProperty('ADMIN_PASSWORD');
    if (!initialPassword) {
      throw new Error('لم يتم إعداد كلمة مرور الإدارة. أضف ADMIN_PASSWORD في Script Properties.');
    }
    if (initialPassword.length < 8) {
      throw new Error('كلمة مرور الإدارة يجب أن تكون 8 أحرف على الأقل. حدّث ADMIN_PASSWORD من Script Properties.');
    }
    salt = Utilities.getUuid() + Utilities.getUuid();
    savedHash = passwordHash_(initialPassword, salt);
    props.setProperties({
      ADMIN_PASSWORD_SALT: salt,
      ADMIN_PASSWORD_HASH: savedHash
    });
    props.deleteProperty('ADMIN_PASSWORD');
  }

  if (!salt || passwordHash_(supplied, salt) !== savedHash) {
    throw new Error('كلمة المرور غير صحيحة.');
  }

  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('admin_session_' + token, '1', APP.adminSessionSeconds);
  return { ok: true, token: token, expiresIn: APP.adminSessionSeconds };
}

function adminLogout(token) {
  if (token) CacheService.getScriptCache().remove('admin_session_' + String(token));
  return { ok: true };
}

function adminGetDashboard(token) {
  requireAdmin_(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss && ss.getSheetByName(APP.sheets.providers);
  if (!sh) throw new Error('ورقة Providers غير موجودة. شغّل setupDatabase أولاً.');

  return readObjects_(sh).map(x => ({
    id: String(x.ID || ''),
    createdAt: x.CreatedAt instanceof Date
      ? Utilities.formatDate(x.CreatedAt, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
      : String(x.CreatedAt || ''),
    status: String(x.Status || 'Pending'),
    name: String(x.Name || ''),
    category: String(x.Category || ''),
    city: String(x.City || ''),
    description: String(x.Description || ''),
    phone: String(x.Phone || ''),
    whatsapp: cleanPhone_(x.WhatsApp || x.Phone || ''),
    featured: x.Featured === true || String(x.Featured).toLowerCase() === 'true',
    paymentStatus: String(x.PaymentStatus || ''),
    paymentAmount: Number(x.PaymentAmount || 0),
    paymentReference: String(x.PaymentReference || ''),
    paymentSubmittedAt: x.PaymentSubmittedAt instanceof Date
      ? Utilities.formatDate(x.PaymentSubmittedAt, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
      : String(x.PaymentSubmittedAt || ''),
    paymentVerifiedAt: x.PaymentVerifiedAt instanceof Date
      ? Utilities.formatDate(x.PaymentVerifiedAt, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
      : String(x.PaymentVerifiedAt || '')
  }));
}

function adminSetProviderStatus(token, id, status) {
  requireAdmin_(token);
  id = clean_(id, 100);
  status = String(status || '');
  if (!id || ['Approved', 'Rejected', 'Pending'].indexOf(status) === -1) {
    throw new Error('الطلب أو الحالة غير صحيحة.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getProvidersSheet_();
    const headers = getHeaders_(sh);
    const idCol = headers.indexOf('ID') + 1;
    const statusCol = headers.indexOf('Status') + 1;
    if (!idCol || !statusCol) throw new Error('أعمدة ورقة Providers غير مكتملة.');

    const row = findProviderRow_(sh, idCol, id);
    if (!row) throw new Error('لم يتم العثور على الطلب. حدّث القائمة وحاول مرة أخرى.');
    const paymentStatusCol = headers.indexOf('PaymentStatus') + 1;
    const paymentStatus = paymentStatusCol ? String(sh.getRange(row, paymentStatusCol).getValue() || '') : '';
    if (status === 'Approved' && paymentStatus && paymentStatus !== 'Verified') {
      throw new Error('لا يمكن نشر الإعلان قبل تأكيد الدفع من زر «تأكيد الدفع ونشر الإعلان».');
    }
    sh.getRange(row, statusCol).setValue(status);
    return { ok: true, status: status };
  } finally {
    lock.releaseLock();
  }
}

function adminConfirmPaymentAndApprove(token, id) {
  requireAdmin_(token);
  id = clean_(id, 100);
  if (!id) throw new Error('معرّف الطلب غير صحيح.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getProvidersSheet_();
    ensureProviderPaymentColumns_(sh);
    const headers = getHeaders_(sh);
    const idCol = headers.indexOf('ID') + 1;
    const statusCol = headers.indexOf('Status') + 1;
    const paymentStatusCol = headers.indexOf('PaymentStatus') + 1;
    const verifiedAtCol = headers.indexOf('PaymentVerifiedAt') + 1;
    if (!idCol || !statusCol || !paymentStatusCol || !verifiedAtCol) {
      throw new Error('أعمدة الدفع غير مكتملة.');
    }
    const row = findProviderRow_(sh, idCol, id);
    if (!row) throw new Error('لم يتم العثور على الطلب.');
    const currentPaymentStatus = String(sh.getRange(row, paymentStatusCol).getValue() || '');
    if (currentPaymentStatus !== 'Submitted' && currentPaymentStatus !== 'Verified') {
      throw new Error('هذا الطلب لا يحتوي على عملية دفع جديدة بانتظار التحقق.');
    }
    const refCol = headers.indexOf('PaymentReference') + 1;
    const reference = refCol ? String(sh.getRange(row, refCol).getValue() || '').trim() : '';
    if (!reference) throw new Error('لا يوجد رقم عملية مسجل لهذا الطلب.');

    sh.getRange(row, paymentStatusCol).setValue('Verified');
    sh.getRange(row, verifiedAtCol).setValue(new Date());
    sh.getRange(row, statusCol).setValue('Approved');
    return { ok: true, status: 'Approved', paymentStatus: 'Verified' };
  } finally {
    lock.releaseLock();
  }
}

function adminSetProviderFeatured(token, id, featured) {
  requireAdmin_(token);
  id = clean_(id, 100);
  if (!id) throw new Error('معرّف الخدمة غير صحيح.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getProvidersSheet_();
    const headers = getHeaders_(sh);
    const idCol = headers.indexOf('ID') + 1;
    const featuredCol = headers.indexOf('Featured') + 1;
    if (!idCol || !featuredCol) throw new Error('أعمدة ورقة Providers غير مكتملة.');

    const row = findProviderRow_(sh, idCol, id);
    if (!row) throw new Error('لم يتم العثور على الخدمة. حدّث القائمة وحاول مرة أخرى.');
    sh.getRange(row, featuredCol).setValue(featured === true);
    return { ok: true, featured: featured === true };
  } finally {
    lock.releaseLock();
  }
}

function requireAdmin_(token) {
  token = String(token || '');
  if (!token || token.length > 200 ||
      CacheService.getScriptCache().get('admin_session_' + token) !== '1') {
    throw new Error('انتهت جلسة الإدارة. سجّل الدخول مرة أخرى.');
  }
}

function passwordHash_(password, salt) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt) + ':' + String(password),
    Utilities.Charset.UTF_8
  );
  return Utilities.base64Encode(bytes);
}

/* -------------------- HELPERS -------------------- */

function getProvidersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss && ss.getSheetByName(APP.sheets.providers);
  if (!sh) throw new Error('ورقة Providers غير موجودة.');
  return sh;
}

function getHeaders_(sh) {
  if (sh.getLastColumn() < 1) return [];
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
}

function findProviderRow_(sh, idCol, id) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  const values = sh.getRange(2, idCol, lastRow - 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === id) return i + 2;
  }
  return 0;
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  } else {
    const existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length))
      .getValues()[0].slice(0, headers.length).map(String);
    if (existing.join('|') !== headers.join('|')) {
      throw new Error('ترتيب أعمدة ورقة ' + name + ' غير مطابق. لم يتم تعديل أي بيانات.');
    }
  }
  return sh;
}

function ensureProviderPaymentColumns_(sh) {
  const headers = getHeaders_(sh);
  APP.paymentColumns.forEach(name => {
    if (headers.indexOf(name) === -1) {
      const col = sh.getLastColumn() + 1;
      sh.getRange(1, col).setValue(name).setFontWeight('bold');
      headers.push(name);
    }
  });
}

function readObjects_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values.shift().map(String);
  return values.filter(row => row.some(v => v !== '' && v !== null))
    .map(row => Object.fromEntries(headers.map((h, i) => [h, row[i]])));
}

function clean_(value, max) {
  return String(value || '').replace(/[<>]/g, '').trim().slice(0, max);
}

function cleanPhone_(value) {
  return String(value || '').replace(/[^\d+]/g, '').slice(0, 20);
}

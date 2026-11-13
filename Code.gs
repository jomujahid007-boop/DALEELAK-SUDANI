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
  listingRateHourly: 3,
  listingRateDaily: 6,
  adminPasswordIterations: 120000,
  paymentColumns: ['PaymentStatus', 'PaymentAmount', 'PaymentMethod', 'PaymentReference', 'PaymentSubmittedAt', 'PaymentVerifiedAt']
};

function doGet(e) {
  // Public read-only JSONP bridge for GitHub Pages.
  // Avoids hidden-iframe postMessage failures on some mobile browsers.
  var params = e && e.parameter ? e.parameter : {};
  var callback = String(params.callback || '');
  if (params.action === 'getPublicData' && /^daleelakJsonp_[A-Za-z0-9_]{8,80}$/.test(callback)) {
    var publicData;
    try {
      publicData = getPublicData();
    } catch (err) {
      publicData = { ok: false, error: err && err.message ? String(err.message) : 'تعذر تحميل بيانات الدليل.' };
    }
    var json = JSON.stringify(publicData)
      .replace(/</g, '\\u003c')
      .replace(/>/g, '\\u003e')
      .replace(/&/g, '\\u0026')
      .replace(/\u2028/g, '\\u2028')
      .replace(/\u2029/g, '\\u2029');
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

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
    if (payloadText.length > 1800000) throw new Error('حجم الصور كبير. قلّل حجم الصور وحاول مرة أخرى.');
    var payload = JSON.parse(payloadText || '{}');
    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
      throw new Error('صيغة الطلب غير صحيحة.');
    }

    switch (action) {
      case 'getPublicData': result = getPublicData(); break;
      case 'submitProviderListing': result = submitProviderListing(payload); break;
      case 'submitPaymentReference': result = submitPaymentReference(payload); break;
      case 'adminLogin': result = adminLogin(payload.password); break;
      case 'adminLogout': result = adminLogout(payload.token); break;
      case 'adminGetDashboard': result = adminGetDashboard(payload.token); break;
      case 'adminSetProviderStatus': result = adminSetProviderStatus(payload.token, payload.id, payload.status); break;
      case 'adminConfirmPaymentAndApprove': result = adminConfirmPaymentAndApprove(payload.token, payload.id); break;
      case 'adminSetProviderFeatured': result = adminSetProviderFeatured(payload.token, payload.id, payload.featured); break;
      case 'adminDeleteProvider': result = adminDeleteProvider(payload.token, payload.id); break;
      case 'adminClearListings': result = adminClearListings(payload.token); break;
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

  ensureProviderImageColumn_(providersSheet);
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
      images: parseProviderImages_(x.Images),
      createdAt: x.CreatedAt instanceof Date ? x.CreatedAt.getTime() : 0
    }))
    .sort((a, b) => Number(b.featured) - Number(a.featured) || b.createdAt - a.createdAt);

  return {
    ok: true, version: APP.version, title: APP.title, categories: categories,
    providers: approved, totalProviders: approved.length, listingFee: APP.listingFee,
    stcPayNumber: PropertiesService.getScriptProperties().getProperty('STCPAY_NUMBER') || '',
    bankIban: PropertiesService.getScriptProperties().getProperty('BANK_IBAN') || '',
    bankAccountName: PropertiesService.getScriptProperties().getProperty('BANK_ACCOUNT_NAME') || '',
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
  const phoneDigits = phone.replace(/\D/g, '');
  const elapsedMs = Number(form.formElapsedMs || 0);
  if (String(form.website || '').trim()) throw new Error('تعذر التحقق من الطلب. حدّث الصفحة وحاول مرة أخرى.');
  if (!Number.isFinite(elapsedMs) || elapsedMs < 2500 || elapsedMs > 7200000) throw new Error('يرجى مراجعة بيانات الإعلان ثم المحاولة مرة أخرى.');
  if (!name || !category || !city || !phone) throw new Error('يرجى تعبئة جميع البيانات المطلوبة للإعلان.');
  if (!/^[+\d\s()-]{7,25}$/.test(phone) || phoneDigits.length < 7) throw new Error('رقم التواصل غير صحيح. أدخل الرقم مع رمز الدولة إن أمكن.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getProvidersSheet_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const categorySheet = ss && ss.getSheetByName(APP.sheets.categories);
    if (!categorySheet) throw new Error('أقسام الخدمات غير متاحة حالياً.');
    const activeCategories = readObjects_(categorySheet)
      .filter(x => x.Active === true || String(x.Active).toLowerCase() === 'true')
      .map(x => String(x.Name || '').trim());
    if (activeCategories.indexOf(category) === -1) throw new Error('القسم المختار غير متاح. حدّث الصفحة واختر قسماً من القائمة.');
    enforceListingRateLimit_(phoneDigits);
    ensureProviderPaymentColumns_(sh);
    ensureProviderImageColumn_(sh);
    const headers = getHeaders_(sh);
    const id = Utilities.getUuid();
    let imageIds = [];
    const record = {
      ID: id, CreatedAt: new Date(), Status: 'Pending',
      Name: name, Category: category, City: city, Description: description,
      Phone: phone, WhatsApp: whatsapp, Featured: false, Images: JSON.stringify(imageIds),
      PaymentStatus: 'AwaitingPayment', PaymentAmount: APP.listingFee,
      PaymentMethod: '', PaymentReference: '', PaymentSubmittedAt: '', PaymentVerifiedAt: ''
    };
    try {
      imageIds = saveProviderImages_(form.images || [], id);
      record.Images = JSON.stringify(imageIds);
      sh.appendRow(headers.map(h => Object.prototype.hasOwnProperty.call(record, h) ? record[h] : ''));
    } catch (err) {
      trashProviderImages_(imageIds);
      throw err;
    }
    recordListingSubmission_(phoneDigits);
    return { ok: true, listingId: id, message: 'تم تسجيل بيانات إعلانك. أكمل الدفع في الصفحة التالية لإرسال الطلب للمراجعة.' };
  } finally { lock.releaseLock(); }
}

function submitPaymentReference(payload) {
  payload = payload || {};
  const id = clean_(payload.listingId, 100);
  const method = String(payload.paymentMethod || '');
  const reference = clean_(payload.paymentReference, 80);
  if (!id || ['BankTransfer', 'STCPay'].indexOf(method) === -1 || !reference) throw new Error('اختر طريقة الدفع وأدخل رقم العملية أو الرقم المرجعي.');
  if (!/^[\w\-\u0600-\u06FF ]{4,80}$/.test(reference)) throw new Error('رقم العملية غير صحيح. راجع الرقم كما يظهر في إيصال التحويل.');
  const props = PropertiesService.getScriptProperties();
  if (method === 'BankTransfer' && !props.getProperty('BANK_IBAN')) throw new Error('بيانات الحساب البنكي غير مكتملة حالياً.');
  if (method === 'STCPay' && !props.getProperty('STCPAY_NUMBER')) throw new Error('بيانات استقبال STC Pay غير متاحة حالياً.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getProvidersSheet_();
    ensureProviderPaymentColumns_(sh);
    const headers = getHeaders_(sh);
    const idCol = headers.indexOf('ID') + 1;
    const row = findProviderRow_(sh, idCol, id);
    if (!row) throw new Error('لم يتم العثور على الإعلان. أعد تسجيل الإعلان من فضلك.');
    const statusCol = headers.indexOf('Status') + 1;
    const paymentStatusCol = headers.indexOf('PaymentStatus') + 1;
    const existingStatus = String(sh.getRange(row, paymentStatusCol).getValue() || '');
    if (String(sh.getRange(row, statusCol).getValue() || '') !== 'Pending' || existingStatus !== 'AwaitingPayment') throw new Error('تم إرسال بيانات الدفع لهذا الإعلان مسبقاً أو لم يعد الطلب متاحاً.');
    const refCol = headers.indexOf('PaymentReference') + 1;
    const normalizedReference = normalizePaymentReference_(reference);
    const refs = sh.getLastRow() >= 2 ? sh.getRange(2, refCol, sh.getLastRow() - 1, 1).getDisplayValues().flat() : [];
    if (refs.some(x => normalizePaymentReference_(x) === normalizedReference)) throw new Error('رقم العملية مستخدم من قبل. تأكد من الرقم الصحيح.');
    sh.getRange(row, headers.indexOf('PaymentMethod') + 1).setValue(method);
    sh.getRange(row, refCol).setValue(reference);
    sh.getRange(row, paymentStatusCol).setValue('Submitted');
    sh.getRange(row, headers.indexOf('PaymentSubmittedAt') + 1).setValue(new Date());
    return { ok: true, message: 'تم إرسال بيانات الدفع بنجاح. طلبك الآن بانتظار مراجعة الإدارة.' };
  } finally { lock.releaseLock(); }
}

/* -------------------- ADMIN AUTHENTICATION -------------------- */

/**
 * First login migrates ADMIN_PASSWORD into a salted hash.
 * Set ADMIN_PASSWORD under Project Settings > Script Properties before use.
 */
function adminLogin(password) {
  const supplied = String(password || '');
  if (!supplied || supplied.length > 256) throw new Error('أدخل كلمة مرور الإدارة الصحيحة.');

  const props = PropertiesService.getScriptProperties();
  let salt = props.getProperty('ADMIN_PASSWORD_SALT');
  let savedHash = props.getProperty('ADMIN_PASSWORD_HASH');
  let hashVersion = props.getProperty('ADMIN_PASSWORD_HASH_VERSION') || 'LEGACY_SHA256';

  if (!savedHash) {
    const initialPassword = props.getProperty('ADMIN_PASSWORD');
    if (!initialPassword) {
      throw new Error('لم يتم إعداد كلمة مرور الإدارة. أضف ADMIN_PASSWORD في Script Properties.');
    }
    if (initialPassword.length < 12) {
      throw new Error('لأمان أفضل، اجعل كلمة مرور الإدارة 12 حرفاً على الأقل ثم أعد المحاولة.');
    }
    salt = Utilities.getUuid() + Utilities.getUuid();
    savedHash = pbkdf2Hash_(initialPassword, salt, APP.adminPasswordIterations);
    props.setProperties({
      ADMIN_PASSWORD_SALT: salt,
      ADMIN_PASSWORD_HASH: savedHash,
      ADMIN_PASSWORD_HASH_VERSION: 'PBKDF2_SHA256_120000'
    });
    props.deleteProperty('ADMIN_PASSWORD');
    hashVersion = 'PBKDF2_SHA256_120000';
  }

  if (!salt) throw new Error('إعداد كلمة مرور الإدارة غير مكتمل. راجع خصائص المشروع.');
  const candidateHash = hashVersion === 'PBKDF2_SHA256_120000'
    ? pbkdf2Hash_(supplied, salt, APP.adminPasswordIterations)
    : passwordHash_(supplied, salt);
  if (!constantTimeEqual_(candidateHash, savedHash)) throw new Error('كلمة المرور غير صحيحة.');

  if (hashVersion !== 'PBKDF2_SHA256_120000') {
    salt = Utilities.getUuid() + Utilities.getUuid();
    savedHash = pbkdf2Hash_(supplied, salt, APP.adminPasswordIterations);
    props.setProperties({
      ADMIN_PASSWORD_SALT: salt,
      ADMIN_PASSWORD_HASH: savedHash,
      ADMIN_PASSWORD_HASH_VERSION: 'PBKDF2_SHA256_120000'
    });
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
    paymentMethod: String(x.PaymentMethod || ''),
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
    const currentListingStatus = String(sh.getRange(row, statusCol).getValue() || '');
    if (currentListingStatus !== 'Pending' || currentPaymentStatus !== 'Submitted') {
      throw new Error('هذا الطلب ليس إعلاناً معلقاً يحتوي على إثبات دفع جديد.');
    }
    const amountCol = headers.indexOf('PaymentAmount') + 1;
    const amount = amountCol ? Number(sh.getRange(row, amountCol).getValue()) : 0;
    if (amount !== APP.listingFee) throw new Error('قيمة رسوم الإعلان المسجلة غير مطابقة للرسوم المعتمدة.');
    const refCol = headers.indexOf('PaymentReference') + 1;
    const reference = refCol ? String(sh.getRange(row, refCol).getValue() || '').trim() : '';
    if (!reference) throw new Error('لا يوجد رقم عملية مسجل لهذا الطلب.');
    const methodCol = headers.indexOf('PaymentMethod') + 1;
    const method = methodCol ? String(sh.getRange(row, methodCol).getValue() || '') : '';
    if (method !== 'BankTransfer' && method !== 'STCPay') throw new Error('طريقة الدفع المسجلة غير صحيحة.');

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

function adminClearListings(token) {
  requireAdmin_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getProvidersSheet_();
    const headers = getHeaders_(sh);
    const idCol = headers.indexOf('ID') + 1;
    const statusCol = headers.indexOf('Status') + 1;
    if (!idCol || !statusCol) throw new Error('أعمدة ورقة Providers غير مكتملة.');
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, deletedCount: 0 };

    const rowCount = lastRow - 1;
    const ids = sh.getRange(2, idCol, rowCount, 1).getValues();
    const statuses = sh.getRange(2, statusCol, rowCount, 1).getValues();
    let deletedCount = 0;
    const updated = statuses.map((cell, index) => {
      const id = String(ids[index][0] || '').trim();
      const status = String(cell[0] || '');
      if (id && status.toLowerCase() !== 'deleted') {
        deletedCount++;
        return ['Deleted'];
      }
      return [cell[0]];
    });
    if (deletedCount) sh.getRange(2, statusCol, rowCount, 1).setValues(updated);
    return { ok: true, deletedCount: deletedCount };
  } finally {
    lock.releaseLock();
  }
}


function adminDeleteProvider(token, id) {
  requireAdmin_(token);
  id = clean_(id, 100);
  if (!id) throw new Error('معرّف الإعلان غير صحيح.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getProvidersSheet_();
    const headers = getHeaders_(sh);
    const idCol = headers.indexOf('ID') + 1;
    const statusCol = headers.indexOf('Status') + 1;
    if (!idCol || !statusCol) throw new Error('أعمدة ورقة Providers غير مكتملة.');

    const row = findProviderRow_(sh, idCol, id);
    if (!row) throw new Error('الإعلان غير موجود أو تم حذفه مسبقاً.');

    // Soft-delete keeps the row for audit/recovery but removes it from public/admin tabs.
    sh.getRange(row, statusCol).setValue('Deleted');
    return { ok: true, status: 'Deleted' };
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

function pbkdf2Hash_(password, salt, iterations) {
  const rounds = Math.max(10000, Math.min(Number(iterations) || 12000, 200000));
  const key = String(password);
  let u = Utilities.computeHmacSha256Signature(String(salt) + '\x00\x00\x00\x01', key, Utilities.Charset.UTF_8);
  const derived = u.slice();
  for (let i = 1; i < rounds; i++) {
    u = Utilities.computeHmacSha256Signature(u, key);
    for (let j = 0; j < derived.length; j++) derived[j] = derived[j] ^ u[j];
  }
  return Utilities.base64Encode(derived);
}

function constantTimeEqual_(left, right) {
  left = String(left || ''); right = String(right || '');
  if (left.length !== right.length) return false;
  let difference = 0;
  for (let i = 0; i < left.length; i++) difference |= left.charCodeAt(i) ^ right.charCodeAt(i);
  return difference === 0;
}

function normalizePaymentReference_(value) {
  return String(value || '').replace(/\s+/g, '').toUpperCase();
}

function listingRateKeys_(phoneDigits) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(phoneDigits), Utilities.Charset.UTF_8);
  const key = Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, '');
  return { hourly: 'listing_h_' + key, daily: 'listing_d_' + key };
}

function enforceListingRateLimit_(phoneDigits) {
  const cache = CacheService.getScriptCache();
  const keys = listingRateKeys_(phoneDigits);
  const hourly = Number(cache.get(keys.hourly) || 0);
  const daily = Number(cache.get(keys.daily) || 0);
  if (hourly >= APP.listingRateHourly) throw new Error('تم الوصول للحد المؤقت لإضافة الإعلانات من هذا الرقم. حاول بعد ساعة.');
  if (daily >= APP.listingRateDaily) throw new Error('تم الوصول للحد اليومي لإضافة الإعلانات من هذا الرقم. حاول غداً.');
}

function recordListingSubmission_(phoneDigits) {
  const cache = CacheService.getScriptCache();
  const keys = listingRateKeys_(phoneDigits);
  const hourly = Number(cache.get(keys.hourly) || 0);
  const daily = Number(cache.get(keys.daily) || 0);
  cache.put(keys.hourly, String(hourly + 1), 3600);
  cache.put(keys.daily, String(daily + 1), 86400);
}

function trashProviderImages_(ids) {
  (ids || []).forEach(id => {
    try { DriveApp.getFileById(String(id)).setTrashed(true); } catch (err) {}
  });
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

function ensureProviderImageColumn_(sh) {
  const headers = getHeaders_(sh);
  if (headers.indexOf('Images') === -1) {
    sh.getRange(1, sh.getLastColumn() + 1).setValue('Images').setFontWeight('bold');
  }
}

function authorizeProviderImageStorage() {
  const props = PropertiesService.getScriptProperties();
  const key = 'PROVIDER_IMAGES_FOLDER_ID';
  let folder = null;
  const savedId = props.getProperty(key);
  if (savedId) {
    try { folder = DriveApp.getFolderById(savedId); } catch (err) {}
  }
  if (!folder) {
    folder = DriveApp.createFolder('دليلك السوداني - صور الإعلانات');
    props.setProperty(key, folder.getId());
  }
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'تم تجهيز مجلد صور إعلانات دليلك السوداني.';
}

function parseProviderImages_(value) {
  try {
    const ids = Array.isArray(value) ? value : JSON.parse(String(value || '[]'));
    if (!Array.isArray(ids)) return [];
    return ids.slice(0, 4).map(id => String(id || '').trim())
      .filter(id => /^[A-Za-z0-9_-]{10,}$/.test(id))
      .map(id => 'https://drive.google.com/uc?export=view&id=' + encodeURIComponent(id));
  } catch (err) { return []; }
}

function saveProviderImages_(images, listingId) {
  if (!Array.isArray(images)) throw new Error('صيغة الصور غير صحيحة.');
  if (images.length > 4) throw new Error('يمكن إضافة 4 صور كحد أقصى لكل إعلان.');
  if (!images.length) return [];
  const key = 'PROVIDER_IMAGES_FOLDER_ID';
  const props = PropertiesService.getScriptProperties();
  let folder = null;
  const savedFolderId = props.getProperty(key);
  if (savedFolderId) {
    try { folder = DriveApp.getFolderById(savedFolderId); } catch (err) {}
  }
  if (!folder) {
    folder = DriveApp.createFolder('دليلك السوداني - صور الإعلانات');
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    props.setProperty(key, folder.getId());
  }
  const prepared = images.map((item, index) => {
    if (!item || typeof item.data !== 'string') throw new Error('إحدى الصور غير صالحة.');
    const match = item.data.match(/^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=]+)$/);
    if (!match) throw new Error('يسمح فقط بصور JPG أو PNG أو WEBP.');
    if (match[2].length > 430000) throw new Error('حجم إحدى الصور كبير. اختر صوراً أصغر.');
    const bytes = Utilities.base64Decode(match[2]);
    if (bytes.length > 320000) throw new Error('يجب ألا يتجاوز حجم الصورة الواحدة 320 كيلوبايت.');
    const b = bytes.map(x => x & 255);
    const mime = match[1];
    const validJpeg = mime === 'image/jpeg' && b[0] === 255 && b[1] === 216 && b[2] === 255;
    const validPng = mime === 'image/png' && b.slice(0, 8).join(',') === '137,80,78,71,13,10,26,10';
    const validWebp = mime === 'image/webp' && String.fromCharCode.apply(null, b.slice(0, 4)) === 'RIFF' && String.fromCharCode.apply(null, b.slice(8, 12)) === 'WEBP';
    if (!validJpeg && !validPng && !validWebp) throw new Error('محتوى إحدى الصور لا يطابق نوع الملف المسموح.');
    const ext = mime === 'image/png' ? 'png' : (mime === 'image/webp' ? 'webp' : 'jpg');
    return { bytes: bytes, mime: mime, name: listingId + '-' + (index + 1) + '.' + ext };
  });
  const ids = [];
  try {
    prepared.forEach(item => {
      const blob = Utilities.newBlob(item.bytes, item.mime, item.name);
      const file = folder.createFile(blob);
      ids.push(file.getId());
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    });
    return ids;
  } catch (err) {
    trashProviderImages_(ids);
    throw err;
  }
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

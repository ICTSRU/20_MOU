/**
 * منصة توقيع مذكرات التفاهم — SRU MoU Signing
 * Google Apps Script backend · v1.7
 *
 * Security model
 *  - Encryption/decryption happens ONLY in the user's browser (Web Crypto).
 *  - This script never receives a PIN or a plain signature image. It stores:
 *      • the signature ciphertext (AES-256-GCM) as a file in a private Drive folder
 *      • an auth hash = SHA-256( PIN-derived tag + server pepper )
 *  - Ciphertext is released only after the PIN-derived tag is verified,
 *    with server-side lockout after MAX_FAILS wrong attempts.
 *
 * Access (two deployments of the SAME script)
 *  - University deployment: "Anyone within Sulaiman Al Rajhi University" → Google identity (@sr.edu.sa).
 *  - Guest deployment:      "Anyone"  → external users listed in the Externals sheet sign in
 *                           with a one-time code sent to their e-mail.
 */

const APP_VERSION    = '1.7';
const ALLOWED_DOMAIN = 'sr.edu.sa';
const MAX_FAILS      = 5;
const LOCK_MINUTES   = 15;
const TZ             = 'Asia/Riyadh';
const SESSION_HOURS  = 6;     // external session length (CacheService max = 6h)
const OTP_MINUTES    = 10;
const MAX_LOGO_CHARS = 45000; // a logo must fit in one Sheet cell

const SH = { SIGNERS:'Signers', DOCS:'Documents', SIGNINGS:'SigningLog', ACCESS:'AccessLog', ADMINS:'Admins', EXT:'Externals' };
const HEAD = {
  Signers:    ['id','name','position','org','fileId','encSalt','encIv','iter','authSalt','authHash','fails','lockUntil','status','createdBy','createdAt','updatedAt','ownerEmail'],
  Documents:  ['docId','refNo','title','status','json','createdBy','createdAt','updatedBy','updatedAt','logoRight','logoLeft','externalEmails'],
  SigningLog: ['timestamp','date','time','email','action','signerId','signerName','docRef','party','docHash','result'],
  AccessLog:  ['email','date','time','action'],
  Admins:     ['email','name','addedAt'],
  Externals:  ['email','name','org','status','addedBy','addedAt','lastLogin']
};
const TEXT_SHEETS = [SH.SIGNERS, SH.DOCS];
const PURPOSES = ['sign','view','preview','unsign','pin-change'];

/* ======================= Setup & migration ======================= */
/** Run once from the editor (also re-run after an update to approve new permissions). */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEAD).forEach(n => {
    const sh = ss.getSheetByName(n) || ss.insertSheet(n);
    const h = HEAD[n];
    if (sh.getLastRow() === 0) sh.appendRow(h);
    else {
      // Migration: add any new columns to the header row (existing columns keep their position)
      const cur = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
      h.forEach((k, i) => { if (cur[i] !== k && cur.indexOf(k) === -1 && !cur[i]) sh.getRange(1, i + 1).setValue(k); });
    }
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#501e8c').setFontColor('#ffffff');
    if (TEXT_SHEETS.indexOf(n) > -1) sh.getRange(1, 1, sh.getMaxRows(), h.length).setNumberFormat('@');
  });
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('FOLDER_ID')) props.setProperty('FOLDER_ID', DriveApp.createFolder('SRU-MoU-Signatures-Encrypted').getId());
  if (!props.getProperty('PEPPER'))    props.setProperty('PEPPER', Utilities.getUuid() + Utilities.getUuid());

  const me = norm_(Session.getEffectiveUser().getEmail());
  if (!rows_(SH.ADMINS).some(r => norm_(r.email) === me)) append_(SH.ADMINS, { email: me, name: 'Owner', addedAt: new Date().toISOString() });

  ['Sheet1', 'ورقة1'].forEach(name => { const s = ss.getSheetByName(name); if (s && ss.getSheets().length > 1 && s.getLastRow() === 0) ss.deleteSheet(s); });
  props.setProperty('SCHEMA', APP_VERSION);
  return 'Setup complete ✔ v' + APP_VERSION;
}

/** Self-healing: runs setup() if a sheet/property is missing or the schema is older than this version. */
let ensured_ = false;
function ensure_() {
  if (ensured_) return;
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('السكربت غير مرتبط بملف Google Sheet — أنشئه من Extensions → Apps Script داخل الملف');
  const props = PropertiesService.getScriptProperties();
  const missing = Object.keys(HEAD).some(n => !ss.getSheetByName(n)) || !props.getProperty('FOLDER_ID') || !props.getProperty('PEPPER') || props.getProperty('SCHEMA') !== APP_VERSION;
  if (missing) setup();
  ensured_ = true;
}

/* ======================= Web app entry ======================= */
function doGet() {
  ensure_();
  const g = googleEmail_();
  const t = HtmlService.createTemplateFromFile('Index');
  if (g) {
    logAccess_(g, 'sign-in');
    t.boot = JSON.stringify(bootFor_({ email: g, external: false }));
  } else {
    // Not an SRU Google session → guest sign-in screen (works only on the "Anyone" deployment)
    t.boot = JSON.stringify({ needsLogin: true, version: APP_VERSION });
  }
  return t.evaluate()
    .setTitle('منصة توقيع مذكرات التفاهم v' + APP_VERSION)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function bootFor_(id) {
  const ext = id.external ? rows_(SH.EXT).find(r => norm_(r.email) === id.email) : null;
  return {
    email: id.email, external: !!id.external, isAdmin: !id.external && isAdmin_(id.email),
    name: ext ? ext.name : '', org: ext ? ext.org : '',
    version: APP_VERSION, maxFails: MAX_FAILS, lockMin: LOCK_MINUTES
  };
}

/* ======================= Guest (external) sign-in ======================= */
/** Step 1: send a one-time code. Always answers the same way so the list of guests can't be probed. */
function extRequestCode(email) {
  ensure_();
  email = norm_(email);
  const generic = { ok: true, msg: 'إذا كان البريد مسجّلاً لدينا فسيصلك رمز الدخول خلال دقيقة.' };
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || !isExternalActive_(email)) { logAccess_(email || '(empty)', 'otp-denied'); return generic; }
  const c = CacheService.getScriptCache();
  const rk = 'otpn:' + email, n = Number(c.get(rk) || 0);
  if (n >= 5) { logAccess_(email, 'otp-rate-limited'); return generic; }
  c.put(rk, String(n + 1), 3600);
  const code = String(parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 12), 16) % 900000 + 100000);
  c.put('otp:' + email, JSON.stringify({ h: sha_(code), tries: 0 }), OTP_MINUTES * 60);
  MailApp.sendEmail({
    to: email,
    subject: 'رمز الدخول — منصة توقيع مذكرات التفاهم | Sulaiman Al Rajhi University',
    htmlBody:
      '<div dir="rtl" style="font-family:Tahoma,Arial;color:#241436">' +
      '<h3 style="color:#3a1464">منصة توقيع مذكرات التفاهم</h3>' +
      '<p>رمز الدخول الخاص بك:</p>' +
      '<p style="font-size:28px;font-weight:bold;letter-spacing:6px;color:#501e8c">' + code + '</p>' +
      '<p>صالح لمدة ' + OTP_MINUTES + ' دقائق ولمرة واحدة. إذا لم تطلب هذا الرمز فتجاهل الرسالة.</p>' +
      '<hr><p style="font-size:12px;color:#6b5a80">Your sign-in code is <b>' + code + '</b> (valid ' + OTP_MINUTES + ' minutes).<br>' +
      'Sulaiman Al Rajhi University — Executive Directorate of Communications &amp; IT</p></div>'
  });
  logAccess_(email, 'otp-sent');
  return generic;
}

/** Step 2: verify the code and open a session (token kept in the browser tab only). */
function extVerifyCode(email, code) {
  ensure_();
  email = norm_(email);
  const c = CacheService.getScriptCache(), key = 'otp:' + email, raw = c.get(key);
  if (!raw) return { ok: false, msg: 'انتهت صلاحية الرمز أو لم يُطلب — اطلب رمزاً جديداً.' };
  const o = JSON.parse(raw);
  if (sha_(String(code || '').trim()) !== o.h) {
    o.tries++;
    if (o.tries >= 5) { c.remove(key); logAccess_(email, 'otp-locked'); return { ok: false, msg: 'تجاوزت عدد المحاولات — اطلب رمزاً جديداً.' }; }
    c.put(key, JSON.stringify(o), OTP_MINUTES * 60);
    logAccess_(email, 'otp-failed');
    return { ok: false, msg: 'رمز غير صحيح — المحاولات المتبقية: ' + (5 - o.tries) };
  }
  c.remove(key);
  if (!isExternalActive_(email)) return { ok: false, msg: 'الحساب غير مفعّل.' };
  const token = Utilities.getUuid() + Utilities.getUuid();
  c.put('tok:' + token, email, SESSION_HOURS * 3600);
  const r = rows_(SH.EXT).find(x => norm_(x.email) === email);
  if (r) update_(SH.EXT, r._row, { lastLogin: new Date().toISOString() });
  logAccess_(email, 'sign-in-external');
  return { ok: true, token, boot: bootFor_({ email, external: true }) };
}

/* ======================= API ======================= */
/** Single API endpoint: google.script.run.api(action, payload, token) */
function api(action, p, token) {
  ensure_();
  const me = identity_(token);
  p = p || {};
  const writes = ['addSigner','unlock','changePin','deleteSigner','saveDoc','addAdmin','removeAdmin','addExternal','setExternalStatus','removeExternal','setGuestUrl','inviteExternal'];
  const lock = LockService.getScriptLock();
  const w = writes.indexOf(action) > -1;
  if (w) lock.waitLock(20000);
  try {
    switch (action) {
      case 'listSigners':       return listSigners_(me);
      case 'addSigner':         return addSigner_(me, p);
      case 'unlock':            return unlock_(me, p);
      case 'changePin':         return changePin_(me, p);
      case 'deleteSigner':      return deleteSigner_(me, p);
      case 'listDocs':          return listDocs_(me);
      case 'loadDoc':           return loadDoc_(me, p.docId);
      case 'saveDoc':           return saveDoc_(me, p);
      // ---- admin ----
      case 'listUsers':         return admin_(me) && listUsers_();
      case 'addAdmin':          return admin_(me) && addAdmin_(me, p);
      case 'removeAdmin':       return admin_(me) && removeAdmin_(me, p);
      case 'addExternal':       return admin_(me) && addExternal_(me, p);
      case 'setExternalStatus': return admin_(me) && setExternalStatus_(me, p);
      case 'removeExternal':    return admin_(me) && removeExternal_(me, p);
      case 'setGuestUrl':       return admin_(me) && setGuestUrl_(p);
      case 'inviteExternal':    return admin_(me) && inviteExternal_(me, p);
      default: throw new Error('إجراء غير معروف');
    }
  } finally { if (w) lock.releaseLock(); }
}

function identity_(token) {
  const g = googleEmail_();
  if (g) return { email: g, external: false, isAdmin: isAdmin_(g) };
  if (token) {
    const e = CacheService.getScriptCache().get('tok:' + token);
    if (e && isExternalActive_(e)) return { email: e, external: true, isAdmin: false };
  }
  throw new Error('انتهت الجلسة أو غير مصرّح — أعد تسجيل الدخول.');
}
function admin_(me) { if (!me.isAdmin) throw new Error('هذه الصلاحية لمسؤولي المنصة فقط'); return true; }

/* ======================= Signers ======================= */
function listSigners_(me) {
  const now = Date.now();
  return rows_(SH.SIGNERS)
    .filter(r => r.status === 'active' && (!me.external || norm_(r.ownerEmail) === me.email))
    .map(r => ({
      id: r.id, name: r.name, position: r.position, org: r.org,
      authSalt: r.authSalt, iter: Number(r.iter), createdAt: String(r.createdAt),
      locked: Number(r.lockUntil) > now,
      ownerEmail: me.isAdmin ? String(r.ownerEmail || '') : '',
      mine: norm_(r.ownerEmail) === me.email
    }));
}

function addSigner_(me, p) {
  const name = clean_(p.name, 120);
  if (!name) throw new Error('الاسم مطلوب');
  // Admins may register a signature for anyone; everyone else only for themselves.
  const owner = me.isAdmin ? norm_(p.ownerEmail || '') : me.email;
  if (owner && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(owner)) throw new Error('بريد صاحب التوقيع غير صالح');
  validEnvelope_(p.enc, p.auth);

  const id = Utilities.getUuid();
  const file = DriveApp.getFolderById(prop_('FOLDER_ID')).createFile(id + '.enc', p.enc.ct, MimeType.PLAIN_TEXT);
  const now = new Date().toISOString();
  append_(SH.SIGNERS, {
    id, name, position: clean_(p.position, 120), org: clean_(p.org, 160), fileId: file.getId(),
    encSalt: p.enc.salt, encIv: p.enc.iv, iter: String(p.enc.iter),
    authSalt: p.auth.salt, authHash: hashTag_(p.auth.tag),
    fails: '0', lockUntil: '0', status: 'active', createdBy: me.email, createdAt: now, updatedAt: now, ownerEmail: owner
  });
  logSign_(me.email, { action: 'add-signer', signerId: id, signerName: name, result: 'success' });
  return { ok: true, id };
}

function unlock_(me, p) {
  const r = findSigner_(p.id, me);
  const purpose = PURPOSES.indexOf(p.purpose) > -1 ? p.purpose : 'view';
  const ctx = p.ctx || {};
  const v = verify_(r, p.tag);
  logSign_(me.email, {
    action: purpose, signerId: r.id, signerName: r.name,
    docRef: clean_(ctx.docRef, 60), party: clean_(ctx.party, 10), docHash: clean_(ctx.docHash, 64),
    result: v.ok ? 'success' : (v.locked ? 'locked' : 'failed')
  });
  if (!v.ok) return v;
  const ct = DriveApp.getFileById(r.fileId).getBlob().getDataAsString();
  return { ok: true, enc: { salt: r.encSalt, iv: r.encIv, ct, iter: Number(r.iter) } };
}

function changePin_(me, p) {
  const r = findSigner_(p.id, me);
  const v = verify_(r, p.tag);
  if (!v.ok) return v;
  validEnvelope_(p.enc, p.auth);
  DriveApp.getFileById(r.fileId).setContent(p.enc.ct);
  update_(SH.SIGNERS, r._row, {
    encSalt: p.enc.salt, encIv: p.enc.iv, iter: String(p.enc.iter),
    authSalt: p.auth.salt, authHash: hashTag_(p.auth.tag), updatedAt: new Date().toISOString()
  });
  logSign_(me.email, { action: 'pin-changed', signerId: r.id, signerName: r.name, result: 'success' });
  return { ok: true };
}

function deleteSigner_(me, p) {
  const r = findSigner_(p.id, me);
  if (!me.isAdmin && norm_(r.ownerEmail) !== me.email) throw new Error('يمكنك حذف توقيعاتك فقط');
  try { DriveApp.getFileById(r.fileId).setTrashed(true); } catch (e) {}
  update_(SH.SIGNERS, r._row, { status: 'deleted', updatedAt: new Date().toISOString() });
  logSign_(me.email, { action: 'delete-signer', signerId: r.id, signerName: r.name, result: 'success' });
  return { ok: true };
}

function verify_(r, tag) {
  const now = Date.now();
  const until = Number(r.lockUntil) || 0;
  if (until > now) return { ok: false, locked: true, msg: 'التوقيع مقفل مؤقتاً — حاول بعد ' + Math.ceil((until - now) / 60000) + ' دقيقة' };
  if (!isB64_(tag, 100)) return { ok: false, msg: 'بيانات تحقق غير صالحة' };
  if (hashTag_(tag) === r.authHash) {
    if (Number(r.fails)) update_(SH.SIGNERS, r._row, { fails: '0', lockUntil: '0' });
    return { ok: true };
  }
  const fails = (Number(r.fails) || 0) + 1;
  if (fails >= MAX_FAILS) {
    update_(SH.SIGNERS, r._row, { fails: '0', lockUntil: String(now + LOCK_MINUTES * 60000) });
    return { ok: false, locked: true, msg: 'تجاوزت عدد المحاولات — تم قفل التوقيع ' + LOCK_MINUTES + ' دقيقة' };
  }
  update_(SH.SIGNERS, r._row, { fails: String(fails) });
  return { ok: false, msg: 'رقم سري غير صحيح — المحاولات المتبقية: ' + (MAX_FAILS - fails) };
}

/* ======================= Documents ======================= */
function docVisible_(me, r) {
  if (!me.external) return true;
  return String(r.externalEmails || '').split(',').map(norm_).indexOf(me.email) > -1;
}

function listDocs_(me) {
  return rows_(SH.DOCS).filter(r => docVisible_(me, r)).map(r => ({
    docId: r.docId, refNo: r.refNo, title: r.title, status: r.status,
    updatedBy: r.updatedBy, updatedAt: String(r.updatedAt)
  })).reverse();
}

function loadDoc_(me, docId) {
  const r = rows_(SH.DOCS).find(x => x.docId === docId);
  if (!r || !docVisible_(me, r)) throw new Error('المذكرة غير موجودة أو غير مصرّح بها');
  const d = JSON.parse(r.json);
  d.logoRight = String(r.logoRight || '');
  d.logoLeft  = String(r.logoLeft || '');
  return d;
}

function saveDoc_(me, p) {
  let d = p.doc || {};
  const existing = d.docId ? rows_(SH.DOCS).find(x => x.docId === d.docId) : null;

  if (me.external) {
    // Guests cannot create or edit documents — they may only add/remove THEIR OWN signature.
    if (!existing || !docVisible_(me, existing)) throw new Error('غير مصرّح');
    const stored = JSON.parse(existing.json);
    stored.signed = stored.signed || {}; stored.sel = stored.sel || {};
    ['p1', 'p2'].forEach(k => {
      const inc = d.signed && d.signed[k], old = stored.signed[k];
      const ownsInc = inc && ownsSigner_(me, inc.signerId), ownsOld = old && ownsSigner_(me, old.signerId);
      if (JSON.stringify(inc || null) === JSON.stringify(old || null)) return;
      if ((inc && ownsInc && !old) || (!inc && ownsOld)) { stored.signed[k] = inc || null; if (inc) stored.sel[k] = inc.signerId; }
      else throw new Error('يمكنك إضافة أو إزالة توقيعك فقط');
    });
    stored.logoRight = existing.logoRight; stored.logoLeft = existing.logoLeft;
    d = stored;
  }

  // Never store signature images: keep only signing metadata
  ['p1', 'p2'].forEach(k => { if (d.signed && d.signed[k]) delete d.signed[k].img; });
  const signedCount = ['p1', 'p2'].filter(k => d.signed && d.signed[k]).length;
  const status = signedCount === 2 ? 'signed' : signedCount === 1 ? 'partially-signed' : 'draft';
  const now = new Date().toISOString();
  d.docId = existing ? existing.docId : Utilities.getUuid();

  const logoRight = cleanLogo_(d.logoRight), logoLeft = cleanLogo_(d.logoLeft);
  delete d.logoRight; delete d.logoLeft;
  const emails = [d.p1email, d.p2email].map(norm_).filter(e => e && !e.endsWith('@' + ALLOWED_DOMAIN)).join(',');
  const json = JSON.stringify(d);
  if (json.length > 48000) throw new Error('نص المذكرة طويل جداً للحفظ (الحد ~48 ألف حرف)');

  const rec = { docId: d.docId, refNo: clean_(d.refNo, 60), title: clean_(d.title, 200), status, json,
                updatedBy: me.email, updatedAt: now, logoRight, logoLeft, externalEmails: emails };
  if (existing) update_(SH.DOCS, existing._row, rec);
  else append_(SH.DOCS, Object.assign(rec, { createdBy: me.email, createdAt: now }));
  return { ok: true, docId: d.docId, status };
}

function cleanLogo_(v) {
  v = String(v || '');
  if (['', 'none', 'sru', 'ictd', 'itoc'].indexOf(v) > -1) return v;
  if (/^data:image\/(png|jpeg|webp);base64,[A-Za-z0-9+/=]+$/.test(v) && v.length <= MAX_LOGO_CHARS) return v;
  throw new Error('صورة الشعار غير صالحة أو كبيرة جداً');
}

/* ======================= Users (admin) ======================= */
function listUsers_() {
  return {
    admins: rows_(SH.ADMINS).map(r => ({ email: String(r.email), name: String(r.name), addedAt: String(r.addedAt) })),
    externals: rows_(SH.EXT).filter(r => r.status !== 'removed').map(r => ({
      email: String(r.email), name: String(r.name), org: String(r.org), status: String(r.status),
      addedBy: String(r.addedBy), addedAt: String(r.addedAt), lastLogin: String(r.lastLogin || '')
    })),
    guestUrl: prop_('GUEST_URL') || ''
  };
}
function addAdmin_(me, p) {
  const email = norm_(p.email);
  if (!email.endsWith('@' + ALLOWED_DOMAIN)) throw new Error('المسؤول يجب أن يكون بحساب @' + ALLOWED_DOMAIN);
  if (rows_(SH.ADMINS).some(r => norm_(r.email) === email)) throw new Error('موجود مسبقاً');
  append_(SH.ADMINS, { email, name: clean_(p.name, 120), addedAt: new Date().toISOString() });
  logAccess_(me.email, 'admin-added:' + email);
  return { ok: true };
}
function removeAdmin_(me, p) {
  const email = norm_(p.email), all = rows_(SH.ADMINS);
  if (email === me.email) throw new Error('لا يمكنك إزالة نفسك');
  if (all.length <= 1) throw new Error('يجب بقاء مسؤول واحد على الأقل');
  const r = all.find(x => norm_(x.email) === email);
  if (!r) throw new Error('غير موجود');
  sheet_(SH.ADMINS).deleteRow(r._row);
  logAccess_(me.email, 'admin-removed:' + email);
  return { ok: true };
}
function addExternal_(me, p) {
  const email = norm_(p.email), name = clean_(p.name, 120);
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('البريد غير صالح');
  if (email.endsWith('@' + ALLOWED_DOMAIN)) throw new Error('منسوبو الجامعة لا يحتاجون إضافة — يدخلون بحسابهم مباشرة');
  if (!name) throw new Error('الاسم مطلوب');
  const r = rows_(SH.EXT).find(x => norm_(x.email) === email);
  const rec = { email, name, org: clean_(p.org, 160), status: 'active', addedBy: me.email, addedAt: new Date().toISOString() };
  if (r) update_(SH.EXT, r._row, rec); else append_(SH.EXT, rec);
  logAccess_(me.email, 'external-added:' + email);
  return { ok: true };
}
function setExternalStatus_(me, p) {
  const r = rows_(SH.EXT).find(x => norm_(x.email) === norm_(p.email));
  if (!r) throw new Error('غير موجود');
  const status = p.status === 'active' ? 'active' : 'disabled';
  update_(SH.EXT, r._row, { status });
  logAccess_(me.email, 'external-' + status + ':' + r.email);
  return { ok: true };
}
function removeExternal_(me, p) {
  const r = rows_(SH.EXT).find(x => norm_(x.email) === norm_(p.email));
  if (!r) throw new Error('غير موجود');
  update_(SH.EXT, r._row, { status: 'removed' });
  logAccess_(me.email, 'external-removed:' + r.email);
  return { ok: true };
}
function setGuestUrl_(p) {
  const u = String(p.url || '').trim();
  if (u && !/^https:\/\/script\.google\.com\/(a\/macros\/[^/]+|macros)\/s\/[\w-]+\/exec$/.test(u)) throw new Error('رابط غير صالح — يجب أن ينتهي بـ ‎/exec');
  PropertiesService.getScriptProperties().setProperty('GUEST_URL', u);
  return { ok: true };
}
function inviteExternal_(me, p) {
  const r = rows_(SH.EXT).find(x => norm_(x.email) === norm_(p.email) && x.status === 'active');
  if (!r) throw new Error('المستخدم غير موجود أو غير مفعّل');
  const url = prop_('GUEST_URL');
  if (!url) throw new Error('اضبط رابط دخول الضيوف أولاً');
  MailApp.sendEmail({
    to: r.email,
    subject: 'دعوة للدخول — منصة توقيع مذكرات التفاهم | Sulaiman Al Rajhi University',
    htmlBody:
      '<div dir="rtl" style="font-family:Tahoma,Arial;color:#241436">' +
      '<p>السلام عليكم ' + esc_(r.name) + '،</p>' +
      '<p>تمت إضافتكم إلى منصة توقيع مذكرات التفاهم بجامعة سليمان الراجحي. للدخول افتح الرابط التالي وأدخل بريدك الإلكتروني، وسيصلك رمز دخول لمرة واحدة:</p>' +
      '<p><a href="' + url + '" style="background:#501e8c;color:#fff;padding:10px 18px;border-radius:4px;text-decoration:none">الدخول إلى المنصة</a></p>' +
      '<hr><p style="font-size:12px;color:#6b5a80">You have been invited to the Sulaiman Al Rajhi University MoU Signing Platform. Open the link above and enter your e-mail to receive a one-time sign-in code.</p></div>'
  });
  logAccess_(me.email, 'external-invited:' + r.email);
  return { ok: true };
}

/* ======================= Helpers ======================= */
function googleEmail_() {
  const e = norm_(Session.getActiveUser().getEmail());
  return e && e.endsWith('@' + ALLOWED_DOMAIN) ? e : '';
}
function isAdmin_(e) { return rows_(SH.ADMINS).some(r => norm_(r.email) === e); }
function isExternalActive_(e) { return rows_(SH.EXT).some(r => norm_(r.email) === norm_(e) && r.status === 'active'); }
function ownsSigner_(me, id) { return rows_(SH.SIGNERS).some(r => r.id === id && norm_(r.ownerEmail) === me.email); }
function findSigner_(id, me) {
  const r = rows_(SH.SIGNERS).find(x => x.id === id && x.status === 'active');
  if (!r) throw new Error('التوقيع غير موجود');
  if (me && me.external && norm_(r.ownerEmail) !== me.email) throw new Error('غير مصرّح باستخدام هذا التوقيع');
  return r;
}
function hashTag_(tag) { return sha_(tag + '|' + prop_('PEPPER')); }
function sha_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}
function validEnvelope_(enc, auth) {
  if (!enc || !auth) throw new Error('بيانات التشفير ناقصة');
  if (!isB64_(enc.salt, 64) || !isB64_(enc.iv, 64) || !isB64_(enc.ct, 1500000) || !isB64_(auth.salt, 64) || !isB64_(auth.tag, 100))
    throw new Error('بيانات التشفير غير صالحة');
  const it = Number(enc.iter);
  if (!(it >= 100000 && it <= 2000000)) throw new Error('إعدادات التشفير غير صالحة');
}
function isB64_(s, max) { return typeof s === 'string' && s.length > 0 && s.length <= max && /^[A-Za-z0-9+/=]+$/.test(s); }
function clean_(s, max) { return String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max); }
function esc_(s) { return String(s || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function norm_(s) { return String(s || '').trim().toLowerCase(); }
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function sheet_(n) { return SpreadsheetApp.getActive().getSheetByName(n); }
function rows_(n) {
  const v = sheet_(n).getDataRange().getValues();
  const h = (v.shift() || []).map(String);
  return v.map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = r[j]); return o; });
}
/** Writes by header name, so column order in the sheet does not matter. */
function headerIndex_(n) { const sh = sheet_(n); return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String); }
function append_(n, o) {
  const h = headerIndex_(n);
  sheet_(n).appendRow(h.map(k => o[k] === undefined ? '' : o[k]));
}
function update_(n, row, o) {
  const h = headerIndex_(n), rg = sheet_(n).getRange(row, 1, 1, h.length), cur = rg.getValues()[0];
  h.forEach((k, j) => { if (k in o) cur[j] = o[k]; });
  rg.setValues([cur]);
}
function stamp_() { const d = new Date(); return { ts: d, date: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'), time: Utilities.formatDate(d, TZ, 'HH:mm:ss') }; }
function logAccess_(email, action) { const s = stamp_(); append_(SH.ACCESS, { email, date: s.date, time: s.time, action }); }
function logSign_(email, o) { const s = stamp_(); append_(SH.SIGNINGS, Object.assign({ timestamp: s.ts, date: s.date, time: s.time, email }, o)); }

/**
 * منصة توقيع مذكرات التفاهم — SRU MoU Signing
 * Google Apps Script backend · v1.5
 *
 * Security model
 *  - Encryption/decryption happens ONLY in the user's browser (Web Crypto).
 *  - This script never receives a PIN or a plain signature image. It stores:
 *      • the signature ciphertext (AES-256-GCM) as a file in a private Drive folder
 *      • an auth hash = SHA-256( PIN-derived tag + server pepper )
 *  - The ciphertext is released only after the PIN-derived tag is verified,
 *    with server-side lockout after MAX_FAILS wrong attempts.
 *  - Web app runs "as me" for "Anyone in sr.edu.sa", so users never get
 *    direct access to the Sheet or the Drive folder.
 */

const APP_VERSION   = '1.5';
const ALLOWED_DOMAIN = 'sr.edu.sa';
const MAX_FAILS     = 5;
const LOCK_MINUTES  = 15;
const TZ            = 'Asia/Riyadh';

const SH = { SIGNERS:'Signers', DOCS:'Documents', SIGNINGS:'SigningLog', ACCESS:'AccessLog', ADMINS:'Admins' };
const HEAD = {
  Signers:    ['id','name','position','org','fileId','encSalt','encIv','iter','authSalt','authHash','fails','lockUntil','status','createdBy','createdAt','updatedAt'],
  Documents:  ['docId','refNo','title','status','json','createdBy','createdAt','updatedBy','updatedAt'],
  SigningLog: ['timestamp','date','time','email','action','signerId','signerName','docRef','party','docHash','result'],
  AccessLog:  ['email','date','time','action'],
  Admins:     ['email','name','addedAt']
};
const PURPOSES = ['sign','view','preview','unsign','pin-change'];

/* ======================= One-time setup ======================= */
/** Run once from the Apps Script editor: creates sheets, Drive folder, secret pepper, and makes you admin. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEAD).forEach(n => {
    const sh = ss.getSheetByName(n) || ss.insertSheet(n);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEAD[n]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEAD[n].length).setFontWeight('bold').setBackground('#501e8c').setFontColor('#ffffff');
    }
    if (n === SH.SIGNERS || n === SH.DOCS) sh.getRange('A:P').setNumberFormat('@'); // keep base64/IDs as plain text
  });
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('FOLDER_ID')) props.setProperty('FOLDER_ID', DriveApp.createFolder('SRU-MoU-Signatures-Encrypted').getId());
  if (!props.getProperty('PEPPER'))    props.setProperty('PEPPER', Utilities.getUuid() + Utilities.getUuid());

  const me = norm_(Session.getEffectiveUser().getEmail());
  if (!rows_(SH.ADMINS).some(r => norm_(r.email) === me)) append_(SH.ADMINS, { email: me, name: 'Owner', addedAt: new Date().toISOString() });

  ['Sheet1', 'ورقة1'].forEach(name => { const s = ss.getSheetByName(name); if (s && ss.getSheets().length > 1 && s.getLastRow() === 0) ss.deleteSheet(s); });
  return 'Setup complete ✔';
}

/* ======================= Web app entry ======================= */
function doGet() {
  let email;
  try { email = email_(); }
  catch (err) {
    return HtmlService.createHtmlOutput(
      '<div dir="rtl" style="font-family:Tahoma;padding:48px;text-align:center;color:#3a1464">' +
      '<h2>غير مصرّح بالدخول</h2><p>يجب تسجيل الدخول بحساب البريد الجامعي <b>@' + ALLOWED_DOMAIN + '</b></p></div>'
    ).setTitle('غير مصرّح');
  }
  logAccess_(email, 'sign-in');
  const t = HtmlService.createTemplateFromFile('Index');
  t.boot = JSON.stringify({ email, isAdmin: isAdmin_(email), version: APP_VERSION, maxFails: MAX_FAILS, lockMin: LOCK_MINUTES });
  return t.evaluate()
    .setTitle('منصة توقيع مذكرات التفاهم v' + APP_VERSION)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Single API endpoint called from the page via google.script.run.api(action, payload) */
function api(action, p) {
  const email = email_();
  p = p || {};
  const writes = ['addSigner', 'unlock', 'changePin', 'deleteSigner', 'saveDoc'];
  const lock = LockService.getScriptLock();
  if (writes.indexOf(action) > -1) lock.waitLock(20000);
  try {
    switch (action) {
      case 'listSigners':  return listSigners_();
      case 'addSigner':    return addSigner_(email, p);
      case 'unlock':       return unlock_(email, p);
      case 'changePin':    return changePin_(email, p);
      case 'deleteSigner': return deleteSigner_(email, p);
      case 'listDocs':     return listDocs_();
      case 'loadDoc':      return loadDoc_(p.docId);
      case 'saveDoc':      return saveDoc_(email, p);
      default: throw new Error('إجراء غير معروف');
    }
  } finally {
    if (writes.indexOf(action) > -1) lock.releaseLock();
  }
}

/* ======================= Signers ======================= */
function listSigners_() {
  const now = Date.now();
  return rows_(SH.SIGNERS).filter(r => r.status === 'active').map(r => ({
    id: r.id, name: r.name, position: r.position, org: r.org,
    authSalt: r.authSalt, iter: Number(r.iter), createdAt: String(r.createdAt),
    locked: Number(r.lockUntil) > now
  }));
}

function addSigner_(email, p) {
  if (!isAdmin_(email)) throw new Error('إضافة التوقيعات متاحة لمسؤولي المنصة فقط');
  const name = clean_(p.name, 120);
  if (!name) throw new Error('الاسم مطلوب');
  validEnvelope_(p.enc, p.auth);

  const id = Utilities.getUuid();
  const folder = DriveApp.getFolderById(prop_('FOLDER_ID'));
  const file = folder.createFile(id + '.enc', p.enc.ct, MimeType.PLAIN_TEXT);
  const now = new Date().toISOString();
  append_(SH.SIGNERS, {
    id, name, position: clean_(p.position, 120), org: clean_(p.org, 160), fileId: file.getId(),
    encSalt: p.enc.salt, encIv: p.enc.iv, iter: String(p.enc.iter),
    authSalt: p.auth.salt, authHash: hashTag_(p.auth.tag),
    fails: '0', lockUntil: '0', status: 'active', createdBy: email, createdAt: now, updatedAt: now
  });
  logSign_(email, { action: 'add-signer', signerId: id, signerName: name, result: 'success' });
  return { ok: true, id };
}

function unlock_(email, p) {
  const r = findSigner_(p.id);
  const purpose = PURPOSES.indexOf(p.purpose) > -1 ? p.purpose : 'view';
  const ctx = p.ctx || {};
  const v = verify_(r, p.tag);
  logSign_(email, {
    action: purpose, signerId: r.id, signerName: r.name,
    docRef: clean_(ctx.docRef, 60), party: clean_(ctx.party, 10), docHash: clean_(ctx.docHash, 64),
    result: v.ok ? 'success' : (v.locked ? 'locked' : 'failed')
  });
  if (!v.ok) return v;
  const ct = DriveApp.getFileById(r.fileId).getBlob().getDataAsString();
  return { ok: true, enc: { salt: r.encSalt, iv: r.encIv, ct, iter: Number(r.iter) } };
}

function changePin_(email, p) {
  const r = findSigner_(p.id);
  const v = verify_(r, p.tag);
  if (!v.ok) return v;
  validEnvelope_(p.enc, p.auth);
  DriveApp.getFileById(r.fileId).setContent(p.enc.ct);
  update_(SH.SIGNERS, r._row, {
    encSalt: p.enc.salt, encIv: p.enc.iv, iter: String(p.enc.iter),
    authSalt: p.auth.salt, authHash: hashTag_(p.auth.tag), updatedAt: new Date().toISOString()
  });
  logSign_(email, { action: 'pin-changed', signerId: r.id, signerName: r.name, result: 'success' });
  return { ok: true };
}

function deleteSigner_(email, p) {
  if (!isAdmin_(email)) throw new Error('الحذف متاح لمسؤولي المنصة فقط');
  const r = findSigner_(p.id);
  try { DriveApp.getFileById(r.fileId).setTrashed(true); } catch (e) {}
  update_(SH.SIGNERS, r._row, { status: 'deleted', updatedAt: new Date().toISOString() });
  logSign_(email, { action: 'delete-signer', signerId: r.id, signerName: r.name, result: 'success' });
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
function listDocs_() {
  return rows_(SH.DOCS).map(r => ({
    docId: r.docId, refNo: r.refNo, title: r.title, status: r.status,
    updatedBy: r.updatedBy, updatedAt: String(r.updatedAt)
  })).reverse();
}

function loadDoc_(docId) {
  const r = rows_(SH.DOCS).find(x => x.docId === docId);
  if (!r) throw new Error('المذكرة غير موجودة');
  return JSON.parse(r.json);
}

function saveDoc_(email, p) {
  const d = p.doc || {};
  // Never store signature images: keep only signing metadata
  ['p1', 'p2'].forEach(k => { if (d.signed && d.signed[k]) delete d.signed[k].img; });
  const signedCount = ['p1', 'p2'].filter(k => d.signed && d.signed[k]).length;
  const status = signedCount === 2 ? 'signed' : signedCount === 1 ? 'partially-signed' : 'draft';
  const now = new Date().toISOString();
  const existing = d.docId ? rows_(SH.DOCS).find(x => x.docId === d.docId) : null;
  d.docId = existing ? existing.docId : Utilities.getUuid();
  const json = JSON.stringify(d);
  if (json.length > 48000) throw new Error('نص المذكرة طويل جداً للحفظ (الحد ~48 ألف حرف)');

  const rec = { docId: d.docId, refNo: clean_(d.refNo, 60), title: clean_(d.title, 200), status, json, updatedBy: email, updatedAt: now };
  if (existing) update_(SH.DOCS, existing._row, rec);
  else append_(SH.DOCS, Object.assign(rec, { createdBy: email, createdAt: now }));
  return { ok: true, docId: d.docId, status };
}

/* ======================= Helpers ======================= */
function email_() {
  const e = norm_(Session.getActiveUser().getEmail());
  if (!e || !e.endsWith('@' + ALLOWED_DOMAIN)) throw new Error('غير مصرّح — يجب الدخول بحساب الجامعة @' + ALLOWED_DOMAIN);
  return e;
}
function isAdmin_(e) { return rows_(SH.ADMINS).some(r => norm_(r.email) === e); }
function findSigner_(id) {
  const r = rows_(SH.SIGNERS).find(x => x.id === id && x.status === 'active');
  if (!r) throw new Error('التوقيع غير موجود');
  return r;
}
function hashTag_(tag) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, tag + '|' + prop_('PEPPER'), Utilities.Charset.UTF_8);
  return bytes.map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
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
function norm_(s) { return String(s || '').trim().toLowerCase(); }
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function sheet_(n) { return SpreadsheetApp.getActive().getSheetByName(n); }
function rows_(n) {
  const v = sheet_(n).getDataRange().getValues();
  const h = v.shift() || [];
  return v.map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = r[j]); return o; });
}
function append_(n, o) { sheet_(n).appendRow(HEAD[n].map(k => o[k] === undefined ? '' : o[k])); }
function update_(n, row, o) {
  const h = HEAD[n], rg = sheet_(n).getRange(row, 1, 1, h.length), cur = rg.getValues()[0];
  h.forEach((k, j) => { if (k in o) cur[j] = o[k]; });
  rg.setValues([cur]);
}
function stamp_() { const d = new Date(); return { ts: d, date: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'), time: Utilities.formatDate(d, TZ, 'HH:mm:ss') }; }
function logAccess_(email, action) { const s = stamp_(); append_(SH.ACCESS, { email, date: s.date, time: s.time, action }); }
function logSign_(email, o) { const s = stamp_(); append_(SH.SIGNINGS, Object.assign({ timestamp: s.ts, date: s.date, time: s.time, email }, o)); }

/**
 * منصة توقيع مذكرات التفاهم — SRU MoU Signing
 * Google Apps Script backend · v2.5
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
 *
 * Live collaboration (v1.8)
 *  - Each open page polls `sync` every few seconds: revision number, who is online, who is editing.
 *  - One editor at a time (soft edit lock); everyone else sees changes after each autosave.
 *  - A freshly inserted signature is shown to the other viewers through an END-TO-END encrypted
 *    relay (ECDH P-256 + AES-GCM between browsers). The server only forwards ciphertext it cannot
 *    read; relays live in CacheService (max 6 h) and are deleted once delivered.
 *
 * Access control (v2.0) — per MoU
 *  - admin  : Admins sheet → everything
 *  - editor : creator or listed in `editors` → view, edit, sign, finalize
 *  - signer : party representative (p1email/p2email) → view + add/remove own signature only
 *  - viewer : listed in `viewers` or in the GlobalViewers sheet (all MoUs) → read-only
 *  - anyone else → the MoU is invisible (not listed, cannot be opened or synced)
 *
 * Final copy (v1.9)
 *  - When the second party signs, that browser (which holds the decrypted signatures) renders the final
 *    A4 PDF with a signature-certificate page and calls `finalizeDoc`.
 *  - The server archives the PDF in a private Drive folder, records its SHA-256, e-mails it to both
 *    parties' representatives and the signature owners, and locks the MoU against unsigning.
 */

const APP_VERSION    = '2.5';
const ALLOWED_DOMAIN = 'sr.edu.sa';
const MAX_FAILS      = 5;
const LOCK_MINUTES   = 15;
const TZ             = 'Asia/Riyadh';
const SESSION_HOURS  = 6;     // external session length (CacheService max = 6h)
const OTP_MINUTES    = 10;
const MAX_LOGO_CHARS = 45000; // a logo must fit in one Sheet cell
const LIVE_TTL_MS    = 15000; // a viewer disappears from "online now" after 15 s without a heartbeat
const EDIT_TTL_MS    = 20000; // edit lock expires after 20 s without a heartbeat
const RELAY_CHUNK    = 90000; // CacheService value limit is 100 KB
const MAX_PDF_B64    = 24000000; // ~18 MB PDF (Gmail attachment limit is 25 MB)

const SH = { SIGNERS:'Signers', DOCS:'Documents', SIGNINGS:'SigningLog', ACCESS:'AccessLog', ADMINS:'Admins', EXT:'Externals', GV:'GlobalViewers' };
const HEAD = {
  Signers:    ['id','name','position','org','fileId','encSalt','encIv','iter','authSalt','authHash','fails','lockUntil','status','createdBy','createdAt','updatedAt','ownerEmail'],
  Documents:  ['docId','refNo','title','status','json','createdBy','createdAt','updatedBy','updatedAt','logoRight','logoLeft','externalEmails','rev','finalPdfId','finalizedAt','finalizedBy','pdfHash','sentTo','editors','viewers','partyEmails'],
  SigningLog: ['timestamp','date','time','email','action','signerId','signerName','docRef','party','docHash','result'],
  AccessLog:  ['email','date','time','action'],
  Admins:     ['email','name','addedAt'],
  Externals:  ['email','name','org','status','addedBy','addedAt','lastLogin'],
  GlobalViewers: ['email','name','addedBy','addedAt']
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
  if (!props.getProperty('FINAL_FOLDER_ID')) props.setProperty('FINAL_FOLDER_ID', DriveApp.createFolder('SRU-MoU-Final-PDFs').getId());

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
  const writes = ['setPdfFont','setPdfFontChoice','setDocAccess','addGlobalViewer','removeGlobalViewer','finalizeDoc','sync','relay','addSigner','unlock','changePin','deleteSigner','saveDoc','addAdmin','removeAdmin','addExternal','setExternalStatus','removeExternal','setGuestUrl','inviteExternal'];
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
      case 'sync':              return sync_(me, p);
      case 'relay':             return relay_(me, p);
      case 'finalizeDoc':       return finalizeDoc_(me, p);
      case 'getFinalPdf':       return getFinalPdf_(me, p);
      // ---- admin ----
      case 'listUsers':         return admin_(me) && listUsers_();
      case 'addAdmin':          return admin_(me) && addAdmin_(me, p);
      case 'removeAdmin':       return admin_(me) && removeAdmin_(me, p);
      case 'addExternal':       return admin_(me) && addExternal_(me, p);
      case 'setExternalStatus': return admin_(me) && setExternalStatus_(me, p);
      case 'removeExternal':    return admin_(me) && removeExternal_(me, p);
      case 'setGuestUrl':       return admin_(me) && setGuestUrl_(p);
      case 'inviteExternal':    return admin_(me) && inviteExternal_(me, p);
      case 'getPdfFontInfo':    return { choice: prop_('PDF_FONT_CHOICE') || 'cairo', name: prop_('PDF_FONT_NAME') || '' };
      case 'getPdfFont':        return getPdfFont_();
      case 'setPdfFont':        return admin_(me) && setPdfFont_(me, p);
      case 'setPdfFontChoice':  return admin_(me) && setPdfFontChoice_(me, p);
      case 'listDocAccess':     return admin_(me) && listDocAccess_();
      case 'setDocAccess':      return admin_(me) && setDocAccess_(me, p);
      case 'addGlobalViewer':   return admin_(me) && addGlobalViewer_(me, p);
      case 'removeGlobalViewer':return admin_(me) && removeGlobalViewer_(me, p);
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
function emails_(v) { return String(v || '').split(',').map(norm_).filter(Boolean); }
function isGlobalViewer_(e) { return rows_(SH.GV).some(r => norm_(r.email) === e); }
function partyEmailsOf_(r) {
  const own = emails_(r.partyEmails);
  return own.length ? own : emails_(r.externalEmails);   // documents saved before v2.0
}
/** Role of the caller on one MoU: 'admin' | 'editor' | 'signer' | 'viewer' | null */
function docRole_(me, r) {
  if (me.isAdmin) return 'admin';
  const e = me.email;
  if (!me.external && (norm_(r.createdBy) === e || emails_(r.editors).indexOf(e) > -1)) return 'editor';
  if (partyEmailsOf_(r).indexOf(e) > -1) return 'signer';
  if (emails_(r.viewers).indexOf(e) > -1) return 'viewer';
  if (!me.external && (me.gv !== undefined ? me.gv : isGlobalViewer_(e))) return 'viewer';
  if (me.external && emails_(r.editors).indexOf(e) > -1) return 'signer';   // guests never edit text
  return null;
}
function docVisible_(me, r) { return !!docRole_(me, r); }

function listDocs_(me) {
  const gv = !me.external && !me.isAdmin && isGlobalViewer_(me.email);
  return rows_(SH.DOCS).map(r => ({ r, role: docRole_(Object.assign({}, me, { gv }), r) })).filter(x => x.role).map(({ r, role }) => ({
    docId: r.docId, refNo: r.refNo, title: r.title, status: r.status, role,
    updatedBy: r.updatedBy, updatedAt: String(r.updatedAt)
  })).reverse();
}

function loadDoc_(me, docId) {
  const r = rows_(SH.DOCS).find(x => x.docId === docId);
  if (!r || !docVisible_(me, r)) throw new Error('المذكرة غير موجودة أو غير مصرّح بها');
  const d = JSON.parse(r.json);
  d.logoRight = String(r.logoRight || '');
  d.logoLeft  = String(r.logoLeft || '');
  d.rev = Number(r.rev) || 0;
  d.role = docRole_(me, r);
  d.final = r.finalPdfId ? { at: String(r.finalizedAt), by: String(r.finalizedBy), pdfHash: String(r.pdfHash), sentTo: String(r.sentTo) } : null;
  return d;
}

function saveDoc_(me, p) {
  let d = p.doc || {};
  const existing = d.docId ? rows_(SH.DOCS).find(x => x.docId === d.docId) : null;
  if (existing && existing.finalPdfId) throw new Error('المذكرة نهائية — صدرت نسختها الموقّعة ولا يمكن تعديلها أو إزالة التوقيعات.');
  const role = existing ? docRole_(me, existing) : (me.external ? null : 'editor');
  if (!role) throw new Error('غير مصرّح بهذه المذكرة');
  if (role === 'viewer') throw new Error('لديك صلاحية المشاهدة فقط لهذه المذكرة');
  const limited = me.external || role === 'signer';   // may only add/remove their own signature
  if (existing && !limited) {
    // Optimistic concurrency: refuse to overwrite a newer version saved by someone else
    if (p.baseRev !== undefined && p.baseRev !== null && Number(existing.rev || 0) !== Number(p.baseRev))
      throw new Error('CONFLICT|تم تعديل المذكرة من مستخدم آخر — تم تحميل أحدث نسخة، راجعها ثم أعد المحاولة.');
    const ed = JSON.parse(cache_().get('edit:' + existing.docId) || 'null');
    if (ed && ed.email !== me.email && Date.now() - ed.ts < EDIT_TTL_MS && !p.signOnly)
      throw new Error('CONFLICT|يقوم ' + ed.email + ' بتحرير المذكرة الآن — انتظر حتى ينتهي.');
  }

  if (limited) {
    // Guests and party representatives may only add/remove THEIR OWN signature.
    if (!existing) throw new Error('غير مصرّح');
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
  delete d.logoRight; delete d.logoLeft; delete d.rev; delete d.final;
  const parties = [d.p1email, d.p2email].map(norm_).filter(e => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e));
  const emails = parties.filter(e => !e.endsWith('@' + ALLOWED_DOMAIN)).join(',');
  const json = JSON.stringify(d);
  if (json.length > 48000) throw new Error('نص المذكرة طويل جداً للحفظ (الحد ~48 ألف حرف)');

  const rev = (existing ? Number(existing.rev || 0) : 0) + 1;
  delete d.rev;
  const rec = { docId: d.docId, refNo: clean_(d.refNo, 60), title: clean_(d.title, 200), status, json,
                updatedBy: me.email, updatedAt: now, logoRight, logoLeft, externalEmails: emails, partyEmails: parties.join(','), rev: String(rev) };
  if (existing) update_(SH.DOCS, existing._row, rec);
  else append_(SH.DOCS, Object.assign(rec, { createdBy: me.email, createdAt: now }));
  cache_().put('rev:' + d.docId, String(rev), 21600);
  return { ok: true, docId: d.docId, status, rev };
}

function cleanLogo_(v) {
  v = String(v || '');
  if (['', 'none', 'sru', 'ictd', 'itoc'].indexOf(v) > -1) return v;
  if (/^data:image\/(png|jpeg|webp);base64,[A-Za-z0-9+/=]+$/.test(v) && v.length <= MAX_LOGO_CHARS) return v;
  throw new Error('صورة الشعار غير صالحة أو كبيرة جداً');
}

/* ======================= Final signed copy ======================= */
function finalizeDoc_(me, p) {
  const r = rows_(SH.DOCS).find(x => x.docId === p.docId);
  if (!r || !docVisible_(me, r)) throw new Error('المذكرة غير موجودة أو غير مصرّح بها');
  if (docRole_(me, r) === 'viewer') throw new Error('لديك صلاحية المشاهدة فقط لهذه المذكرة');
  if (r.status !== 'signed') throw new Error('لا يمكن إصدار النسخة النهائية قبل توقيع الطرفين');
  if (r.finalPdfId && !(p.resend && !me.external && me.isAdmin)) return { ok: true, already: true, at: String(r.finalizedAt) };
  if (!isB64_(p.pdf, MAX_PDF_B64)) throw new Error('ملف PDF غير صالح أو كبير جداً');
  const bytes = Utilities.base64Decode(p.pdf);
  if (String.fromCharCode.apply(null, bytes.slice(0, 5).map(b => (b + 256) % 256)) !== '%PDF-') throw new Error('الملف ليس PDF');

  const d = JSON.parse(r.json);
  const name = (clean_(r.refNo, 60) || 'MoU').replace(/[^\w\u0600-\u06FF.-]+/g, '-') + '-signed.pdf';
  const pdfHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
  const blob = Utilities.newBlob(bytes, 'application/pdf', name);
  let file;
  if (r.finalPdfId) { try { DriveApp.getFileById(r.finalPdfId).setTrashed(true); } catch (e) {} }
  file = DriveApp.getFolderById(prop_('FINAL_FOLDER_ID')).createFile(blob);

  // Recipients: party representatives + signature owners + document creator
  const signers = rows_(SH.SIGNERS);
  const owner = id => { const s = signers.find(x => x.id === id); return s ? norm_(s.ownerEmail) : ''; };
  const valid = e => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e);
  const to = [norm_(d.p1email), norm_(d.p2email), owner(d.signed && d.signed.p1 && d.signed.p1.signerId),
              owner(d.signed && d.signed.p2 && d.signed.p2.signerId), norm_(r.createdBy)]
    .filter((e, i, a) => e && valid(e) && a.indexOf(e) === i);

  const audit = rows_(SH.SIGNINGS).filter(x => x.docRef === r.refNo && x.action === 'sign' && x.result === 'success');
  const now = new Date().toISOString(), s = stamp_();
  const partyRow = k => {
    const sg = d.signed && d.signed[k]; if (!sg) return '';
    const a = audit.filter(x => x.party === k && x.signerId === sg.signerId).pop();
    return '<tr><td style="padding:6px;border:1px solid #ddd">' + (k === 'p1' ? 'الطرف الأول' : 'الطرف الثاني') + '<br><small>' + esc_(d[k + 'org']) + '</small></td>' +
      '<td style="padding:6px;border:1px solid #ddd"><b>' + esc_(sg.signerName) + '</b><br><small>' + esc_(sg.position || '') + '</small></td>' +
      '<td style="padding:6px;border:1px solid #ddd">' + (a ? esc_(a.date + ' ' + a.time) : esc_(sg.at)) + '<br><small dir="ltr">' + esc_(a ? a.email : '') + '</small></td></tr>';
  };
  const html =
    '<div dir="rtl" style="font-family:Tahoma,Arial;color:#241436;max-width:640px">' +
    '<h3 style="color:#3a1464;margin-bottom:4px">تم توقيع مذكرة التفاهم من الطرفين</h3>' +
    '<p style="margin-top:0;color:#6b5a80">' + esc_(r.title) + ' — رقم <b dir="ltr">' + esc_(r.refNo) + '</b></p>' +
    '<table style="border-collapse:collapse;font-size:13px;width:100%"><tr style="background:#501e8c;color:#fff">' +
    '<th style="padding:6px">الطرف</th><th style="padding:6px">الموقّع</th><th style="padding:6px">وقت التوقيع (توقيت الرياض)</th></tr>' +
    partyRow('p1') + partyRow('p2') + '</table>' +
    '<p style="font-size:12px;color:#6b5a80">بصمة الملف المرفق (SHA-256):<br><span dir="ltr" style="font-family:monospace">' + pdfHash + '</span></p>' +
    '<p>مرفق النسخة النهائية الموقّعة بصيغة PDF، وتُحفظ نسخة أرشيفية لدى الجامعة.</p>' +
    '<hr><p style="font-size:12px;color:#6b5a80" dir="ltr">The Memorandum of Understanding <b>' + esc_(r.refNo) + '</b> has been signed by both parties. ' +
    'The final signed PDF is attached (SHA-256 above).</p>' +
    '<p style="font-size:12px;color:#6b5a80">جامعة سليمان الراجحي — الإدارة التنفيذية للاتصالات وتقنية المعلومات<br>Sulaiman Al Rajhi University — Executive Directorate of Communications &amp; IT</p></div>';

  if (to.length) MailApp.sendEmail({
    to: to.join(','),
    subject: 'تم توقيع مذكرة التفاهم ' + r.refNo + ' — النسخة النهائية | MoU fully signed',
    htmlBody: html, attachments: [blob], name: 'منصة توقيع مذكرات التفاهم — SRU'
  });

  const rev = Number(r.rev || 0) + 1;   // bump revision so every open page refreshes and shows the final banner
  update_(SH.DOCS, r._row, { finalPdfId: file.getId(), finalizedAt: now, finalizedBy: me.email, pdfHash, sentTo: to.join(','), rev: String(rev) });
  cache_().put('rev:' + r.docId, String(rev), 21600);
  logSign_(me.email, { action: p.resend ? 'final-resent' : 'finalized', docRef: r.refNo, docHash: pdfHash.slice(0, 64), result: 'sent:' + to.length });
  return { ok: true, at: now, sentTo: to, pdfHash, date: s.date };
}

function getFinalPdf_(me, p) {
  const r = rows_(SH.DOCS).find(x => x.docId === p.docId);
  if (!r || !docVisible_(me, r) || !r.finalPdfId) throw new Error('لا توجد نسخة نهائية لهذه المذكرة');
  const f = DriveApp.getFileById(r.finalPdfId);
  logSign_(me.email, { action: 'final-downloaded', docRef: r.refNo, result: 'success' });
  return { name: f.getName(), b64: Utilities.base64Encode(f.getBlob().getBytes()) };
}

/* ======================= Live collaboration ======================= */
function cache_() { return CacheService.getScriptCache(); }

/** Ensures the caller may see this document; returns { rev, role }. Role is cached 60 s per user. */
function liveDoc_(me, docId) {
  const c = cache_();
  const aclKey = 'acl:' + docId + ':' + me.email;
  let rev = c.get('rev:' + docId), role = c.get(aclKey);
  if (rev === null || !role) {
    const r = rows_(SH.DOCS).find(x => x.docId === docId);
    role = r ? docRole_(me, r) : null;
    if (!role) throw new Error('المذكرة غير موجودة أو غير مصرّح بها');
    rev = String(Number(r.rev) || 0);
    c.put('rev:' + docId, rev, 21600);
    c.put(aclKey, role, 60);
  }
  return { rev: Number(rev), role };
}

function sync_(me, p) {
  const docId = clean_(p.docId, 64);
  if (!docId) return { rev: 0, editor: null, presence: [], relays: [] };
  const { rev, role: myRole } = liveDoc_(me, docId);
  const c = cache_(), now = Date.now(), pk = 'pres:' + docId, ek = 'edit:' + docId;

  const pres = JSON.parse(c.get(pk) || '{}');
  Object.keys(pres).forEach(e => { if (now - pres[e].ts > LIVE_TTL_MS) delete pres[e]; });
  if (p.leave) delete pres[me.email];
  else {
    const mine = pres[me.email] || { name: displayName_(me), external: me.external };
    mine.ts = now; mine.role = myRole;
    if (isJwk_(p.pub)) mine.pub = { kty: 'EC', crv: 'P-256', x: p.pub.x, y: p.pub.y };
    pres[me.email] = mine;
  }

  let ed = JSON.parse(c.get(ek) || 'null');
  if (ed && (now - ed.ts > EDIT_TTL_MS || !pres[ed.email])) ed = null;
  if ((myRole === 'admin' || myRole === 'editor') && p.wantEdit && !p.leave && (!ed || ed.email === me.email)) ed = { email: me.email, ts: now };
  if (ed && ed.email === me.email && (p.releaseEdit || p.leave)) ed = null;
  if (ed) c.put(ek, JSON.stringify(ed), 60); else c.remove(ek);
  c.put(pk, JSON.stringify(pres), 120);

  return {
    rev, editor: ed ? ed.email : null,
    presence: Object.keys(pres).map(e => ({ email: e, name: pres[e].name, role: pres[e].role || '', external: !!pres[e].external, pub: pres[e].pub || null, self: e === me.email })),
    relays: p.leave ? [] : takeRelays_(docId, me.email)
  };
}

/** Forward an end-to-end encrypted signature image to another online viewer. */
function relay_(me, p) {
  const docId = clean_(p.docId, 64), to = norm_(p.to), party = p.party === 'p2' ? 'p2' : 'p1';
  liveDoc_(me, docId);
  const pres = JSON.parse(cache_().get('pres:' + docId) || '{}');
  if (!pres[me.email] || !pres[to] || to === me.email) throw new Error('المستلم غير متصل بالمذكرة');
  if (!isB64_(p.iv, 64) || !isB64_(p.ct, 1200000) || !isJwk_(p.fromPub)) throw new Error('بيانات غير صالحة');
  const c = cache_(), base = 'rly:' + docId + ':' + to + ':' + party, put = {};
  const n = Math.ceil(p.ct.length / RELAY_CHUNK);
  for (let i = 0; i < n; i++) put[base + ':' + i] = p.ct.slice(i * RELAY_CHUNK, (i + 1) * RELAY_CHUNK);
  put[base + ':m'] = JSON.stringify({ from: me.email, fromName: pres[me.email].name, party, iv: p.iv, n,
    fromPub: { kty: 'EC', crv: 'P-256', x: p.fromPub.x, y: p.fromPub.y },
    signerId: clean_(p.signerId, 64), hash: clean_(p.hash, 64) });
  c.putAll(put, 21600);
  const ik = 'rlyidx:' + docId + ':' + to, idx = JSON.parse(c.get(ik) || '[]');
  if (idx.indexOf(party) === -1) idx.push(party);
  c.put(ik, JSON.stringify(idx), 21600);
  return { ok: true };
}

function takeRelays_(docId, to) {
  const c = cache_(), ik = 'rlyidx:' + docId + ':' + to, idx = JSON.parse(c.get(ik) || '[]');
  if (!idx.length) return [];
  const out = [];
  idx.forEach(party => {
    const base = 'rly:' + docId + ':' + to + ':' + party, metaRaw = c.get(base + ':m');
    if (!metaRaw) return;
    const m = JSON.parse(metaRaw), keys = [];
    for (let i = 0; i < m.n; i++) keys.push(base + ':' + i);
    const parts = c.getAll(keys);
    if (keys.some(k => parts[k] == null)) return;
    m.ct = keys.map(k => parts[k]).join('');
    out.push(m);
    c.removeAll(keys.concat([base + ':m']));
  });
  c.remove(ik);
  return out;
}

function displayName_(me) {
  if (me.external) { const x = rows_(SH.EXT).find(r => norm_(r.email) === me.email); if (x && x.name) return String(x.name); }
  const a = rows_(SH.ADMINS).find(r => norm_(r.email) === me.email);
  if (a && a.name && a.name !== 'Owner') return String(a.name);
  return me.email.split('@')[0];
}
function isJwk_(k) {
  return !!k && typeof k === 'object' && typeof k.x === 'string' && typeof k.y === 'string' &&
    /^[A-Za-z0-9_-]{30,60}$/.test(k.x) && /^[A-Za-z0-9_-]{30,60}$/.test(k.y);
}

/* ======================= Users (admin) ======================= */
function listUsers_() {
  return {
    admins: rows_(SH.ADMINS).map(r => ({ email: String(r.email), name: String(r.name), addedAt: String(r.addedAt) })),
    externals: rows_(SH.EXT).filter(r => r.status !== 'removed').map(r => ({
      email: String(r.email), name: String(r.name), org: String(r.org), status: String(r.status),
      addedBy: String(r.addedBy), addedAt: String(r.addedAt), lastLogin: String(r.lastLogin || '')
    })),
    guestUrl: prop_('GUEST_URL') || '',
    globalViewers: rows_(SH.GV).map(r => ({ email: String(r.email), name: String(r.name), addedBy: String(r.addedBy), addedAt: String(r.addedAt) }))
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

/* ======================= PDF font (v2.5) ======================= */
/** Cairo is embedded in the page; an admin may upload a licensed font (e.g. Khaled Bold) kept privately in Drive. */
function getPdfFont_() {
  const choice = prop_('PDF_FONT_CHOICE') || 'cairo', id = prop_('PDF_FONT_ID');
  if (choice !== 'custom' || !id) return { choice: 'cairo' };
  const c = cache_(), key = 'pdffont:' + id;
  let b64 = null;
  const meta = c.get(key + ':n');
  if (meta) { const keys = []; for (let i = 0; i < Number(meta); i++) keys.push(key + ':' + i); const parts = c.getAll(keys); if (keys.every(k => parts[k] != null)) b64 = keys.map(k => parts[k]).join(''); }
  if (!b64) {
    b64 = Utilities.base64Encode(DriveApp.getFileById(id).getBlob().getBytes());
    const put = {}, n = Math.ceil(b64.length / RELAY_CHUNK);
    for (let i = 0; i < n; i++) put[key + ':' + i] = b64.slice(i * RELAY_CHUNK, (i + 1) * RELAY_CHUNK);
    put[key + ':n'] = String(n);
    try { c.putAll(put, 21600); } catch (e) {}
  }
  return { choice: 'custom', name: prop_('PDF_FONT_NAME') || '', ext: prop_('PDF_FONT_EXT') || 'ttf', b64 };
}
function setPdfFont_(me, p) {
  const ext = String(p.ext || '').toLowerCase();
  if (['ttf', 'otf', 'woff', 'woff2'].indexOf(ext) === -1) throw new Error('صيغة الخط غير مدعومة (TTF أو OTF أو WOFF أو WOFF2)');
  if (!isB64_(p.b64, 4200000)) throw new Error('ملف الخط غير صالح أو كبير جداً (الحد 3MB)');
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('FONT_FOLDER_ID')) props.setProperty('FONT_FOLDER_ID', DriveApp.createFolder('SRU-MoU-Fonts').getId());
  const old = props.getProperty('PDF_FONT_ID');
  if (old) { try { DriveApp.getFileById(old).setTrashed(true); } catch (e) {} }
  const name = clean_(p.name, 120) || ('font.' + ext);
  const file = DriveApp.getFolderById(props.getProperty('FONT_FOLDER_ID')).createFile(Utilities.newBlob(Utilities.base64Decode(p.b64), 'application/octet-stream', name));
  props.setProperties({ PDF_FONT_ID: file.getId(), PDF_FONT_NAME: name, PDF_FONT_EXT: ext });
  logAccess_(me.email, 'pdf-font-uploaded:' + name);
  return { ok: true };
}
function setPdfFontChoice_(me, p) {
  const choice = p.choice === 'custom' ? 'custom' : 'cairo';
  if (choice === 'custom' && !prop_('PDF_FONT_ID')) throw new Error('ارفع ملف الخط المخصص أولاً');
  PropertiesService.getScriptProperties().setProperty('PDF_FONT_CHOICE', choice);
  logAccess_(me.email, 'pdf-font-choice:' + choice);
  return { ok: true };
}

/* ======================= Per-MoU access (admin) ======================= */
function listDocAccess_() {
  return rows_(SH.DOCS).map(r => ({
    docId: r.docId, refNo: String(r.refNo), title: String(r.title), status: String(r.status),
    createdBy: String(r.createdBy), editors: emails_(r.editors), viewers: emails_(r.viewers), parties: partyEmailsOf_(r)
  })).reverse();
}
function checkAccessEmail_(e) {
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) throw new Error('بريد غير صالح: ' + e);
  if (!e.endsWith('@' + ALLOWED_DOMAIN) && !isExternalActive_(e))
    throw new Error(e + ' ليس من منسوبي الجامعة — أضفه أولاً في "المستخدمون الخارجيون"');
}
function setDocAccess_(me, p) {
  const r = rows_(SH.DOCS).find(x => x.docId === p.docId);
  if (!r) throw new Error('المذكرة غير موجودة');
  const uniq = a => (a || []).map(norm_).filter((e, i, x) => e && x.indexOf(e) === i);
  const editors = uniq(p.editors), viewers = uniq(p.viewers).filter(e => editors.indexOf(e) === -1);
  editors.concat(viewers).forEach(checkAccessEmail_);
  const before = emails_(r.editors).concat(emails_(r.viewers));
  update_(SH.DOCS, r._row, { editors: editors.join(','), viewers: viewers.join(',') });
  before.forEach(e => cache_().remove('acl:' + r.docId + ':' + e));      // revoke live access immediately
  logAccess_(me.email, 'doc-access:' + r.refNo + ' editors=' + editors.length + ' viewers=' + viewers.length);
  return { ok: true };
}
function addGlobalViewer_(me, p) {
  const email = norm_(p.email);
  if (!email.endsWith('@' + ALLOWED_DOMAIN)) throw new Error('قائمة المشاهدة العامة لمنسوبي الجامعة فقط (@' + ALLOWED_DOMAIN + ')');
  if (rows_(SH.GV).some(r => norm_(r.email) === email)) throw new Error('موجود مسبقاً');
  append_(SH.GV, { email, name: clean_(p.name, 120), addedBy: me.email, addedAt: new Date().toISOString() });
  logAccess_(me.email, 'global-viewer-added:' + email);
  return { ok: true };
}
function removeGlobalViewer_(me, p) {
  const r = rows_(SH.GV).find(x => norm_(x.email) === norm_(p.email));
  if (!r) throw new Error('غير موجود');
  sheet_(SH.GV).deleteRow(r._row);
  logAccess_(me.email, 'global-viewer-removed:' + r.email);
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

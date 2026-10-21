# CHANGELOG — منصة توقيع مذكرات التفاهم / SRU MoU Signing

## v2.5 — 2026-10-21
- **The PDF now really uses the chosen font.** Previously the capture engine could not load web fonts, so PDFs used a browser fallback font. Pages are now rendered by a built-in renderer (SVG page image with embedded fonts → canvas → jsPDF); html2canvas is no longer needed.
- **Cairo embedded** (SIL Open Font License, Arabic + Latin, Regular/Bold, ~78 KB) — works offline, no external font link.
- **Custom PDF font (Users tab → "خط ملف PDF"):** admins can upload a licensed TTF/OTF/WOFF/WOFF2 (e.g. Khaled Bold, ≤3 MB), stored privately in Drive folder `SRU-MoU-Fonts`, and switch between Cairo and the custom font.
- **"معاينة PDF" button** (toolbar and Users tab): downloads a watermarked preview of the current MoU without saving or sending anything; unsigned parties show "لم يُوقَّع بعد".

## v2.4 — 2026-10-21
- PDF header line fixed: reference number, Gregorian date and Hijri date are now separate, non-breaking items on one line; English values (reference, date, e-mail) are isolated left-to-right so they no longer split ("ICTD-MOU-" / "2026-03") or overlap the divider.
- Gregorian date shown as DD/MM/YYYY م in the PDF.
- Fixed duplicated Hijri suffix ("هـ هـ") in the app header and the PDF.
- Certificate page: reference number and e-mail rendered left-to-right.

## v2.3 — 2026-10-21
- **Root-cause fix** for `SyntaxError: Invalid or unexpected token` (line 457:41) on the live Apps Script deployment, present since v1.9: Apps Script mangled the PDF-library web addresses (`https://…`) written literally inside the script, leaving the string unterminated, so the whole page script stopped and the MoU appeared empty.
- The CDN address is now assembled at run time; no literal web address remains inside any script. Developer note: never write `https://` URLs as string literals inside `<script>` in Apps Script HTML files.

## v2.2 — 2026-10-21
- Fix: `SyntaxError: Invalid or unexpected token` on the live deployment that stopped the page before the MoU loaded.
- All scripts in `Index.html` are now **ASCII-only**: Arabic text and emoji inside JavaScript are written as `\uXXXX` escapes (displayed identically), so copy/paste or storage in the Apps Script editor can no longer corrupt them.
- End-of-file marker: if `Index.html` is pasted incompletely, the page says so explicitly.
- Error bar now shows line and column.

## v2.1 — 2026-10-21
- Fix: the MoU could appear empty (no reference number, date, parties, text or signature boxes) when start-up stopped early — most likely when `Index` and `Code.gs` were from different versions.
- Start-up is now step-by-step: a failing step no longer blanks the MoU.
- On-screen diagnostics: an orange bar warns when `Code.gs` and `Index` versions differ; a red bar shows any script error with its line number (reporter loads before the main script).
- Backward compatible with older servers: if the server doesn't send a role, the page falls back to the editor view (the server still enforces access).

## v2.0 — 2026-10-21
- **Per-MoU access control:** roles admin / editor / signer (party representative) / viewer, enforced on the server for listing, opening, saving, live sync, relay and finalization. Users without a role can't list, open or sync the MoU.
- **Users tab → "صلاحيات المذكرات":** pick a MoU, add editors or viewers by e-mail, remove with ×; shows automatic roles (creator, party representatives, admins, global list).
- **Users tab → "قائمة المشاهدة العامة":** new `GlobalViewers` sheet — members (SRU only) can view **all** MoUs read-only.
- Read-only mode in the MoU for viewers (banner, locked fields, no signing controls) and signer mode for party representatives (text locked, own signature only).
- Role shown in the "online now" bar and in the MoU picker (👁 viewer, ✍️ signer). Edit lock only for editors/admins (role computed on the server).
- Revoking access takes effect within seconds, even on open pages. Access changes are logged.
- **Behavior change:** staff no longer see every MoU by default — only their own or shared ones.
- Migration: new `Documents` columns `editors`, `viewers`, `partyEmails`.

## v1.9 — 2026-10-21
- **Final signed copy:** when the second party signs, the browser renders an A4 PDF (logos, MoU text, both signatures, document fingerprint) plus an **Electronic Signature Certificate** page, with page numbers.
- Server archives the PDF in the private Drive folder `SRU-MoU-Final-PDFs`, records SHA-256, finalizer and recipients (new `Documents` columns), and **e-mails it as an attachment** to both representatives, the signature owners and the creator (bilingual e-mail with signing audit table).
- All open devices refresh and show the "final copy issued" banner with a **Download PDF** button.
- After finalization the server rejects any edit or unsigning; remove-signature buttons are hidden.
- If a signature image isn't available on the finalizing device, the PDF uses an e-signature certificate block (user is asked first).
- PDF engine: manual A4 pagination (breaks between paragraphs, never mid-line) + browser-native text rendering for correct Arabic shaping.
- `setup` creates the final-PDF folder automatically.

## v1.8 — 2026-10-21
- **Live collaboration:** both parties can open the same MoU on different devices; pages sync every 4 s.
- **"Online now" bar:** names and e-mails of everyone viewing the MoU (guests marked), with the current editor highlighted.
- **Auto-refresh + autosave:** edits autosave 2.5 s after typing stops and appear on the other device automatically.
- **Single editor lock:** the first person to edit holds a soft lock; others become read-only with a banner; lock frees after inactivity or when the MoU is signed.
- **Signature appears instantly on the other device** via an end-to-end encrypted relay (ECDH P-256 + AES-GCM between browsers); the server forwards ciphertext only and keeps nothing on disk.
- **Conflict protection:** revision number per MoU (`rev` column); stale saves are rejected and the latest version is loaded.
- Visual cue (toast + flash) when the other party signs.

## v1.7 — 2026-10-21
- **Users tab (admins):** add/remove platform admins; add, disable, delete and e-mail-invite external users.
- **External (guest) access:** guests listed in the new `Externals` sheet sign in through a second "Anyone" deployment with a one-time 6-digit code e-mailed to them (10-minute expiry, 5 tries, rate-limited, 6-hour session).
- Guests see only MoUs where their e-mail is a party representative, cannot edit MoU text, and can add/remove only their own signature (enforced on the server).
- Any user can register **their own** signature (linked to their e-mail); admins can register for anyone and link an owner e-mail.
- **MoU logos:** right and left logo slots above the MoU (none / SRU / ICTD / ITOC / upload), saved per document.
- Official SRU, ICTD and ITOC logos embedded in the platform and added to the landing page `assets/`.
- Landing page: optional "Guest sign-in" button.
- Automatic schema migration of existing sheets (new columns + `Externals` sheet); writes are now by column name.
- `appsscript.json`: removed fixed web-app access so two deployments (University / Anyone) can coexist.

## v1.6 — 2026-10-21
- Fix: "TypeError: Cannot read properties of null (reading 'appendRow')" on first sign-in when `setup` had not been run.
- Backend is now self-healing: missing sheets, Drive folder or server secret are created automatically on first use (`ensure_`).
- Clear error message if the script is not bound to the Google Sheet.

## v1.5 — 2026-10-21
- Landing page: Web App URL placed directly on the sign-in button (`href`); removed the script check and the "WEB_APP_URL not set" alert, so the button works even with a cached or blocked script.
- Version label updated to v1.5.

## v1.4 — 2026-10-21
- Landing page: Web App URL configured (`WEB_APP_URL` → production `/exec` link); sign-in button now opens the platform.
- Landing page: stricter URL check and no-cache headers so GitHub Pages/browser caching doesn't serve an old copy.
- Version label updated to v1.4 across landing page and backend.

## v1.3 — 2026-10-21 (built on v1.1)
- Backend moved to Google Apps Script + Google Sheets; encrypted signature files kept in a private Drive folder.
- Sign-in limited to @sr.edu.sa accounts (Web App: "Anyone within SRU"); ITOC landing page on GitHub Pages.
- Every sign-in logged (email, date, time) in `AccessLog`; every sign/view/unsign/PIN change/failed attempt logged in `SigningLog`.
- Signatures always encrypted: the PIN never leaves the browser; the server stores only ciphertext + SHA-256(PIN-derived tag + pepper).
- Server-side PIN verification and lockout (5 attempts → 15 min) — cannot be bypassed from the browser.
- Stronger PIN policy: at least 8 characters with letters and digits.
- Admin roles (`Admins` sheet): only admins add/delete signers.
- MoUs saved to the Sheet (draft / partially signed / signed); reopening a signed MoU requires the PIN to display signatures.
- Removing a signature now requires the signer's PIN.
- Tamper warning when a document no longer matches its fingerprint at signing.
- Watermark now includes the viewer's e-mail.

## v1.2 — 2026-10-21 (deferred)
- Electron desktop wrapper with OS-level screen-capture exclusion. Postponed to a later phase.

## v1.1 — 2026-10-21
- Blurred signatures (press-and-hold to reveal), watermark, cover screen on focus loss, capture/save shortcuts intercepted, copy/right-click disabled.

## v1.0 — 2026-10-21
- Editable MoU, two-party signatures with PIN, document lock + SHA-256 fingerprint, local encrypted signature register.

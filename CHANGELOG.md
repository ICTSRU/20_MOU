# CHANGELOG — منصة توقيع مذكرات التفاهم / SRU MoU Signing

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

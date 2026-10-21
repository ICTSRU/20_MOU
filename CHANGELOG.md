# CHANGELOG — منصة توقيع مذكرات التفاهم / SRU MoU Signing

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

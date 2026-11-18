# QA Addendum — v15.6.1
- Source code and test-suite changes include pre-upgrade profile backup integrity/idempotency.
- Real user data is never written during build or QA; installer test uses a different QA-only AppId.
- Authenticode signature is not available on this Windows machine.

# QA Addendum 15.6.0
- Source self, database, export, period PDF, PDF design and Windows UI smoke tests were exercised.
- New synthetic regression covers employee search, read-only safe migration and continuous PDF.
- Windows installer and compiled EXE remain unsigned until trusted code signing is available.
- Android old 15.3 package ID differs; in-place APK upgrade is not possible from that version.
# QA Report — Daily Work Reports 15.5.0 Flexible Public Edition

## Public configuration
- No employee names, section labels, custom report titles, field labels, or dropdown values are provided at first launch.
- Windows first-run wizard test: PASS (user supplied title, category, role caption and first field; persisted configuration and successful reopening).
- Blank default configuration and normalization tests: PASS.
- Original private staff-name scan (24 names): 0 matches in published source.
- Sensitive data files removed from public source ZIP: no SQLite DB, CSV, PDF, generated logs, caches, private config or APK/EXE.
- Deliberately public contact: Abu Omar / protech.2025.2026@gmail.com / +966503276927.

## Windows 15.5.0
- Python source self-test: PASS.
- Source database smoke: PASS.
- Source export smoke: PASS.
- Source UI smoke: PASS.
- Structure unit tests: 8/8 PASS.
- PyInstaller one-file EXE build: PASS.
- Compiled EXE self-test: PASS, exit 0.
- Compiled EXE database smoke: PASS, exit 0.
- Compiled EXE export smoke: PASS, exit 0.
- Compiled EXE UI smoke: PASS, exit 0.
- First-run wizard is confirmed bundled in the compiled EXE archive.
- One non-fatal Tcl/Tk after-cleanup warning was observed during a source structure test; no failed assertions.

## Android 15.5.0
- Gradle offline assembleDebug: PASS.
- APK archive integrity: PASS.
- Developer email, phone and name: present in DEX.
- Old employee-name seed scan in DEX: zero known matches.
- Runtime emulator / physical device test: NOT EXECUTED (ADB reported no connected devices).
- Android APK is debug-signed; production signing and device acceptance tests remain necessary.

## Release security
- Public repository source has an allowlist of reviewed files.
- No real employee database or existing personalized company config copied.
- Code signing certificates are not supplied. Desktop EXE may trigger SmartScreen.
- Licensing update (2026-11-07): MIT adopted for original ProTech-owned code. This post-QA documentation change does not add any runtime test results.

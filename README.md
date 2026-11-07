# Daily Work Reports — Flexible Public Edition 15.5.0

**Developer:** Abu Omar
**Email:** protech.2025.2026@gmail.com
**Phone / WhatsApp:** +966503276927

An Arabic-first customizable daily work reporting application for Windows and Android.

## Clean first launch
The application ships with no real or preselected staff names, section names, custom
report titles, data-entry columns, or dropdown choices. Every user defines these
according to their own workflow.

Windows: a guided first-run form asks the user to create the report title, first
section, person/role caption, and first data field. More sections and fields can be
added using the structure manager.

Android: open an empty dashboard, create a first section, then add people and
data fields. Set the report title through Settings.

Generic navigation labels (Save, Settings, Add section, etc.) remain in place so
the application stays usable. Synthetic fixture text appears only in tests.

## Build on Windows
Requirements: Windows 10/11, Python 3.13, plus libraries listed in
windows/requirements_15.2.txt.

1. Open a PowerShell terminal inside windows.
2. Run: py -3.13 -m venv .venv
3. Run: .venv\Scripts\python.exe -m pip install -r requirements_15.2.txt
4. Run: .venv\Scripts\python.exe DailyWorkReports_15.2.py

Source QA: python scripts/verify_public.py

Data is stored separately in LOCALAPPDATA/DailyWorkReportsFlexible_v155.

## Build on Android
Requirements: Android Studio, Android SDK 36, JDK 17 and Gradle 8.11.1.
Open the android folder and run assembleDebug. For production, generate a
separately signed release with your own private key.
The application ID is com.protech.dailyreports.flexible, isolated from the older app.

## Distribution, privacy, and licensing
No private user data is included. Do not commit user DBs, exported reports,
private config files, backup archives or signing keys.
The developer email and phone above are published intentionally.
Read PUBLISH_GITHUB.md and QA_REPORT.md before releasing.
No open source license has been selected; see LICENSE_DECISION.md.

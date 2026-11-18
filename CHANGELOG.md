# Windows v15.7.1 — Startup Compatibility Hotfix

- Fix KeyError('technicians') when legacy stored templates are normalized with an empty public default category catalog.
- No deletion or overwrite of user profiles, staff, or reports.
- Tests cover the original historical template database and private startup using isolated data copies.

# Windows v15.7.0
- Staff KPI dashboard, date-aware report opening, and safe prevention of accidental switching with unsaved work.
- Public setup ships without data; a separate local-only private installer is built outside the repository.

# Release 15.6.1
- In-place updater with historical installer AppId and verified backups of older profiles.
- Does not uninstall or delete saved reports or configuration.

# Release 15.6.0
- Employee substring search + Arabic-friendly matching on Windows; searchable Android entry.
- Compact flowing PDF tables and continuous daily report export.
- Improved navigation and quick actions; Ctrl+F employee search.
- Safe opt-in migration from old Windows profiles without overwriting old or non-empty new data.
- New Windows Inno Setup installer recipe, build tests, and regression checks.
# Release 15.5.0
- No preconfigured employee list, section titles, report titles or data fields.
- Windows interactive first-run configuration.
- Android empty state and custom report title / role label editing.
- Public developer name Abu Omar, email and phone shown in the applications.
- Separate Windows user-data storage and Android package identifier.
- Private employee lists are not shipped.
- Clean GitHub source archive, QA checks and build instructions.

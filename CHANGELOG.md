## 2.0.4 — Advanced cleanup

- Advanced: Remove Backscan size (MB) and Backscan retry minutes controls. Code continues using current defaults.
- Advanced: Remove buttons “Force backscan (missing zones)” and “Replace CoV Faction from CSV)”.

## 2.0.3 — Version visible in tray

- Tray: Show app version in the tray tooltip when hovering the icon.
- Tray menu: Add a disabled header item with the current version at the top.

## 2.0.2 — Update check UX

- Tray: “Check for updates…” now shows dialogs for Up to Date, Update Available, and error/no‑updater cases.
- Keeps existing auto‑download and “Restart Now” prompt when the update is ready.

## 2.0.1 — Desktop shortcut icon fix

- Build: Generate `build/icon.ico` from `assets/tray-256.png` so the desktop shortcut matches the system tray icon.
- Packaging: Run the icon generator automatically before `electron-builder`.

## 1.8.0 — Zone tracking by source file

- Feature: Track and upsert zones per log file to avoid collisions when the same character name exists on multiple servers (e.g., Zoic on P1999Green vs Real-Test).
- Webhook + local CSV now emit per-file zone rows; Google Apps Script updated to key by “Source Log File”.
- Inventory rows now link to the most recent zone source for each character.

Notes: After deploying the new Apps Script, clear existing rows in the “Zone Tracker” tab once to remove entries keyed by Character. Subsequent upserts will repopulate using the new key.

## 1.8.1 — Backscan last zone on first sight

## 1.8.2 — Configurable backscan + local CSV convenience

## 1.8.3 — Force backscan tool

## 1.8.4 — Unlimited and optimized backscan

- Backscan can be unlimited: set Backscan size (MB) to 0 to scan the entire file.
- Optimized reverse scanning to find the last “You have entered …” by reading from end in chunks and scanning lines from the end, minimizing re-processing.

- Tray: Add “Force backscan (missing zones)” to immediately backscan all logs that currently lack a recorded zone and update CSV + webhook.
- Helpful for diagnosing missing rows like Nsac/Poqet without waiting for periodic retries.

- Settings: Add Backscan size (MB, 5–20) and Backscan retry minutes (0 disables) to control how much of each log is searched and how often to retry until a zone is found.
- Scanner: Uses configured backscan size; periodically retries backscan for files without a recorded zone.
- Tray: Add “Open local CSV folder” menu item to jump directly to the CSV output directory.
- You can already choose the CSV output directory in Settings → Local CSV output (now easier to test with a repo folder).

- Improvement: When a log file is seen for the first time, backscan up to 10MB from the end to find the most recent "You have entered …" line. This seeds Zone Tracker immediately even if the last zone message falls outside the initial 256KB tail.
- Adds a small log message: "Backscan seeded last zone" with character and zone for traceability.
## 1.9.0 — Smarter /con heuristics (invis + combat)

- Track self invisibility start/stop times and treat invis as active up to 20 minutes (configurable). If a /con returns Indifferent while invis is active, prefer previous stable standing instead of locking in Indifferent.
- Track recent attacks per mob (by name). If a /con occurs soon after attacking that mob (default 5 minutes), treat hostile results as combat-biased and prefer the previous stable standing.
- Keeps short look-behind line heuristics but prioritizes time-window rules for reliability.
- Notes now include previous standing context: when fallback is applied due to invis or combat, the note includes e.g. “(invis; prev=Kindly)” or “(combat; prev=Warmly)”.
## 1.9.1 — CoV list viewer + easy overrides

- Settings: Add CoV Mob List section to view the bundled default list and manage overrides.
- Users can add mobs or hide defaults with simple Add/Remove controls. Changes apply immediately and affect /con matching.
- Default list updated to include “a wyvern”.
## 1.9.2 — Settings cleanup + spell hit parsing

- Settings: Remove CoV Mob List section (now available under tray → CoV Mob List…).
- Combat parsing: Add spell hit detection (Your <spell> hits <mob>…, You blast/smite <mob>…, <mob> has taken N damage from your …) to mark recent combat per mob.
## 1.9.3 — Advanced page + classic icons

- Moved Force Backscan, Full Refresh to Sheet, and Backscan configuration to a new Advanced window (tray → Advanced…).
- Settings page simplified (advanced options removed).
- Added simple classic-style icons in assets: `assets/simple-eq-1999.svg`, `assets/simple-xp-orb.svg`, `assets/simple-xp-shield.svg`.

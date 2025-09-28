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

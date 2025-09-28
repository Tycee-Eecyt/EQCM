## 1.8.0 — Zone tracking by source file

- Feature: Track and upsert zones per log file to avoid collisions when the same character name exists on multiple servers (e.g., Zoic on P1999Green vs Real-Test).
- Webhook + local CSV now emit per-file zone rows; Google Apps Script updated to key by “Source Log File”.
- Inventory rows now link to the most recent zone source for each character.

Notes: After deploying the new Apps Script, clear existing rows in the “Zone Tracker” tab once to remove entries keyed by Character. Subsequent upserts will repopulate using the new key.


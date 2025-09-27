# EQ Character Manager (Author: Tyler A)

System tray app for EverQuest logs that:
- Tracks last zone per character
- Infers CoV faction standing from `/con` lines (with invis/combat fallback)
- Parses inventory files and writes CSVs
- Checks a **Raid Kit** across characters
- Sends upserts to Google Sheets and can **push a character's inventory** to a dedicated sheet tab (like the screenshot).

## Quick Start
```bash
npm install
npm run start
```

## Settings
- Pick your EverQuest `logs` folder and base folder (where `*-Inventory.txt` live).
- (Optional) Paste your Google Sheets spreadsheet URL (used when opening the sheet).
- (Optional) Deploy the Apps Script (see `docs/deploy-sheets.html`) and paste the `/exec` URL + secret.

## Tray
- **Push inventory to sheet → <Character>**: creates a new tab in your spreadsheet with a header row:
  `Inventory for | File | Created On | Modified On` and a table of items below (Location, Name, ID, Count, Slots).
- Start/pause scanning and choose scan interval.

## Local CSVs
Written to `~/.eq-character-manager/data/sheets` (or your custom path):
- `Zone Tracker.csv`
- `CoV Faction.csv`
- `Inventory Summary.csv` (includes Raid Kit columns)
- `Inventory Items - <Character>.csv`

## Google Sheets backend
Open `docs/deploy-sheets.html` in a browser and follow the steps. Copy+paste the `Code.gs` into your Apps Script project.

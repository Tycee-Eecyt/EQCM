## Deployment & Sync Pipeline Ideas

- Buffer all incoming EQCM webhook payloads into a database (e.g., MongoDB or Postgres). Suggested schema: users, characters, zones, factions, inventory, each with timestamps and a `needs_sync` flag.
- Maintain precomputed summaries per character (latest zone, faction standing, inventory) so scheduled sync jobs can grab fresh state quickly.
- Run a scheduler/job queue (BullMQ, Agenda, etc.) that every minute pulls a subset of characters needing an update (e.g., last sync ≥ 30 minutes) and posts their data to Google Sheets; stagger jobs to avoid API spikes.
- Sheets sync worker:
  - Fetch the latest rows from the DB.
  - Build the existing payload shapes (zones, factions, inventory, inventoryDetails).
  - Call the Google Sheets API (use the Express webhook endpoint) to write batches per tab.
  - Record `last_sheet_push` timestamps and clear `needs_sync`.
- Handle quota management via concurrency limits and exponential backoff on 429 responses; request quota increases from Google if needed.
- Keep credentials secure; if multi-tenant, store per-tenant sheet IDs/secrets.

## Branching / Release Strategy

- Keep `main` stable for production. Create and maintain a `beta` branch for ongoing integration (e.g., the Express backend).
- Workflow:
  - `git checkout -b beta` from `main`.
  - Merge feature branches into `beta`, push regularly.
  - Produce beta builds/releases from `beta` (tag them, mark as pre-release).
  - When ready, merge `beta` back into `main` (`git checkout main && git merge --ff-only beta`) and tag a stable release.
  - Handle hotfixes by branching from `main`, then merge fixes into both `main` and `beta`.

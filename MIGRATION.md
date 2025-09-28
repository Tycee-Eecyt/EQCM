EQ Character Manager v2 Migration Notes

- Why major: Inventory Summary schema can change at runtime; CSV `Raid Kit Summary.csv` renamed to `Raid Kit.csv`. Advanced webhook debounced by content digest.

- Action items
  - Redeploy Apps Script (`docs/Code.gs`) to your /exec deployment and update the app’s Apps Script URL if it changes.
  - Open the Raid Kit window and click Save & Close once to trigger a full Replace All.
  - If columns still look stale, use Advanced → Force Replace All.

- Notable changes
  - Raid Kit UI rebuilt (add/edit/remove, reset to defaults).
  - Wort/Cloak are fixed columns; hiding a fixed item removes its column from CSV and Sheet.
  - Custom Count extras write blank when zero (were 0).
  - Periodic sync posts only when CSV content changed.
  - Replace All is skipped if no data changed; Force Replace All bypasses skip.

- Backwards compatibility
  - If downstream tools expect a stable Inventory Summary schema, pin a fixed set of columns or consume the provided headers row dynamically.


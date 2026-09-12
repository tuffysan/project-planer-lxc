# Excel Round-trip

Project Planer remains the system of record while Excel can be used as an offline working surface.

## Safety model
1. Every exported editable record has a hidden `_ID`.
2. `_Hash` stores the record snapshot at export time.
3. On import, the Excel row is compared with both the export snapshot and the current database row.
4. If both sides changed, the item is flagged as a conflict.
5. Nothing is written until the user confirms the preview.
6. Missing Excel rows never trigger deletion.
7. A SQLite online backup is created before commit.
8. Import activity is written to the audit log.

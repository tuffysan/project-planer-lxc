# Project Planer v8.1.0 — Excel Round-trip Edition

## New
- Full-project Excel export directly from Project Workspace.
- Focused exports for Plan, Risk & Change, Resources and Finance.
- Round-trip metadata with hidden row IDs and snapshot hashes.
- Excel import with server-side preview before any database write.
- New rows in Excel create new entities.
- Existing rows update by stable internal ID.
- Conflict detection when both Project Planer and Excel changed after export.
- Conflict policy: keep app values or use Excel values.
- No automatic deletion when rows disappear from Excel.
- SQLite online backup before every committed import.
- Audit entry after export and import.
- Data validation dropdowns for common statuses, priorities, link types and yes/no values.
- Responsive import/preview UI integrated into Project Workspace.

## Workbook sheets
- Projektinformation
- Uppgifter
- Milstolpar
- Risker
- Ändringsärenden
- Resurser
- Kostnader
- Beroenden
- Beslut
- Möten
- Åtgärder
- Nyttor
- _Metadata (very hidden)

Milstolpar is an informational view derived from Uppgifter. Edit milestone records in the Uppgifter sheet.

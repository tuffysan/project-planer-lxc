# Project Planer v5.0.2 – Full application audit

The v5.0.1 codebase was statically audited route-by-route, template-by-template and against the SQLite schema.

## Runtime failures fixed

- Project Control queried `raid_items.kind`; the table uses `item_type`.
- Automation processing inserted notification fields `title` and `body`; the notification table uses `message` and `link`.
- Intelligence action creation inserted a non-existent `action_items.created_at` field.
- Automation project selection queried a non-existent `projects.archived` field; it now uses `archived_at` and `deleted_at`.
- Report Pack linked to the non-existent Flask endpoint `export_excel`; corrected to `export_project`.
- Project Workspace status report fields were aligned to `report_date` and `overall_rag`.

## Project Control

Project Control was rebuilt for clearer RAID, Change Request, decision and approval workflows. Forms now work with the actual database schema and the page is responsive.

## Mobile and tablet

Responsive hardening was added globally:
- two-column layouts collapse on tablet/mobile
- forms become one-column on phones
- touch-friendly buttons
- horizontal navigation/tabs where appropriate
- responsive tables become card rows on phones when marked responsive
- Gantt remains horizontally scrollable rather than being squeezed
- inputs/selects/textareas fit small screens
- panels and grids no longer force viewport overflow

## Release validation

`VERIFY-PYTHON.py` now checks:
- Python syntax
- Jinja template syntax
- template endpoint existence
- required `url_for()` route parameters
- SQLite schema creation
- `ensure_column()` migrations
- all static SQL statements against the generated schema

This prevents several classes of HTTP 500 regressions from being published unnoticed.

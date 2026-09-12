# Project Planer v5.0.4 – Visual UX Refinement

This release is a second Visual UX pass focused on correctness, visual hierarchy and mobile/tablet usability.

## Correctness fixes
- Main dashboard now uses `avg_progress` correctly instead of a non-existent `progress` field from the portfolio query.
- Main dashboard now receives an overdue task list instead of an integer, so the attention list and counters render correctly.
- Archived/deleted projects are excluded from dashboard portfolio metrics.

## Visual improvements
- RAG status is visible on the main dashboard, not only on a secondary page.
- Health score, trend, status reasons, progress and attention counters are shown together.
- Portfolio health donut and RAG distribution.
- Project status cards now explain *why* a project is green/orange/red.
- High-risk and overdue attention feed.
- Upcoming 21-day timeline.
- Portfolio Health search, filter and sort.
- Non-colour status labels for accessibility.
- Project cockpit quick navigation and status explanation.

## Mobile/tablet
- Mobile bottom navigation for Home, Portfolio, My Work and Search.
- Compact project cards and responsive filters.
- Touch-friendly controls.
- Horizontally scrollable project quick navigation.
- Improved small-screen visual hierarchy.

# Project Planer v8.0.1 — Complete UX Edition

## UX principles
1. Important work must be reachable in one or two interactions.
2. Global navigation is grouped by user intent, not by release/feature name.
3. Dashboards should lead to an action.
4. Advanced functionality is progressively disclosed.
5. Desktop, tablet and mobile have deliberate layouts.
6. Status is never communicated by color alone.
7. Forms use consistent controls, touch targets and submit feedback.
8. Empty states, focus states and reduced-motion preferences are supported.

## Changes
- Kept the cleaned v8 navigation from the v8.0.0 navigation fix.
- Rebuilt the v8 home page around attention and next actions rather than a feature catalogue.
- Added consistent responsive rules for cards, forms, tables, action rows and dashboards.
- Added 44px minimum touch targets.
- Added keyboard focus visibility and a skip-to-content link.
- Added reduced-motion support.
- Added safe-area handling for mobile bottom navigation.
- Added form submit feedback (`Sparar…`) to reduce double submission.
- Improved tablet breakpoints and mobile single-column behavior.
- Improved table overflow behavior on narrow screens.
- Added consistent section headers, action groups and help-text patterns.

## Quality gate
Static validation includes Python compilation, Jinja parsing, endpoint/url_for validation, shell syntax and ZIP CRC integrity.
Actual browser/runtime validation in the target Proxmox LXC remains the final gate.

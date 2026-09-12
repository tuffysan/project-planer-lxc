# Project Planer v14.0.0 — Ultimate Edition

v14.0.0 is a cumulative release built from v13.0.0.

## Product direction

The release deliberately reduces visible complexity while keeping advanced project-management capabilities available in context.

Primary navigation:
- Idag
- Mitt arbete
- Projekt
- Portfolio

Global functions such as search, create, reporting, resource planning, Azure DevOps and administration are accessed through contextual actions or the command palette.

## New in v14

- Role-aware `Idag` home page with attention-first prioritization.
- Unified project cockpit at `/projects/<id>/ultimate`.
- Deterministic project health based on overdue work, blocking work, high risks and budget variance.
- Global Ctrl/Cmd+K command palette.
- Simplified desktop and mobile navigation.
- Project view switcher for overview, list, board, Gantt, calendar, risks and collaboration.
- Contextual export menu for Excel Pro, Report Studio, PDF and PowerPoint.
- Attention queue with direct next actions.
- Team, Azure DevOps and recent collaboration context in the project cockpit.
- Responsive layout and reduced-motion support.
- Login now lands on the new user-centered home page.

## Product principles

1. Show what needs attention before showing everything.
2. Keep one mental model across desktop and mobile.
3. Put advanced tools in context instead of the primary navigation.
4. Prefer actions and decisions over passive dashboards.
5. Preserve specialist views for users who need them.
6. Make common work possible without knowing the product architecture.

## Validation

Static validation covers Python compilation, Jinja parsing, route references, shell syntax and ZIP integrity.
A real Proxmox runtime test is still required before calling the release production-verified.

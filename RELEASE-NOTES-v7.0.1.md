# Project Planer v7.0.1 — Full Quality Audit

Stability and quality release for Project Planer 7.

Highlights:
- fixed cross-project visibility in Task Experience Pro;
- enforced write permissions in Quick Add, Board, task quick updates and Planning Engine scheduling;
- aligned Project Wizard / Automation / Goals authorization with PM/admin responsibilities;
- moved late-roadmap tables into startup migrations;
- SQLite WAL + busy timeout hardening;
- archived/deleted project filtering;
- server-side intake validation;
- controlled 500 errors with request reference;
- `/health/live` plus stricter `/health/ready` smoke testing;
- stronger static verifier;
- additional desktop/tablet/mobile overflow and touch-target hardening.

See `AUDIT-v7.0.1.md` for the complete audit.

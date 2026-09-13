# Project Planer v16.1.1 — Unified UI Hotfix

## Fixar 500 på Start
v16.1.0:s nya `/start` använde två namn som inte finns i applikationen:
- `db_connect()`
- `require_project_access()`

Dessutom refererade SQL-frågan till `users.name`, medan Project Planer använder
`users.display_name`.

v16.1.1 använder nu de etablerade funktionerna:
- `db()`
- `current_user()`
- `visible_projects_for_user()`
- `project_or_404()`

Start-sidan använder samma accessmodell som resten av applikationen.

## En enda meny
Den extra `ux-v1600-nav` som lades ovanför den befintliga topbaren är borttagen.
Project Planers befintliga topbar är nu den enda huvudmenyn.

Första menyvalet heter **Start** och Project Planer-logotypen länkar också till `/start`.

## Projektöversikt
`/projects/<id>/overview` är också korrigerad:
- använder `project_or_404()`
- använder `db()`
- använder `users.display_name`

## Login
Efter normal inloggning landar användaren på den nya Start-sidan.

Ingen databasändring krävs.

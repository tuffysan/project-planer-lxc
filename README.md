# Project Planer LXC v2.1.0 – Multilingual Branding Edition

v2.1.0 bygger vidare på Complete Edition och lägger till språkval och en ny visuell identitet.

## Nytt i v2.1.0

- Ny Project Planer-appikon.
- Favicon och appikon i flera storlekar.
- Den nya grafiska Project Planer-identiteten inkluderas i appen.
- Språkväljare i toppmenyn.
- Språkvalet sparas i användarens session.
- Svenska.
- English.
- Deutsch.
- Norsk.
- Dansk.
- Suomi.
- Gemensamt översättningslager i backend för vidare lokalisering.
- Navigering, login, lösenordsbyte och centrala dashboardtexter använder språkstödet.
- `/about` visar den nya Project Planer-brandingen.
- Alla funktioner från v2.0.0 finns kvar.

## Språk

Välj språk direkt från menyn uppe till höger. Språkstödet är byggt så att fler språk enkelt kan läggas till i `LANGUAGES` och `TRANSLATIONS` i `app/app.py`.

## Branding

Assets ligger i:

```text
app/static/assets/project-planer-icon.png
app/static/assets/project-planer-icon-192.png
app/static/assets/project-planer-icon-128.png
app/static/assets/project-planer-icon-64.png
app/static/assets/project-planer-icon-32.png
app/static/assets/project-planer-branding.png
```

## Installation

```bash
VERSION=2.1.0 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

## Uppgradering

```bash
CTID=140 VERSION=2.1.0 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/update-lxc.sh)"
```

## Publicering

```powershell
.\PUBLISH.cmd
```

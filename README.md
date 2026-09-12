# Project Plan LXC

En enkel självhostad projektplans-app för Proxmox LXC.

## Funktioner

- Flera projekt
- Projektinformation: namn, kund, projektledare, beskrivning, start/slutdatum
- Aktiviteter med:
  - WBS
  - aktivitet
  - ansvarig
  - start- och slutdatum
  - status
  - prioritet
  - progress 0–100 %
  - milstolpe
  - beroenden
  - kommentar
- Dashboard med projektsammanställning
- Excel-export per projekt (`.xlsx`)
- Excel-export av alla projekt (`.xlsx`)
- SQLite-databas
- Systemd-tjänst
- Automatisk installation i Debian 12 LXC
- Automatisk backup av databasen före uppgradering

## Snabbinstallation i Proxmox

Kör på Proxmox-host:

```bash
VERSION=1.0.7 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

Installern hämtar applikationen från GitHub-taggen `v1.0.7`, inte från den senaste koden på `main`.

Du kan också ange CTID:

```bash
CTID=140 VERSION=1.0.7 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

Standard:
- CTID: första lediga ID från 200
- Hostname: `project-planer`
- Port: `8080`
- Debian 12
- 1 CPU
- 1 GB RAM
- 8 GB disk

## Uppgradering

Kör på Proxmox-host:

```bash
CTID=140 VERSION=1.0.7 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/update-lxc.sh)"
```

## Excel-export

Öppna ett projekt och välj **Exportera Excel**.

Arbetsboken innehåller:
- `Projektplan`
- `Sammanfattning`
- `Milstolpar`

Excel-filen har autofilter, frysta rubriker, kolumnbredder, datumformat och progressfält.

## Backup

Databasen ligger i:

```text
/opt/project-plan/data/projectplan.db
```

Vid uppgradering skapas backup i:

```text
/opt/project-plan/backups/
```

## Manuell servicehantering

Inne i LXC:

```bash
systemctl status project-plan
systemctl restart project-plan
journalctl -u project-plan -f
```

## Avinstallera

På Proxmox-host:

```bash
pct stop <CTID>
pct destroy <CTID>
```

## Publicera ny version

Kör på Windows:

```powershell
.\PUBLISH.cmd
```

### Vad v1.0.7 ändrar

v1.0.7 kräver **inte** GitHub-behörigheten `workflow`.

Tidigare version lade till `.github/workflows/validate.yml`. GitHub blockerar
push av en sådan fil när den Personal Access Token som används saknar
`workflow`-scope.

I v1.0.7:

- `.github/workflows/validate.yml` är borttagen.
- Om filen ligger kvar från v1.0.6 tar `PUBLISH.ps1` bort den automatiskt.
- `dist/` ligger i `.gitignore`.
- Om `dist/` råkade bli Git-trackad av v1.0.6 tas den automatiskt bort ur Git-index.
- Release-ZIP byggs fortfarande i `dist/` och laddas upp till GitHub Release.
- Python-validering sker lokalt och måste lyckas innan publicering.
- GitHub-autentisering, `fetch`, `push`, taggning och release sker som tidigare.

Det gör att en vanlig token med repository-innehållsbehörighet räcker för
publish-flödet.

### Installera v1.0.7 på Proxmox

```bash
VERSION=1.0.7 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

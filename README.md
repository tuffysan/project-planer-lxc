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
VERSION=1.0.9 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

Installern hämtar applikationen från GitHub-taggen `v1.0.9`, inte från den senaste koden på `main`.

Du kan också ange CTID:

```bash
CTID=140 VERSION=1.0.9 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
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
CTID=140 VERSION=1.0.9 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/update-lxc.sh)"
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

Kör:

```powershell
.\PUBLISH.cmd
```

### Nytt i v1.0.9

Release-steget är korrigerat.

I v1.0.8 skickades release notes direkt via `gh release create --notes`.
Installationskommandot innehåller `curl -fsSL`, vilket kunde feltolkas av
kommandoraden så att GitHub CLI såg `-fsSL` som en egen flagga.

v1.0.9 använder i stället en temporär Markdown-fil och `--notes-file`.

Dessutom kan publish-scriptet nu återhämta sig om Git-taggen redan finns men
själva GitHub Release saknas. Då återanvänds taggen och endast releasen skapas.

### Installera

```bash
VERSION=1.0.9 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

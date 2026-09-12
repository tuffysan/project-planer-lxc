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
VERSION=1.0.6 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

Installern hämtar applikationen från GitHub-taggen `v1.0.6`, inte från den senaste koden på `main`.

Du kan också ange CTID:

```bash
CTID=140 VERSION=1.0.6 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
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
CTID=140 VERSION=1.0.6 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/update-lxc.sh)"
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

### GitHub-autentisering

v1.0.6 hanterar även gamla eller felaktiga `GITHUB_TOKEN` / `GH_TOKEN`
miljövariabler.

GitHub CLI prioriterar sådana variabler framför sin sparade inloggning.
Om en sådan token finns men inte fungerar gör publish-scriptet därför:

1. testar befintlig GitHub-autentisering,
2. upptäcker om `GH_TOKEN` eller `GITHUB_TOKEN` stör inloggningen,
3. rensar variabeln endast i den aktuella publish-processen,
4. testar sparad `gh`-inloggning igen,
5. öppnar webbinloggning om det fortfarande behövs,
6. kör `gh auth setup-git`,
7. fortsätter med Git fetch/push och release.

Dina permanenta Windows-miljövariabler ändras inte.

### Installera v1.0.6 på Proxmox

```bash
VERSION=1.0.6 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

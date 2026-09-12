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
bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-plan-lxc/main/install-lxc.sh)"
```

Standard:
- CTID: första lediga ID från 200
- Hostname: `project-plan`
- Port: `8080`
- Debian 12
- 1 CPU
- 1 GB RAM
- 8 GB disk

Du kan styra värden:

```bash
CTID=140 \
HOSTNAME=project-plan \
STORAGE=local-lvm \
BRIDGE=vmbr0 \
IP_CONFIG=dhcp \
bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-plan-lxc/main/install-lxc.sh)"
```

Efter installation visas URL till webbappen.

## Uppgradering

Kör på Proxmox-host:

```bash
CTID=140 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-plan-lxc/main/update-lxc.sh)"
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

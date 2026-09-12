# Project Planer v14.0.4 — Update Verification Fix

Fixar ett quoting-fel i `update-lxc.sh` som gjorde att själva applikationsuppgraderingen lyckades
men att den sista verifieringen avslutades med:

`unexpected EOF while looking for matching "'"`

## Viktigt
Loggen från v14.0.3 visade att själva appen faktiskt hade uppgraderats korrekt:
- Current release: `/opt/project-plan/releases/v14.0.3`
- Current venv: `/opt/project-plan/venvs/v14.0.3`
- `/health` rapporterade `14.0.3`

Felet låg enbart i den sista verifieringsraden i updateraren.

## Ändring
- Tar bort känslig nested `python3 -c` quoting.
- Runtime-version läses nu ur `/health` med `sed`.
- Uppgraderingen avslutas bara som lyckad om runtime-versionen matchar vald release.

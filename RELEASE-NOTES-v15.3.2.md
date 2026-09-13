# Project Planer v15.3.2 — Excel Comment Runtime Fix

## Fix
v15.3.1 stoppades korrekt av installationens riktiga XLSX-smoketest.

Multi-Project Excel använder en kommentar på cell A1 i bladet Projekt för att
förklara hur automatiska projektkoder fungerar, men `Comment` från openpyxl
saknades i importerna.

Felet var:

`NameError: name 'Comment' is not defined`

v15.3.2 lägger till:

`from openpyxl.comments import Comment`

## Verifiering
Installations-smoketestet för både single-project och multi-project Excel ligger
kvar och måste passera innan releasen aktiveras.

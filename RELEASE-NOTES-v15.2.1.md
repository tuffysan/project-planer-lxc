# Project Planer v15.2.1 — Language Encoding Fix

## Root cause fixed
The language menu corruption was caused by Windows PowerShell 5 reading `app/app.py`
as the system ANSI code page during `PUBLISH.ps1`. UTF-8 emoji/characters were then
written back as already-corrupted text.

## Changes
- `PUBLISH.ps1` now uses explicit .NET UTF-8 read/write functions.
- Publishing aborts if common mojibake sequences (`Ã`, `Â`, `ðŸ`) are detected.
- Emoji country flags have been removed from the language menu.
- Languages now use robust ASCII badges: SV, EN, DE, NO, DA, FI.
- HTML/text responses explicitly declare `charset=utf-8`.
- Existing v15.2 controlled planning, resource leveling and Azure DevOps two-way sync remain unchanged.

This release fixes the source of the corruption, not only the CSS/menu rendering.

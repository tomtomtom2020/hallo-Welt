# Notizen: MS-Access-Formular (responsive)

Stand: 2026-11-16 – umgesetzt in `access/` (alle offenen Fragen mit Standard beantwortet).

## Umsetzung
- `access/modResponsive.bas` – Standardmodul. Import: VBA-Editor (Alt+F11) > Datei > Datei importieren.
- `access/Form_frmBeispiel.txt` – Code für das Formularmodul (Load + Resize).
- Buttons: Marke/Tag `flex:1`, `flex:2`, … im Formularfuß → Umbruch wie CSS flex-wrap.
- Felder: Marke/Tag `stretch` → wachsen bis zum rechten Rand.
- Ausrichtung wählbar: `flexLeft` (Standard), `flexCenter`, `flexRight`.
- Nicht in echtem Access getestet (keine Access-Umgebung vorhanden).

## Arbeitsweise
- Nichts doppelt einlesen: bereits gelesene Dateien/Routinen nur einmal laden.
- Bestehende Routinen (VBA-Module) übernehmen statt neu schreiben.

## Anforderungen
- Formular in MS Access.
- Responsive: Formular passt sich der Fenstergröße an.
- Buttons ordnen sich je nach Formularbreite automatisch an (flex-artig: nebeneinander, bei Platzmangel Umbruch in nächste Zeile).

## Mögliche Umsetzung (später)
- Access-Bordmittel: Anker (`HorizontalAnchor` / `VerticalAnchor`) für Felder.
- Für Button-Umbruch: VBA-Routine im `Form_Resize`-Ereignis, die Buttons nach `InsideWidth` zeilenweise positioniert (Abstand, Mindestbreite).

## Offene Fragen
- Wo liegen die bestehenden Routinen / die .accdb? (In den Repos nicht gefunden.)
- Welche Buttons / Felder enthält das Formular?
- Ausrichtung der Buttons: links, rechts, zentriert, oder unten fix?

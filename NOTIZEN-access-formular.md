# Notizen: MS-Access-Formular (responsive)

Stand: 2026-11-16 – **noch nicht programmieren**, nur Anforderungen festhalten.

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

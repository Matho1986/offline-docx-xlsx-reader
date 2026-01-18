# Offline DOCX/XLSX Reader

Eine werbefreie, vollständig offline arbeitende Android-App zum **Anzeigen von .docx-, .xlsx- und .pdf-Dateien** (nur Lesen). Die Dateien werden lokal über `content://`-URIs geöffnet – ohne Cloud, ohne Tracking, ohne Netzwerkzugriffe.

## Features

- ✅ Öffnen über **Datei-Dialog (Storage Access Framework)**
- ✅ Unterstützung für **"Öffnen mit" / Teilen** (`ACTION_VIEW`, `ACTION_SEND`)
- ✅ **DOCX-Anzeige** mit einfacher Formatierung (Überschriften, Absätze, Listen)
- ✅ **XLSX-Anzeige** mit Sheet-Auswahl (Dropdown) und Tabellenansicht
- ✅ **PDF-Anzeige** (Bonus, nur Darstellung; Text-Export ggf. eingeschränkt)
- ✅ **Text markieren, kopieren und teilen** (DOCX/XLSX)
- ✅ **Alles kopieren** (DOCX: Text, XLSX: TSV des aktuellen Sheets)
- ✅ **Teilen** der kopierten Inhalte als `text/plain`
- ✅ **Unterstützt Bildschirmrotation** (Dokument bleibt sichtbar)
- ✅ **Offline & werbefrei** (keine Netzwerkzugriffe)

## Status

- MVP abgeschlossen
- Stabil für den Alltagsgebrauch

## Getestet auf

- Android Smartphone (z. B. Samsung, aktuelles Android)

## Screenshots

> Platzhalter – Screenshots folgen

- `screenshots/docx.png`
- `screenshots/xlsx.png`

## Build & Run

1. Repository klonen
2. In **Android Studio** öffnen
3. Gradle-Sync ausführen
4. App starten (minSdk 26)

CLI (optional):

```bash
./gradlew :app:assembleDebug
```

## Datenschutz

- Keine Werbung
- Kein Tracking
- Keine Cloud
- Komplett offline

Alle Dokumente bleiben lokal auf deinem Gerät.

## Roadmap (optional)

- 🔍 Volltextsuche in Dokumenten
- 📄 PDF-Export
- 🔤 Schriftgröße anpassen
- 🔊 Vorlesen (TTS)

## Lizenz

MIT – siehe `LICENSE` im Repository.

# Robustes Portfolio – Zeitphasenmodell V5.0

**Bestehendes Portfolio importieren, Zeitphasen planen & Operative Manager steuern**

---

## 🎯 Was ist neu in Version 5.0 (Zeitphasenmodell)?

- **⏱️ Zeitphasenmodell-Layout:** Die Schichten 1 bis 6 stehen jetzt direkt unter den jeweiligen Zeitphasen 1 bis 6 (von der Kapitalreserve in Zeitphase 1 bis hin zum Engsten Benchmarking in Zeitphase 6).
- **🏦 Tagesgeld & Unternehmerisches Risiko:** Diese beiden Bausteine stehen nebeneinander und sind losgelöst von den Zeitphasen platziert.
- **🏷️ Haupt- & Sub-Headlines:**
  - Haupt-Headline: **„Zeitphasenmodell“**
  - Sub-Headline: **„Denken in Zeitphasen und Managementansätzen“** mit der Unterzeile *„Sie planen Ihre Ziele - wir die passende Strategie dazu.“*
- **📋 Rechte Spalte („Auswahl operativer Manager“):**
  - Einzeilige Überschrift direkt über der Gesamtgewichtungs-Fortschrittsanzeige
  - 4 Aktions-Icons rechtsbündig in der ersten Zeile: ⬆️ Setup laden, ⬇️ Setup speichern, ↺ Reset, 💾 PDF Export
- **🌍 Funktion „Länder aktualisieren“:**
  - Administrator-Funktion zur zentralen Aktualisierung der Ländergewichtungen (systemgeschützt).
- **📥 V4.4 Depot Import Engine & Delisted Fonds:**
  - Liest CSV-, Excel- (.xlsx) und PDF-Depotauszüge
  - Übernimmt delistete Fonds als **„Von Empfehlungsliste genommen“** mit interaktivem Hover-Popup (WKN, Beschreibung, Einmalbeitrag, Sparrate)

---

## ✅ Voraussetzungen

Keine Installation erforderlich. Das Dashboard ist eine reine HTML/JS-Webseite – kein Cloud-Dienst, keine Registrierung notwendig.

| System | Empfohlener Browser |
|--------|---------------------|
| 🍎 macOS | Chrome, Safari, Firefox |
| 🪟 Windows | Chrome, Edge, Firefox |
| 🐧 Linux | Chrome, Firefox |

> **Datensicherheit:** Alle Daten bleiben ausschließlich lokal im Browser (LocalStorage). Es werden keine Kundendaten ins Internet übertragen.

---

## 🚀 Dashboard öffnen

Die Datei **`index.html`** im Browser öffnen:

- **macOS:** Doppelklick auf `index.html`, oder Rechtsklick → „Öffnen mit“ → gewünschter Browser  
- **Windows:** Doppelklick auf `index.html`, oder Rechtsklick → „Öffnen mit“ → Chrome / Edge  

> **Tipp:** Bei Aktualisierungen den Browser-Cache mit **Cmd+Shift+R** (macOS) bzw. **Strg+Shift+R** (Windows) leeren, damit Änderungen an `app_v5.0.js` sofort wirksam werden.

---

## 🧭 Zeitphasen & Schichten-Zuordnung

| Zeitphase | Dauer | Schicht / Baustein | Anlageschwerpunkt |
|-----------|-------|--------------------|-------------------|
| **Zeitphase 1** | < 1 Jahr | **Kapitalreserve** | Kurzfristige Kasse & Geldmarkt |
| **Zeitphase 2** | > 2 Jahre | **Defensive Vermögensverwalter** | Stabilitätsorientiert & Anleihen |
| **Zeitphase 3** | > 4 Jahre | **Ausgewogene Vermögensverwalter** | Balance aus Chance & Sicherheit |
| **Zeitphase 4** | > 6 Jahre | **Dynamische Vermögensverwalter** | Chancenorientierte Vermögensverwalter |
| **Zeitphase 5** | > 10 Jahre | **Märkte Weites Benchmarking** | Weite Aktienindizes & Globales Benchmarking |
| **Zeitphase 6** | > 10 Jahre | **Märkte Enges Benchmarking** | Fokus-Aktienindizes & Enges Benchmarking |
| *Losgelöst* | - | **Tagesgeld** | Kurzfristige Liquiditätsreserve |
| *Losgelöst* | - | **Unternehmerisches Risiko** | Themen, Beteiligungen, Spezialitäten |

---

## 📥 Bestehendes Depot laden – Schritt für Schritt

Der empfohlene Weg ist der **CSV- oder Excel-Export aus dem MLP-Portal**.

### ⭐ Empfohlen: Depot über CSV-, Excel- oder PDF-Datei importieren

1. Dashboard (`index.html`) im Browser öffnen  
2. Oben rechts auf den Button **„Depot laden (CSV / Excel)“** klicken  
3. Im Dateiauswahldialog die exportierte CSV-, XLSX- oder PDF-Datei auswählen  
4. Das **Import-Ergebnis-Fenster** zeigt eine Zusammenfassung:
   - **Erkannte Fonds** aus der VEM-Datenbank
   - **Delistete Fonds** („Von Empfehlungsliste genommen“) mit automatischer Schichtzuordnung
5. Auf **„▶ Jetzt ins Portfolio laden“** klicken
6. Fertig – Fonds und Beträge sind eingetragen, die Gesamtgewichtung wird neu berechnet.

---

## 🌍 Ländergewichtungen aktualisieren

Die Funktion **„Länder aktualisieren“** steht ausschließlich dem **Systemadministrator** zur Verfügung und ist passwortgeschützt. 

Diese Funktion dient dazu, die Informationen zur Ländergewichtung der Fonds aus der **Empfehlungsliste / MLP-Vermögensdepotliste** zentral zu sammeln und in regelmäßigen Abständen zu aktualisieren.

---

## 📋 Rechte Spalte: „Auswahl operativer Manager“

Im rechten Panel verwalten Sie die aktiven Portfoliopositionen:

| Symbol | Funktion |
|--------|----------|
| ⬆ | **Setup laden** – eine gespeicherte `.json`-Datei importieren |
| ⬇ | **Setup speichern** – aktuelles Portfolio als `.json`-Datei sichern |
| ↺ | **Reset** – aktuelles Portfolio zurücksetzen (mit Bestätigungsdialog) |
| 💾 | **PDF-Export** – gewählte Manager-Liste als PDF exportieren |

---

## 🔴 Hard Reset – bei Problemen mit gespeicherten Daten

Der rote **„Hard Reset“**-Button in der Navigationsleiste löscht **alle** im Browser gespeicherten Daten (aktives Portfolio, Bibliothek, Empfehlungslisten-Zustände) und lädt die Anwendung sauber neu.

---

## 🗂️ Dateien in Version 5.0

| Datei | Beschreibung |
|-------|-------------|
| `index.html` | Hauptanwendung (Zeitphasenmodell V5.0) – im Browser öffnen |
| `styles_v5.0.css` | Stylesheet für Zeitphasen-Layout, Kacheln und Modals |
| `app_v5.0.js` | Anwendungslogik V5.0 (Zeitphasen, Import, Administrator-Zugang, Bibliothek) |
| `fund_data.js` | Fondsdatenbank (WKN, Anlageklassen, Top-5-Länderdaten) |
| `crawler_server.py` | Lokaler Crawler-Server für Länderaktualisierung (Port 8765) |
| `update_country_data.py` | Scraper-Skript für Fondsgesellschaften |
| `README.md` | Diese Anleitung |
| `VEM Depots/` | Ordner mit Beispiel-Importdateien (CSV, XLSX, PDF) |

---

## 📋 Changelog

### V5.0 (September 2026) – Zeitphasenmodell
- **✨ Zeitphasen-Architektur:** Direkte Zuordnung der Schichten 1-6 zu den Zeitphasen 1-6.
- **✨ Entkoppelte Bausteine:** Tagesgeld und Unternehmerisches Risiko als eigenständige Bausteine platziert.
- **✨ Layout & Header:** Neue Headlines „Zeitphasenmodell“ & „Denken in Zeitphasen und Managementansätzen“.
- **✨ Manager-Panel:** Zweizeilige Header-Struktur mit einzeiliger Überschrift über der Gesamtgewichtung und rechtsbündigen Aktions-Icons.
- **✨ Administrator-Zugang:** Systemgeschützte Aktualisierung der Ländergewichtungen für Fonds aus der Empfehlungsliste.
- **✨ Delisted Fonds Popup:** Interaktive Hover-Vorschau aller vom Empfehlungsblatt genommenen Fonds mit WKN & Beträgen.

---

*Robustes Portfolio · Zeitphasenmodell V5.0 · Stand: September 2026*

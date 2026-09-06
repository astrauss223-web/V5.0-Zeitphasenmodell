# Robustes Portfolio – Dashboard V4.4

**Bestehendes Portfolio importieren & Zuzahlungen planen**

---

## ✅ Voraussetzungen

Keine Installation erforderlich. Das Dashboard ist eine reine HTML-Webseite – kein Server, kein Cloud-Dienst, keine Internetverbindung notwendig (außer für die erste Anzeige von Schrift- und Symbol-CDNs).

| System | Empfohlener Browser |
|--------|---------------------|
| 🍎 macOS | Chrome, Safari, Firefox |
| 🪟 Windows | Chrome, Edge, Firefox |
| 🐧 Linux | Chrome, Firefox |

> **Datensicherheit:** Alle Daten bleiben ausschließlich lokal im Browser (LocalStorage). Es werden keine Daten ins Internet übertragen.

---

## 🚀 Dashboard öffnen

Die Datei **`index.html`** im Browser öffnen:

- **macOS:** Doppelklick auf `index.html`, oder Rechtsklick → „Öffnen mit" → gewünschter Browser  
- **Windows:** Doppelklick auf `index.html`, oder Rechtsklick → „Öffnen mit" → Chrome / Edge  

> **Tipp:** Bei wiederholtem Öffnen den Browser-Cache mit **Cmd+Shift+R** (macOS) bzw. **Strg+Shift+R** (Windows) leeren, damit Änderungen an `app_v4.4.js` sofort sichtbar werden.

---

## 📥 Bestehendes Depot laden – Schritt für Schritt

Der empfohlene Weg ist der **CSV- oder Excel-Export aus dem MLP-Portal**.

### ⭐ Empfohlen: Depot über CSV- oder Excel-Datei importieren

**Export aus dem MLP-Portal:**

1. Im MLP-Portal einloggen und die Depotübersicht öffnen
2. Den **CSV-Export** oder **Excel-Export (.xlsx)** herunterladen – die Datei enthält WKN-Nummern und investierte Beträge
3. Die Datei muss nicht umbenannt werden

**Depot ins Dashboard laden:**

1. Dashboard (`index.html`) im Browser öffnen  
2. Oben rechts auf den Button **„Depot laden (CSV / Excel)"** klicken  
3. Im Dateiauswahldialog die exportierte CSV- oder XLSX-Datei auswählen  
4. Ein **Import-Ergebnis-Fenster** erscheint mit einer Zusammenfassung:
   - **Anzahl erkannter Fonds** (direkt aus der VEM-Datenbank zugeordnet)
   - **Fonds mit eindeutiger Schicht** → werden automatisch in „Von Empfehlungsliste genommen" eingetragen
   - **Fonds mit unklarer Schicht** → werden als „Fonds delistet"-Badge gespeichert
5. Auf **„▶ Jetzt ins Portfolio laden"** klicken
6. Fertig – alle Fonds und Beträge sind eingetragen, das Gesamtvermögen wird korrekt berechnet

> **Encodings:** Das Dashboard erkennt sowohl **UTF-8** als auch **Windows-1252 / Latin-1** (MLP-Portal-Standard) automatisch. Beides funktioniert.

> **Hinweis zum Dateiformat:** Einige Exportdateien aus dem MLP-Portal haben ein leicht abweichendes Format, bei dem der Betrag des ersten Fonds eine Zeile oberhalb der WKN steht. Das Dashboard erkennt dieses Sonderformat automatisch und verarbeitet es korrekt.

---

### 📄 Alternative: JSON-Datei laden (Bibliothek)

Wenn ein Import bereits einmal gemacht und gespeichert wurde:

1. Auf den Button **„Bibliothek"** klicken  
2. Gespeicherte Portfolios anzeigen, per **„▶ Laden"** wiederherstellen  
3. Oder: eine `.json`-Datei per **„JSON-Datei importieren"** in die Bibliothek aufnehmen

> Die Bibliothek speichert Portfolios im Browser-LocalStorage und bleibt auch nach dem Schließen erhalten.

---

### 📋 Alternative: Setup-Datei laden (Panel-Buttons)

Im Panel „Mein Portfolio" (rechte Seite):

| Symbol | Funktion |
|--------|----------|
| ⬆ | **Setup laden** – eine gespeicherte `.json`-Datei importieren |
| ⬇ | **Setup speichern** – aktuelles Portfolio als `.json`-Datei sichern |
| ↺ | **Reset** – aktuelles Portfolio zurücksetzen (Bibliothek bleibt erhalten) |
| 🖨️ | Aktuelle Fondsliste drucken / als PDF exportieren |

---

### 📑 Fallback: PDF-Import

Maschinenlesbare PDFs (mit echter Textebene) werden ebenfalls unterstützt – einfach über **„Depot laden (CSV / Excel)"** auswählen, der Dialog akzeptiert auch PDF-Dateien.

> ⚠️ Eingescannte PDFs ohne Textebene können **nicht** automatisch gelesen werden. In diesem Fall bitte den CSV-Export aus dem MLP-Portal verwenden.

---

## 🔴 Hard Reset – bei Problemen mit gespeicherten Daten

Im Header oben rechts befindet sich ein roter **„Hard Reset"**-Button. Dieser löscht **alle** im Browser gespeicherten App-Daten:

- Aktives Portfolio & alle Beträge
- Depot-Bibliothek (alle gespeicherten Mappeneinträge)
- Delistete Fonds & Empfehlungsliste-Daten

**Wann verwenden:**
- Die angezeigte Summe stimmt nicht mit der Depot-Datei überein
- Das Dashboard zeigt nach einem Update noch alte, fehlerhafte Werte
- Ein Import-Fehler hat inkonsistente Daten hinterlassen

**Ablauf nach einem Hard Reset:**
1. Seite lädt leer neu
2. Direkt die CSV- oder XLSX-Datei neu importieren → korrekte Summe wird angezeigt
3. Optional: erneut in der Bibliothek speichern

> ⚠️ Der Hard Reset kann **nicht rückgängig** gemacht werden. Vorher wichtige Portfolios per ⬇ als JSON sichern.

---

## 🗂️ Was passiert mit nicht erkannten Fonds?

Das Dashboard unterscheidet zwei Fälle:

### 1. Schicht eindeutig erkennbar (z. B. „Aktien Welt", „Vermögensverwalter ausgewogen")

Der Betrag wird automatisch in das Feld **„Von Empfehlungsliste genommen"** der passenden Schicht eingetragen. Dieses Feld erscheint am Ende jeder Schicht im Portfolio-Panel (nur wenn ein Betrag vorhanden ist) und ist editierbar.

Beim **Hovern über das Feld** erscheint ein Popup mit den einzelnen betroffenen Fonds:
- Fondsname
- WKN
- Anlageschwerpunkt aus der CSV
- Einmalbetrag / Sparrate

### 2. Schicht nicht eindeutig

Diese Fonds werden als **🔶 Fonds delistet**-Badge in der Kapitalreservefonds-Schicht angezeigt.  
Beim Überfahren mit der Maus erscheint ein Popup mit:
- Fondsname
- WKN
- Anlageschwerpunkt aus der CSV
- Einmalbetrag / Sparrate

---

## 📊 Anlageplanung & Budgetlogik nach dem Import

Nach einem erfolgreichen Import gilt folgende Logik:

| Feld | Inhalt |
|------|--------|
| **Anzulegendes Gesamtvermögen** | Wird automatisch auf die **Summe aller importierten Beträge** gesetzt – einschließlich „Von Empfehlungsliste genommen"-Fonds |
| **Einmal Verplant** | Identisch mit dem Gesamtvermögen → **Verbleibend = 0,00 €** |
| **Von Empfehlungsliste genommen** | Ist Teil des Gesamtvermögens und wird in der jeweiligen Schicht separat ausgewiesen |

**Workflow nach dem Import:**
1. Gesamtvermögen und Einmal Verplant sind deckungsgleich (Verbleibend = 0 €)
2. Das Gesamtvermögen kann manuell **erhöht** werden (z. B. um geplante Zuzahlungen)
3. Der entstehende Restbetrag (Verbleibend) kann dann auf neue Schichten verteilt werden

> **Hinweis zur Gesamtgewichtung:** Die prozentualen Anteile im Panel „Mein Portfolio" beziehen sich **nur auf die regulär zugeordneten Fonds** – die Empfehlungsliste-Beträge fließen nicht in die %-Berechnung ein, damit die Summe korrekt 100 % ergibt.

---

## 🔁 Deduplication beim Import

Einige Fonds sind sowohl in **Kapitalreservefonds** als auch in **Tagesgeld** hinterlegt (z. B. iShares iBonds, ZinsPlus, Basis-Fonds I Nachhaltig). Beim Import wird jede WKN **nur einmal** zugeordnet – der **Kapitalreservefonds hat Vorrang** vor Tagesgeld.

---

## 💾 Depot-Bibliothek

Die **Depot-Bibliothek** (Button oben rechts) ist eine Verwaltungs-Zentralstation für gespeicherte Portfolios:

| Funktion | Beschreibung |
|----------|-------------|
| **Laden** | Portfolio aus der Bibliothek ins Dashboard laden |
| **JSON** | Portfolio als `.json`-Datei herunterladen (z. B. für andere Geräte) |
| **🗑️** | Eintrag aus der Bibliothek löschen |
| **JSON-Datei importieren** | Eine auf einem anderen Gerät exportierte `.json`-Datei in die Bibliothek aufnehmen |

> Nach einem erfolgreichen CSV-Import kann das Portfolio direkt in der Bibliothek gespeichert werden – Name ist frei wählbar.

---

## 🏗️ Turm-Schichten (unten → oben)

| Schicht | Anlagestrategie |
|---------|-----------------|
| Tagesgeld | Kurzfristige Einlagen |
| Kapitalreservefonds | Kasse / Wertsicherung |
| Defensive Vermögensverwalter | Stabilitätsorientiert |
| Ausgewogene Vermögensverwalter | Balance aus Chance & Sicherheit |
| Dynamische Vermögensverwalter | Weites Benchmarking auf Aktien |
| Märkte | **Weites & Enges Benchmarking auf Aktien** |
| Spezialitäten / Themen | Immobilien, PE, Gold, Beteiligungen |

### Von Empfehlungsliste genommen

Jede Schicht enthält am Ende ein editierbares Feld **„Von Empfehlungsliste genommen"** (orangefarben). Das Feld erscheint **nur, wenn ein Betrag vorhanden ist** (entweder durch automatischen Import oder manuelle Eingabe).

Beim Hovern über das Feld erscheint ein **Popup mit den Einzelfonds**, die nicht mehr auf der VEM-Empfehlungsliste stehen – analog zur Darstellung im Import-Fenster.

---

## 🗂️ Dateien in diesem Ordner

| Datei | Beschreibung |
|-------|-------------|
| `index.html` | Hauptdatei – im Browser öffnen |
| `styles_v4.4.css` | Stylesheet / Design |
| `app_v4.4.js` | Anwendungslogik (Import, Bibliothek, Panel, Hard Reset) |
| `fund_data.js` | Fondsdatenbank (WKN, Anlageklassen, Länderdaten) |
| `portfolio_import.json` | Beispiel-Portfolio (VEM AW, Stand 15.04.2026) |
| `crawler_server.py` | Hintergrundservice zum Aktualisieren von Ländergewichtungen |
| `README.md` | Diese Anleitung |
| `VEM Depots/` | Ordner mit Import-Dateien (CSV, XLSX, PDF) |

---

## ❓ Häufige Fragen

**F: Muss ich die CSV-Datei zuerst laden, bevor ich auf „Setup laden" klicke?**  
A: **Nein.** Beide Funktionen sind unabhängig. Der CSV-Import und der JSON-Setup-Import funktionieren jeweils für sich.

**F: Was ist der Unterschied zwischen dem ↺ Reset und dem roten Hard Reset?**  
A:  
- **↺ Reset** (Panel) = Löscht nur das aktive Portfolio (Fonds & Beträge). Die **Bibliothek bleibt erhalten**.  
- **🔴 Hard Reset** (Header) = Löscht **alles** – auch die Bibliothek und alle gespeicherten Zustände. Seite lädt danach neu.

**F: Warum erscheinen nach dem CSV-Import manche Fonds nicht im Portfolio?**  
A: Fonds, die nicht in der VEM-Datenbank sind, werden entweder als „Von Empfehlungsliste genommen" (wenn die Schicht eindeutig ist) oder als „Fonds delistet"-Badge (wenn die Schicht unklar ist) behandelt.

**F: Die angezeigte Gesamtsumme stimmt nicht mit meiner Depot-Datei überein – was tun?**  
A: Den roten **„Hard Reset"**-Button oben rechts klicken, bestätigen, und anschließend die CSV- oder XLSX-Datei **neu importieren**. Der korrigierte Parser liest ab V4.4 alle Fonds korrekt aus.

**F: Funktioniert das Dashboard auf Windows genauso wie auf dem Mac?**  
A: **Ja.** Es ist eine reine Webseite – im Browser öffnen, fertig. Kein Unterschied zwischen Betriebssystemen.

**F: Werden meine Daten gespeichert oder übertragen?**  
A: Alle Daten bleiben nur im **Browser-LocalStorage** des jeweiligen Geräts. Bei geleertem Browser-Cache gehen sie verloren – daher regelmäßig per ⬇ als JSON sichern oder in die Bibliothek speichern.

**F: Was ist der Unterschied zwischen Bibliothek speichern und Setup speichern (⬇)?**  
A:  
- **Bibliothek** = Speichert im Browser-LocalStorage, bleibt auf diesem Gerät erhalten, kein Download  
- **Setup speichern (⬇)** = Lädt eine `.json`-Datei herunter, portierbar auf andere Geräte

**F: Welche Datei-Formate werden beim Import unterstützt?**  
A: **CSV** (MLP-Portal-Export, UTF-8 und Latin-1/Windows-1252), **Excel (.xlsx/.xls)** und **PDF** (nur maschinenlesbare PDFs mit Textebene).

---

## 📋 Changelog

### V4.4 (April 2026) – Bugfixes & Verbesserungen

- **🐛 Fix: CSV-Parser** – Erkennt jetzt das Sonderformat, bei dem der Betrag des ersten Fonds in einer eigenen Zeile oberhalb der WKN steht (betrifft z. B. ZinsPlus / A0MUWS). Dadurch wurde dieser Betrag zuvor immer übergangen.
- **🐛 Fix: XLSX-Parser** – Liest Beträge korrekt aus Spalte 4 (mit `€`-Zeichen und non-breaking spaces). Erkennt WKN-Zellen die als Zahl (Integer) statt als String gespeichert sind. Unterstützt ebenfalls das Sonderformat (Betrag in Vorherzeile).
- **🐛 Fix: Gesamtvermögen-Berechnung** – `Anzulegendes Gesamtvermögen` wird jetzt immer als Summe **aller** importierten `fundInvestments` berechnet (inkl. Empfehlungsliste-Fonds). Zuvor wurden Empfehlungsliste-Beträge nicht addiert.
- **🐛 Fix: Bibliothek-Import** – Beim Laden eines Portfolios aus der Bibliothek wird `totalInvestment` immer neu aus `fundInvestments` berechnet – veraltete gespeicherte Werte werden überschrieben.
- **✨ Neu: Hard Reset** – Roter Button im Header löscht alle App-Daten aus dem Browser-Speicher und lädt die Seite neu. Behebt dauerhaft inkonsistente Datenzustände ohne manuelle Browser-Eingriffe.

---

*Version 4.4 · April 2026 – Letzte Änderung: 23.04.2026*

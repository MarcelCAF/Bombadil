# BOMBADIL – Arbeitsanweisung

**Paketverwaltung & OrcaScan-Anbindung**

- **Version:** 1.0.88
- **Datum:** 10. Juni 2026
- **Erstellt für:** Apothekenmitarbeiter

---

## Inhalt

1. [Überblick](#1-überblick)
2. [Programm starten & Oberfläche](#2-programm-starten--oberfläche)
3. [Allgemeine Bedienung aller Tabellen](#3-allgemeine-bedienung-aller-tabellen)
4. [Report – Startansicht](#4-report--startansicht)
5. [Statistik](#5-statistik)
6. [PU heute](#6-pu-heute)
7. [Abholbereit](#7-abholbereit)
8. [Zahlung offen](#8-zahlung-offen)
9. [> 7 Tage](#9--7-tage)
10. [Kissel > 2W](#10-kissel--2w)
11. [Unstimmigkeiten](#11-unstimmigkeiten)
12. [Gestern](#12-gestern)
13. [DHL Express & DHL heute](#13-dhl-express--dhl-heute)
14. [DHL Sendungssuche (Gimli)](#14-dhl-sendungssuche-gimli)
15. [PU Sendungssuche (Legolas)](#15-pu-sendungssuche-legolas)
16. [Cleanup](#16-cleanup)
17. [Datensicherung & Cloud-Backup](#17-datensicherung--cloud-backup)
18. [Einstellungen & Master-PC](#18-einstellungen--master-pc)
19. [Update-Mechanismus](#19-update-mechanismus)
20. [Häufige Fragen (FAQ)](#20-häufige-fragen-faq)

---

## 1. Überblick

Bombadil ist eine Windows-Desktop-Anwendung zur Verwaltung von Paketzustellungen und Pickup-Aufträgen (PU). Sie verbindet sich automatisch mit **OrcaScan** (dem Scan-System) und **Google Drive** (für Archiv und Backups).

**Hauptaufgaben:**

- Echtzeit-Übersicht über alle abholbereiten Pakete
- Verwaltung der täglichen Pickup-Aufträge („PU heute") mit Tour-Zuweisung (T1/T2)
- Abgleich der Tagesboten-Liste mit der Abholer-Datenbank
- DHL-Sendungsverfolgung (Normal und Express)
- **Sendungssuche** für DHL-Pakete (Gimli) und Pickups (Legolas)
- Statusänderungen per Rechtsklick (Retoure, Storno, Abgeholt, Bezahlt)
- Statistiken, Diagramme und Excel-Exporte
- Automatisches Update beim Start (von GitHub)
- **Automatische tägliche Datensicherung in die Cloud** (läuft auch bei ausgeschaltetem PC)

> **Fachbegriffe kurz erklärt:**
> - **OrcaScan** = das System, in das die Pakete eingescannt werden. Bombadil liest die Daten von dort.
> - **Abholer_DB** = die Datenbank mit allen Pickup-Paketen (Name, Status, Zeitpunkte, Ziel-Kiosk).
> - **Barcode / Sendungsnummer / PU-ID** = die Nummer auf dem Paket-Aufkleber.
> - **Tagesbote** = die Liste der Pakete, die heute auf Tour gehen.

---

## 2. Programm starten & Oberfläche

### 2.1 Starten

- **Doppelklick** auf `Bombadil.py` oder die zugehörige `.bat`-Datei.
- Beim Start prüft Bombadil automatisch, ob eine **neue Version** auf GitHub vorliegt, und aktualisiert sich ggf. selbst (siehe Abschnitt 19).
- Danach erscheint ein **Ladescreen** (dunkler Hintergrund, grüner Fortschrittsbalken 🌿) mit wechselnden Tipps, während die Daten geladen werden. Das **erste** Laden dauert etwas länger – das ist normal.
- Im **Fenstertitel** steht die aktuelle Versionsnummer (z. B. „Bombadil v1.0.88").

### 2.2 Header (oben)

Zeigt links Logo und Programmname, rechts die **Auto-Refresh-Buttons**. Ist ein Button grün, lädt Bombadil die Daten regelmäßig automatisch nach. Darunter steht die Uhrzeit des letzten Ladevorgangs.

### 2.3 Seitenmenü (Navigation, links)

Dunkelblaue Leiste, der aktive Bereich ist hervorgehoben. Die Einträge sind in drei Gruppen gegliedert:

| Gruppe | Einträge |
|---|---|
| **ÜBERSICHT** | 📊 Report · 📊 Statistik |
| **PAKETE** | 🚐 PU heute · 📦 Abholbereit · 💰 Zahlung offen · ⚠️ > 7 Tage · 🏪 Kissel > 2W · 🔍 Unstimmigkeiten · 📅 Gestern |
| **DHL** | 🚚 DHL Express · 🚛 DHL heute · 🔍 DHL Sendungssuche (Gimli) · 📦 PU Sendungssuche (Legolas) |

**Steuerung unten:** **Neu laden (F5)** lädt alle Daten frisch aus OrcaScan, **Cleanup** archiviert abgeschlossene Pakete (siehe Abschnitt 16).

### 2.4 Tastenkürzel

- **F5** – Daten neu laden
- **Strg+C** – markierte Zellen kopieren
- **Strg+V** – einfügen (in editierbaren Zellen)
- **Rechtsklick** – Kontextmenü (Aktionen, Kopieren, Sortieren …)

---

## 3. Allgemeine Bedienung aller Tabellen

- **Suche oben:** filtert über alle Spalten gleichzeitig.
- **Editierbare Zellen:** Doppelklick → Wert tippen → Enter (wird sofort in OrcaScan gespeichert).
- **Sortierung:** Klick auf eine Spaltenüberschrift; Standard meist nach Wartezeit/Status (dringendste zuerst).
- **Farb-Legende** erscheint **rechts neben der Tabelle** und erklärt die Farben.

### 3.1 Rechtsklick-Menü (statt Buttons)

In den Aktions-Tabs (Abholbereit, > 7 Tage, Kissel, Zahlung offen) erfolgen Statusänderungen per **Rechtsklick**:

- **Einzelne Zeile:** Rechtsklick → Aktion auswählen.
- **Mehrere Zeilen:** Zeilen erst mit **Strg-Klick** markieren, dann Rechtsklick auf eine markierte Zeile → Aktion wirkt auf alle.
- **Keine Zeile markiert:** Der Rechtsklick wählt automatisch die angeklickte Zeile.

### 3.2 Sicherheitsfrage bei Mehrfach-Aktionen

Werden **mehr als 3 Pakete** in einem Schritt geändert (Retoure / Storno / Abgeholt), erscheint eine Rückfrage:

> „Du möchtest X Pakete auf 'Y' setzen. Bist du sicher?"

So werden versehentliche Massen-Änderungen vermieden. Bei 1–3 Zeilen erscheint kein Dialog.

---

## 4. Report – Startansicht

Startansicht mit Echtzeit-Kacheln und Balkendiagramm.

### 4.1 Kacheln

Klick auf eine Kachel öffnet die Detailliste. **Trend-Pfeil** neben der Zahl: ▲ grün = gestiegen, ▼ rot = gesunken (gegenüber Vortag).

- **PU heute:** zeigt den Fortschritt der heutigen Pickups (verpackt / offen).
- **Kissel > 2W:** wird rot, sobald Einträge vorhanden sind (Warnung).

### 4.2 Balkendiagramm

Zeigt die Abholungen der letzten 7 Tage. Hover über einen Balken zeigt die genaue Anzahl.

---

## 5. Statistik

Kombinierte Auswertung für **PU** (Pickups) und **DHL** (Normal + Express). Pro Bereich gibt es ein Diagramm mit Kacheln.

### 5.1 Ansichten (umschaltbar)

- **Tagesansicht** – letzte 4 Wochen, tageweise
- **Wochenübersicht** – letzte 4 Kalenderwochen
- **Monatsübersicht** – letzte 6 Monate
- **Zeitraum** – ein **frei wählbarer Zeitraum**: Von-/Bis-Datum über den Mini-Kalender wählen, dann „Anzeigen". Die Summe für den Zeitraum wird angezeigt.

### 5.2 PU Statistik

- Kacheln: Lieferungen und Abholungen (Woche/Monat)
- **Zielkiosk-Tabelle** (rechts): Pakete je Kiosk (umschaltbar Woche / Monat / 6 Monate)

### 5.3 DHL Statistik

- Kacheln: Gesamt, DHL Normal, DHL Express, Pickup, Same day, Next day (Woche + Monat)
- Daten aus Live-OrcaScan + Google-Drive-Archiv

### 5.4 Monatsziel

In der **Monatsübersicht** auf einen Monat im Diagramm klicken → Ziel eintragen. Das Ziel erscheint als Linie im Diagramm.

> ℹ️ **Wie die Statistik entsteht:** Nur der **Master-PC** berechnet die Statistik und legt sie auf Google Drive ab. Alle anderen PCs zeigen exakt diese Zahlen – so sehen alle dasselbe. Bei Anzeige-Problemen hilft meist ein Neuaufbau des Zwischenspeichers (Cache) auf dem Master-PC.

---

## 6. PU heute

Zeigt die heutigen Pickup-Aufträge, abgeglichen mit der Abholer-Datenbank, mit automatischer **Tour-Zuweisung**:

- **Tour 1 (T1):** verpackt **vor** 11:16 Uhr
- **Tour 2 (T2):** verpackt **ab** 11:16 Uhr

### 6.1 Buttons

- **📋 Tagesbote** – öffnet/schließt das Seitenpanel mit dem Tagesboten-Abgleich
- **Tour 1 / Tour 2 abgeschickt** – setzt die Pakete der jeweiligen Tour auf Abholbereit
- **Export Tour 1 / Export Tour 2** – Excel-Export der jeweiligen Tour (zusätzlich auf Drive + NAS abgelegt)

### 6.2 Spalten & Farben

| Spalte | Inhalt |
|---|---|
| Tour | T1 (Hellblau) / T2 (Hellgrün) |
| Paket-Barcode | eindeutige Paket-ID |
| Name | Empfänger |
| Packstatus | Status aus Tagesbote |
| In DB | ✔ wenn in Abholer_DB vorhanden |
| Verpackt / Abholbereit | Zeitpunkte aus Abholer_DB |
| Ziel-Kiosk | – |

Zeilenfarbe = Status: 🟢 Salbeigrün = abholbereit · 🔵 Stahlblau = abgeholt · 🟠 Apricot = Retoure · 🟡 Buttergelb = verpackt · 🌸 Altrosa = offen.

### 6.3 Tour manuell umstellen ⭐ (neu)

Manchmal muss eine Tour nach dem Stempeln korrigiert werden:

> **Rechtsklick auf die Tour-Zelle** → **T1**, **T2** oder **Tour entfernen**.

Die Umstellung betrifft beide betroffenen Touren (das Paket wechselt sauber von einer zur anderen).

### 6.4 Weitere Rechtsklick-Aktionen

- **→ Abholbereit setzen** (einzelne Zeile)
- **🗑️ Aus Tagesboten löschen** – entfernt das Paket **nur** aus dem Tagesboten-Sheet; die Abholer_DB bleibt unverändert.

### 6.5 Tagesbote-Seitenpanel

Klick auf **📋 Tagesbote** öffnet rechts ein Panel zum Abgleich der Tagesboten-Liste mit der Abholer_DB. Fehlende Pakete lassen sich direkt übernehmen.

### 6.6 Auto-Fix

Zeigt der Tagesbote ein Paket als „Offen", die Abholer_DB hat aber einen „Verpackt"-Zeitpunkt von **HEUTE**, korrigiert Bombadil den Status automatisch auf „Verpackt" (in der Anzeige und in OrcaScan). Greift nur bei heutigen Zeitstempeln.

---

## 7. Abholbereit

Vollständige Liste aller Pakete mit Status „Abholbereit" (älteste zuerst).

**Spalten:** Paket-Barcode · Abholbereit_At · Name · Ziel-Kiosk (per Doppelklick editierbar) · Wartezeit.

**Legende:** 🟡 kein Datum · 🟠 > 3 Tage · 🔴 > 7 Tage.

**Aktionen (Rechtsklick):** ↩ Retoure · ✗ Storno · ✓ Abgeholt. (Bei > 3 Zeilen Sicherheitsfrage.)

---

## 8. Zahlung offen

Alle Pakete mit offener Zahlung (Status „Vor Ort", „Unbezahlt" oder „Offen") – unabhängig vom Paketstatus.

**Spalten:** Paket-Barcode · Name · Bestellwert (per Doppelklick editierbar) · Status · Abholbereit_At · Wartezeit.

**Legende (Farbe nach Status):** 🌸 Altrosa = offen · 🟡 Buttergelb = verpackt · 🟢 Salbeigrün = abholbereit · 🔵 Stahlblau = abgeholt.

**Sortierung:** Abholbereit → Verpackt → Offen → Abgeholt (je Gruppe älteste zuerst).

**Aktionen:**

- **Bestellwert eintragen:** Doppelklick auf die Zelle → Wert → Enter (sofort in OrcaScan gespeichert).
- **Als bezahlt markieren:** Rechtsklick → **✔ Als bezahlt markieren** → Zeile verschwindet sofort.

> ℹ️ Die T1/T2-Tournummer im Status erscheint, sobald „PU heute" einmal geladen wurde.

---

## 9. > 7 Tage

Abholbereite Pakete seit mehr als 7 Tagen (älteste zuerst).

**Legende:** 🟠 7–14 Tage · 🔴 15–30 Tage · 🟥 > 30 Tage.

**Aktionen:** Rechtsklick → Retoure / Storno / Abgeholt (wie Abholbereit).

---

## 10. Kissel > 2W

Pakete am Standort Kissel, die länger als **2 Wochen (10 Werktage)** abholbereit sind. Werktage = Mo–Fr (Wochenenden zählen nicht).

**Legende:** 🟠 2–6 Wochen · 🔴 6 Wochen–3 Monate · 🟥 > 3 Monate.

**Aktionen:** Rechtsklick → Retoure / Storno / Abgeholt.

---

## 11. Unstimmigkeiten

Pakete mit Status „Verpackt" in der Abholer_DB, die **nicht** im heutigen Tagesboten erscheinen (mögliche Ursache: Paket eingetragen, aber kein PU-Auftrag angelegt).

**Aktion:** „Alle Pakete → Abholbereit setzen". Wird automatisch aktualisiert, wenn Abholer_DB oder PU heute neu geladen wird.

---

## 12. Gestern

Alle gestern abgeholten Pakete (neueste zuerst). Keine weiteren Aktionen.

---

## 13. DHL Express & DHL heute

- **🚚 DHL Express:** DHL-Express-Scans von heute (Barcode + Scan-Zeitpunkt).
- **🚛 DHL heute:** DHL-Normal-Scans von heute (aus dem OrcaScan DHL_Normal-Sheet). Export als `YYMMDD.xlsx` möglich.

---

## 14. DHL Sendungssuche (Gimli)

Sucht eine **DHL-Sendungsnummer** in allen geladenen DHL-Daten – **Express + Normal**, inklusive **Drive-Archiv** und **Live-OrcaScan**.

**Bedienung:**

1. Sendungsnummer (Barcode) eingeben – **mit oder ohne** führende `00`.
2. „Suchen" klicken (oder Enter).
3. Beim **ersten Mal** erscheint ein grüner Ladebalken, während Live-Daten + Archiv eingelesen werden. Weitere Suchen kommen sofort.
4. Treffer zeigen Barcode, Typ (Normal/Express), Scan-Zeitpunkt und Quelle.

> ℹ️ Auch ein **Teil** der Sendungsnummer reicht für die Suche.

---

## 15. PU Sendungssuche (Legolas)

Sucht eine **PU-ID (Paket-Barcode)** in der Abholer_DB (**Live + Drive-Archiv**) und zeigt den **Status-Verlauf als Zeitleiste**:

```
Scan → Verpackt → Abholbereit → Abgeholt
```

Pro Treffer werden außerdem Name, Status, Ziel-Kiosk, Tour (T1/T2) und Quelle angezeigt. Erledigte Schritte sind grün, offene grau markiert. Auch hier reicht ein Teil des Barcodes.

> ℹ️ Beide Sendungssuchen funktionieren **auf jedem PC** (nicht nur am Master-PC), sofern der PC OrcaScan- und Drive-Zugang hat.

---

## 16. Cleanup

Löscht abgeschlossene Einträge aus OrcaScan und archiviert sie in Google Drive.

**Lösch-Kriterien:**

- Abgeholt_At gesetzt + älter als 3 Werktage
- Status „Abgeholt" + kein Abholbereit_At + Scan-Datum älter als 7 Tage

**Ablauf:** Daten laden → Kandidaten filtern → Archiv-Excel nach Drive hochladen → Einträge aus OrcaScan löschen. Nach Abschluss erscheint eine Statusmeldung.

---

## 17. Datensicherung & Cloud-Backup

Damit der PC abends **ausgeschaltet** werden kann, sichert sich Bombadil auf mehreren Wegen ab.

### 17.1 Automatisches Cloud-Backup (PC-unabhängig) ⭐

- Läuft **täglich automatisch in der GitHub-Cloud** – komplett **ohne** PC und ohne geöffnetes Bombadil.
- Gesichert werden: **Abholer_DB, DHL Normal, DHL Express, Tagesbote** → als Excel auf Google Drive.
- Geplante Zeit: **abends** (kann sich GitHub-bedingt um bis zu ~1–2 Std. verschieben – unkritisch).

> ✅ **Das bedeutet:** Du kannst den PC abends bedenkenlos ausschalten. Das Backup macht GitHub von selbst.

### 17.2 Weitere Sicherungen (wenn der PC läuft)

- Der **Master-PC** sichert zusätzlich auf Google Drive und das NAS.
- Die **Statistik** wird mehrschichtig abgesichert: Live-OrcaScan + Drive-Backups + Tour-Dateien.

### 17.3 Datenquellen im Überblick

| Quelle | Wofür |
|---|---|
| **OrcaScan (Live)** | aktuelle Pakete (Abholer_DB, DHL Normal/Express, Tagesbote) |
| **Google Drive** | Archiv + tägliche Backups + Statistik-Cache + Tourlisten |
| **NAS** | zusätzliches Schreibziel für Backups und Tour-Export |

---

## 18. Einstellungen & Master-PC

### 18.1 Menü Einstellungen

- **Exportordner wählen** – Ordner für Excel/CSV-Exporte festlegen
- **Einstellungen zurücksetzen** – Exportordner auf Downloads zurücksetzen
- **Manuelles Backup** – auf jedem PC über das Menü auslösbar

### 18.2 Master-PC

Nur **ein** PC (Marcels Rechner) ist als „Master" eingestellt. Dieser **berechnet die Statistik** und macht die **automatischen Backups**. Alle anderen PCs zeigen die Daten nur an.

> ⚠️ Bitte die Master-Einstellung nicht ohne Rücksprache ändern – sonst würden mehrere PCs gleichzeitig Backups hochladen.

### 18.3 Menü Hilfe

- **Funktionsübersicht** – Kurzbeschreibung aller Tabs
- **Tastenkürzel** – alle Kürzel
- **Über Bombadil** – Versionsinformationen

---

## 19. Update-Mechanismus

Bombadil prüft beim Start automatisch, ob eine neue Version auf GitHub verfügbar ist.

1. Start → Bombadil vergleicht die lokale Version mit GitHub.
2. Bei neuerer Version: Dialog „Update verfügbar – Jetzt aktualisieren?" → **Ja** klicken.
3. Bombadil sichert die alte Datei als `Bombadil.backup_<alte_Version>.py`.
4. Lädt die neue `Bombadil.py` und startet sich neu.

> ℹ️ Funktioniert auch auf Netzlaufwerken (NAS). Ohne Internet wird die alte Version still gestartet.

> ⚠️ Hat ein Kollege eine sehr alte Version (vor dem Update-Fix), die aktuelle `Bombadil.py` einmalig manuell rüberkopieren – ab dann läuft Auto-Update.

---

## 20. Häufige Fragen (FAQ)

**Bombadil zeigt veraltete Daten – was tun?**
F5 drücken oder „Neu laden". Der Auto-Refresh-Button sollte grün sein.

**Ein Paket zeigt falschen Status?**
Status in OrcaScan prüfen/korrigieren, dann F5.

**Tour-Zuweisung stimmt nicht?**
Tour wird aus dem Verpackt-Zeitpunkt berechnet. Notfalls in „PU heute" per Rechtsklick auf die Tour-Zelle manuell umstellen (Abschnitt 6.3).

**Die Sendungssuche zeigt keine Treffer / keinen Ladebalken?**
Beim ersten Suchen nach dem Start dauert es einen Moment (Live + Archiv laden). Erscheint dauerhaft nichts, fehlt ggf. der Drive-Zugang auf diesem PC – dann findet die Suche nur aktuelle (Live-)Pakete, kein Archiv.

**Die Statistik zeigt falsche/veraltete Zahlen?**
Die Zahlen kommen vom Master-PC. Hilft meist: Statistik-Cache neu aufbauen lassen (Master-PC).

**Ein Feiertag zeigt 0 Pakete – Fehler?**
Nein. An arbeitsfreien Tagen gibt es keine Scans → 0 ist korrekt.

**Läuft das Backup, wenn der PC aus ist?**
Ja – das Cloud-Backup läuft automatisch in der GitHub-Cloud (Abschnitt 17.1).

**Wie markiere ich ein Paket als bezahlt?**
Tab „Zahlung offen" → Rechtsklick → „✔ Als bezahlt markieren".

**Wie ändere ich den Status mehrerer Pakete auf einmal?**
Zeilen mit Strg-Klick markieren → Rechtsklick → Aktion. Bei > 3 Zeilen Sicherheitsfrage.

**Wo finde ich den Tagesboten-Abgleich?**
Tab „PU heute" → Button „📋 Tagesbote" → Seitenpanel klappt auf.

**Wo finde ich eine alte Bombadil-Version?**
Im selben Ordner liegt nach jedem Update eine Backup-Datei (`Bombadil.backup_<Version>.py`).

---

*Diese Arbeitsanweisung beschreibt den Stand der Bombadil-Version 1.0.88. Bei neuen Versionen können Funktionen hinzukommen oder sich ändern.*

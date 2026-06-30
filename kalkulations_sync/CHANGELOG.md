# Changelog

Alle nennenswerten Änderungen an `kalkulations_sync` werden hier dokumentiert.
Format: [Keep a Changelog](https://keepachangelog.com/de/1.0.0/) · Versionierung nach
Odoo-Manifest (`19.0.major.minor.patch`).

---

## [19.0.1.2.0] – 2026-06-30

### Changed
- **Internationalization.** The module's source language is now **English**. All
  user-facing strings (views, fields, wizard, messages) were translated to English.
- Added a complete **German translation** in `i18n/de.po` plus the extraction
  template `i18n/kalkulations_sync.pot`. German installations keep the previous
  German labels with no regression.
- Manifest re-licensed to **LGPL-3** for the open-source release; added `LICENSE`,
  English `README.md` and user guides under `docs/`.

### Notes
- Template header markers accept **English and German** labels
  (e.g. `[Quantity]` / `[Menge]`, `[Unit Price]` / `[Preis je Einheit]`).
- Boolean import still accepts `Yes`/`True`/`1` **and** German `Ja`/`Wahr`.

---

## [19.0.1.1.3] – 2026-06-18

### Behoben
- **Zweite Korruptionsursache: ungültiger `modified`-Zeitstempel.** Bei in UTC
  laufendem Server schreibt openpyxl die Core-Property „modified" als
  `…+00:00Z` — ungültiges W3CDTF (Offset *und* Z). Excel meldet das bei der
  „file level validation and repair". `sanitize_export_xlsx` entfernt das
  überzählige `Z` jetzt (→ gültiges `…+00:00`). Per Excel verifiziert:
  vollständig bereinigter Export öffnet ohne Reparaturdialog (`OPENED_CLEAN`).

---

## [19.0.1.1.2] – 2026-06-18

### Behoben
- **Export weiterhin als „beschädigt" gemeldet (Excel-Reparaturdialog) bei
  Vorlagen aus SharePoint.** Solche Vorlagen tragen in `docProps/custom.xml`
  SharePoint-Metadaten (`ContentTypeId`, `_dlc_DocIdItemGuid`,
  `MediaServiceImageTags`). openpyxl schreibt diesen Part beim Speichern in einer
  Form um, die Excel als „unlesbaren Inhalt" ablehnt (in Word die Meldung „custom
  XML elements no longer supported"). Der Export entfernt den Part und seine zwei
  Referenzen (Content-Types-Override + Root-Relationship) jetzt vollständig.
  Per Excel verifiziert: vorher Reparaturdialog, nachher sauberer Öffnen-Vorgang.
- `strip_empty_formula_caches` → `sanitize_export_xlsx` umbenannt/erweitert
  (führt beide Nachbearbeitungen in einem Durchgang aus).

---

## [19.0.1.1.1] – 2026-06-17

### Behoben
- **Korrupte Export-Datei / „Formelzelle ohne berechneten Wert" beim Import.**
  openpyxls `load(data_only=False)`→`save`-Zyklus hinterließ auf jeder Formelzelle
  einen leeren `<v></v>`-Cache-Knoten. Folge: Excel meldete die Datei als beschädigt
  („unlesbaren Inhalt wiederherstellen"), und da der Cache leer war, las der Import
  alle Formelspalten als leer → Fehlerstatus. Der Export entfernt diese leeren Knoten
  jetzt nachträglich (`strip_empty_formula_caches`) und setzt `fullCalcOnLoad`, sodass
  Excel die Datei sauber öffnet und die Formeln neu berechnet. Nach dem Speichern in
  Excel stehen echte Werte im Cache, die der Import korrekt zurückliest.

---

## [19.0.1.1.0] – 2026-06-17

### Hinzugefügt
- Gebündeltes Standard-Template `static/templates/kalkulation_template.xlsx`
  (Holztüren-Kalkulation mit OZ `gaeb_oz`, Positionsart `gaeb_pos_type`, Menge,
  Bezeichnung, EP sowie Zuschlags-/Deckungsbeitrags-Block).
- `post_init_hook`: Bei Frischinstallation wird das gebündelte Template automatisch
  in alle Firmen ohne eigenes Template vorinstalliert (Einstellungen → Verkauf →
  Kalkulations-Sync → Excel-Template). Bereits hochgeladene Templates bleiben unberührt.

---

## [19.0.1.0.1] – 2026-06-03

### Behoben
- Anlegen neuer Positionen (`N`-Zeilen) brach mit `ValueError: Invalid field
  'product_uom' in 'sale.order.line'` ab. In Odoo 19 heißt das Feld `product_uom_id`.
  Beim Setzen der Standardprodukt-Einheit wird jetzt der korrekte Feldname verwendet.

---

## [19.0.1.0.0] – Odoo 19

Aktueller Funktionsstand (Dokumentations-Baseline, erfasst 2026-06-03):

- Template-gesteuerter Excel-Export/-Import für `sale.order.line` (Standard- und Custom-Felder).
- Zwei Template-Varianten: `{{line.feldname}}`-Platzhalter (Variante A) und
  `[feldname]`-Header-Marker (Variante B).
- Import-Wizard mit Feld-Diff und Status (`changed`/`unchanged`/`error`/`new`/`missing`/`ignored`).
- Concurrency-Check über Export-Zeitstempel; neue Positionen über `N` in der ID-Spalte.
- Formel-Unterstützung (Zeilen-Offset pro Position, Summenformel-Expansion).
- Verstecktes Sheet `kalksync_meta` mit Zeitstempel, Auftrags-ID und Spalten-Mapping.
- Firmenfelder mit Prefix `amf_kalksync_*` auf `res.company` (Template, Dateiname,
  Standardprodukt für neue Positionen).
- Export/Import über Buttons im Auftrags-Header (Status `draft`/`sent`).

---

### Historie (v17-Linie, vor der Portierung)

- **17.0.1.2.0** – Export/Import über dedizierte Header-Buttons; `res.company`-Felder von
  `kalksync_*` auf `amf_kalksync_*` umbenannt (AMF-Prefix-Convention).
- **17.0.1.1.0** – Erstveröffentlichung: Template-gesteuerter Export/Import, Import-Wizard
  mit Diff-Ansicht, Concurrency-Check, neue Positionen, Formel-Support, Testsuite.

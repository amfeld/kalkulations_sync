# Kalkulations-Sync – Template-Vorbereitung

Das Excel-Template wird einmalig unter **Einstellungen → Verkauf → Kalkulationstemplate** hochgeladen. Beim Export füllt Odoo das Template pro Auftragsposition aus. Beim Import liest Odoo alle markierten Spalten zurück — unabhängig davon, von welchem Addon das Feld stammt.

---

## Das Wichtigste zuerst: Was wird importiert?

**Das Template bestimmt, was importiert wird — nicht der Code.**

Jedes beschreibbare Feld von `sale.order.line` kann importiert werden, sobald es im Template als Placeholder oder Marker auftaucht. Standard-Felder, Custom-Felder von anderen Addons (`x_gaeb_*`, `material_*`, `efb_*`) — alles funktioniert gleich.

Automatisch **nicht** importiert werden:
- Berechnete, nicht-gespeicherte Felder (z.B. `x_gaeb_gp = Menge × Preis`)
- Relationsfelder (Many2one, One2many, Many2many)
- Binary- und Serialized-Felder

---

## Pflichtbestandteil: `{{line.id}}`

In genau **einer Zelle** des Templates muss der Platzhalter `{{line.id}}` stehen. Diese Zelle markiert die **Masterzeile** — die Zeile, die Odoo für jede Position dupliziert.

- Ohne `{{line.id}}` bricht der Export mit einem Fehler ab.
- Die ID-Zelle darf im exportierten Excel **nie** verändert werden.

---

## Odoo-Werte ausgeben: `{{line.feld}}`

Jedes Feld von `sale.order.line` kann als Platzhalter in die Masterzeile geschrieben werden. Odoo befüllt die Zelle beim Export mit dem aktuellen Wert.

| Platzhalter | Bedeutung |
|---|---|
| `{{line.id}}` | Positions-ID (Pflicht) |
| `{{line.product_uom_qty}}` | Menge |
| `{{line.price_unit}}` | VK-Preis |
| `{{line.purchase_price}}` | EK-Preis (nur mit `sale_margin`) |
| `{{line.name}}` | Positionsbezeichnung |
| `{{line.x_gaeb_oz}}` | GAEB-Ordnungszahl |
| `{{line.x_gaeb_menge}}` | GAEB-Menge |
| `{{line.material_cost_per_unit}}` | Materialkosten/Einheit |
| `{{line.efb_lohn_pct_manual}}` | EFB-Lohnanteil (manuell) |
| `{{line.<beliebiges_feld>}}` | Jedes weitere Feld |

Platzhalter-Felder werden beim Export ausgefüllt **und** beim Import zurückgelesen — sofern das Feld beschreibbar ist.

---

## Importierbare Spalten einrichten — zwei Varianten

### Variante A — Platzhalter direkt in der Masterzeile (einfachste Methode)

Platzhalter direkt in die Masterzeile schreiben. Das System erkennt `{{line.feldname}}` automatisch als importierbare Spalte:

```
Spalte C:   {{line.product_uom_qty}}
Spalte D:   {{line.price_unit}}
Spalte E:   {{line.x_gaeb_menge}}
```

Beim Export wird der Wert eingetragen, beim Import der geänderte Wert zurückgelesen.

### Variante B — Marker in der Kopfzeile (für Formel-Spalten)

Wenn der Wert **nicht** direkt als Platzhalter in der Masterzeile steht, sondern per Formel berechnet wird, kann die Spalte trotzdem importiert werden. Dafür muss in einer Zeile **oberhalb der Masterzeile** ein Marker in eckigen Klammern stehen:

```
Zeile N-1 (Marker):  [product_uom_qty]   [price_unit]   [x_gaeb_menge]   GP %
Zeile N   (Anzeige): Menge               VK-Preis       LV-Menge         (keine Aktion)
Zeile N+1 (Master):  {{line.id}}         =Formel EP     =Formel Menge    =Formel GP
```

- Der Marker steht **in derselben Spalte** wie die Formel
- Der Prefix `line.` ist optional: `[price_unit]` und `[line.price_unit]` sind gleichwertig
- Deutsche Aliasnamen werden erkannt (Groß-/Kleinschreibung ignoriert):

| Marker | Entspricht |
|---|---|
| `[Menge]` | `product_uom_qty` |
| `[Preis je Einheit]` oder `[Preis je ME]` | `price_unit` |
| `[Kosten je Einheit]` oder `[Kosten je ME]` | `purchase_price` |
| `[Bezeichnung]` | `name` |
| `[x_gaeb_menge]` | direkt der technische Feldname |
| `[material_cost_per_unit]` | direkt der technische Feldname |

Spalten **ohne** Marker werden beim Import ignoriert (z.B. GP %-Formel-Spalte).

---

## Unterstützte Feldtypen

| Typ | Vergleich | Beispielfelder |
|---|---|---|
| Float / Integer / Monetary | Numerisch, Toleranz 1e-6 | `product_uom_qty`, `price_unit`, `x_gaeb_menge` |
| Char / Text / Html | Zeichenkette exakt | `name`, `x_gaeb_oz`, `x_gaeb_kurztext` |
| Boolean | Ja/Nein | `x_gaeb_manuell`, `efb_use_custom` |

**Boolean-Werte in Excel:** `Ja`, `ja`, `1`, `True`, `yes`, `wahr` → `True`; alles andere → `False`.

---

## Kopfzeilen-Platzhalter: `{{object.feld}}`

Außerhalb der Masterzeile können Auftrags-Felder eingebettet werden (werden beim Import ignoriert):

```
{{object.name}}              → Angebotsnummer
{{object.partner_id.name}}   → Kundenname
{{object.date_order}}        → Angebotsdatum
{{object.user_id.name}}      → Zuständiger Verkäufer
```

---

## Neue Positionen anlegen

Wenn in einer Zeile die ID-Zelle mit **`N`** (oder `n`) befüllt wird, erkennt der Import diese Zeile als neue Position und legt beim Bestätigen eine neue `sale.order.line` an.

| ID-Zelle | Verhalten |
|---|---|
| leer | Zeile wird stillschweigend ignoriert |
| `N` oder `n` | Neue Position wird angelegt |
| Zahl (bestehende ID) | Bestehende Position wird aktualisiert |
| anderer Text | Fehler → Import blockiert |

Alle importierbaren Felder aus der Zeile werden beim Anlegen der neuen Position gesetzt.

---

## Typischer Template-Aufbau

```
Zeile 1   │ Logo / Projektinfo / {{object.name}} / {{object.partner_id.name}}
Zeile 2   │ (leer oder Überschrift)
Zeile 3   │ [product_uom_qty]  [price_unit]  [x_gaeb_menge]  GP %
Zeile 4   │ Menge              VK-Preis      LV-Menge        (ignoriert)
Zeile 5   │ {{line.id}}  {{line.product_id.name}}  =Formel  =Formel  =GP-Formel
          │ ↑ Masterzeile
Zeile 6+  │ (Odoo fügt hier weitere Positionen ein, Formeln werden angepasst)
...
Zeile X   │ Summe: =SUM(C5:C5)   ← expandiert automatisch auf =SUM(C5:C[letzte])
```

- **Masterzeile** = die Zeile mit `{{line.id}}`
- **Marker-Zeile** = eine oder mehrere Zeilen oberhalb der Masterzeile (beliebig weit oben, nächstgelegene gewinnt)
- Odoo dupliziert die Masterzeile für jede Position und passt Formel-Zellbezüge automatisch an

---

## Was beachten beim Template-Erstellen

1. **`{{line.id}}` muss vorhanden sein** — ohne diese Zelle bricht der Export ab.
2. **Nur beschreibbare Felder importieren** — berechnete Felder (z.B. `x_gaeb_gp`) werden automatisch übersprungen, auch wenn sie im Template stehen.
3. **Marker in richtiger Spalte** — `[feldname]` muss exakt in der Spalte stehen, die beim Import zurückgelesen werden soll.
4. **ID-Spalte nicht verändern** — die Zelle mit der Positions-ID darf im exportierten Excel nicht editiert werden (außer für neue Positionen: `N`).
5. **Nach externer Bearbeitung speichern** — Formeln müssen in Excel geöffnet und gespeichert werden, bevor importiert wird (gecachte Formelwerte).
6. **`kalksync_meta`-Sheet nicht löschen** — das versteckte Sheet wird vom System beim Export erzeugt und beim Import benötigt.
7. **Feldname prüfen** — technische Feldnamen können in Odoo unter Einstellungen → Technisch → Felder nachgeschlagen werden. Der Feldname für den Marker ist identisch mit dem Python-Attributnamen auf `sale.order.line`.

---

## Checkliste vor dem Hochladen

- [ ] `{{line.id}}` ist in genau einer Zelle vorhanden
- [ ] Alle importierbaren Spalten haben entweder `{{line.feldname}}` in der Masterzeile oder `[feldname]` in einer darüberliegenden Zeile
- [ ] Kein Marker auf computed non-stored Felder (würde beim Import übersprungen)
- [ ] Formel-Spalten: Marker steht in derselben Spalte wie die Formel
- [ ] Die Datei lässt sich in Excel öffnen und alle Formeln werden berechnet
- [ ] Nach dem Export: Datei in Excel öffnen → Werte prüfen → speichern → erst dann importieren

---

## Mitgeliefertes Beispiel-Template

Unter `static/templates/kalkulation_template.xlsx` liegt ein einsatzbereites Starter-Template.
Es demonstriert beide Varianten gleichzeitig und kann direkt unter
**Einstellungen → Verkauf → Kalkulations-Sync → Excel-Template** hochgeladen werden.

### Aufbau des Starter-Templates

```
Zeile 1  │ A: Angebot:  B: {{object.name}}   C: Kunde:  D: {{object.partner_id.name}}
Zeile 2  │ A: ID   B: Produkt   C: Menge   D: VK-Preis   E: Summe   F: Bezeichnung
Zeile 3  │ A: {{line.id}}   B: {{line.product_id.name}}   C: {{line.product_uom_qty}}
         │ D: {{line.price_unit}}   E: =C3*D3   F: {{line.name}}
         ↑ Masterzeile
```

| Spalte | Inhalt | Importierbar? |
|--------|--------|---------------|
| A | `{{line.id}}` — Positions-ID (Pflicht, nie editieren) | nein |
| B | `{{line.product_id.name}}` — Produktname (Many2one, nur Ausgabe) | nein |
| C | `{{line.product_uom_qty}}` — Menge | **ja** |
| D | `{{line.price_unit}}` — VK-Preis | **ja** |
| E | `=C3*D3` — Summe (Formelzelle, kein Marker → nicht importiert) | nein |
| F | `{{line.name}}` — Positionsbezeichnung | **ja** |

Zeile 1 zeigt Kopfzeilen-Platzhalter (`{{object.*}}`): werden beim Export befüllt,
beim Import ignoriert.

Spalte E zeigt die Formel-Expansion: Odoo kopiert `=C3*D3` für jede weitere Position
mit angepasstem Zeilen-Offset (`=C4*D4`, `=C5*D5` usw.).

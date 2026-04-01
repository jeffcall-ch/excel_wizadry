# FFD Bar-Count Calculation — Beam Sections

_How the tool determines how many 6 m beam bars to order from cut-piece data in the BoM._

---

## Overview

The BoM lists individual cut pieces, e.g.:

| Qty | Cut length |
|-----|-----------|
| 4 | 900 mm |
| 2 | 1 800 mm |
| 7 | 125 mm |
| 1 | 2 400 mm |
| 3 | 450 mm |

The naive answer — `ceil( total_net_cut ÷ 6 000 )` — is always too low because it ignores kerf and end-loss. **First Fit Decreasing (FFD)** packs the pieces optimally into real bars and gives the true minimum.

---

## Parameters

| | Value |
|---|---|
| Physical bar length | 6 000 mm |
| End-loss (uncuttable bar end) | 100 mm |
| **Effective bar length** | **5 900 mm** |
| Kerf per cut | 3 mm |
| Workshop buffer | 10 % |

---

## Step-by-Step Example

**Step 1 — Net cut length**

`4×900 + 2×1800 + 7×125 + 1×2400 + 3×450 = 11 825 mm = 11.825 m`

**Step 2 — Naive minimum** _(reference only)_

`ceil(11 825 ÷ 6 000) = 2 bars` — impossible in practice (ignores waste).

**Step 3 — Sort pieces largest-first**

`[2400, 1800, 1800, 900, 900, 900, 900, 450, 450, 450, 125×7]`

Placing big pieces first leaves offcuts suited for smaller pieces — this is why FFD is efficient.

**Step 4 — FFD placement** _(effective bar = 5 900 mm, kerf = 3 mm)_

Each piece needs `piece + 3 mm` of remaining bar space. Place into the first bar that fits; open a new bar if none does.

| Piece | Slot needed | Bar 1 rem. | Bar 2 rem. | Bar 3 rem. |
|-------|-------------|-----------|-----------|-----------|
| 2 400 | 2 403 | 5900→**3497** | — | — |
| 1 800 | 1 803 | 3497→**1694** | — | — |
| 1 800 | 1 803 | 1694 too small | 5900→**4097** | — |
| 900 | 903 | 1694→**791** | — | — |
| 900 | 903 | 791 too small | 4097→**3194** | — |
| 900 | 903 | — | 3194→**2291** | — |
| 900 | 903 | — | 2291→**1388** | — |
| 450 | 453 | 791→**338** | — | — |
| 450 | 453 | 338 too small | 1388→**935** | — |
| 450 | 453 | — | 935→**482** | — |
| 125 | 128 | 338→**210** | — | — |
| 125 | 128 | 210→**82** | — | — |
| 125 | 128 | 82 too small | 482→**354** | — |
| 125 | 128 | — | 354→**226** | — |
| 125 | 128 | — | 226→**98** | — |
| 125 | 128 | 82 too small | 98 too small | 5900→**5772** |
| 125 | 128 | — | — | 5772→**5644** |

**FFD result: 3 bars**

**Step 5 — Apply 5% workshop buffer**

`ceil(3 × 1.05) = ceil(3.15) = `**4 bars ordered**

**Step 6 — Final outputs**

| Output | Value | Calculation |
|--------|-------|-------------|
| `Qty_pcs` | 2 | naive minimum |
| `Order_Qty_pcs` | 4 | FFD → buffer |
| `Spare_pcs` | 2 | 4 − 2 |
| `Total_Order_Length_m` | 24.0 m | 4 × 6 m |
| `Utilisation_pct` | 49.3 % | 11.825 ÷ 24.0 |
| `Order_Weight_kg` | scaled | net weight × (24.0 ÷ 11.825) |

---

## How the Margins Stack

```
Physical bar  6 000 mm
  − 100 mm end-loss  →  Effective bar  5 900 mm
     each piece + 3 mm kerf  →  FFD bar count (minimum)
   × 1.05, round up     →  ordered bars
```

---

---

# FFD Stabanzahl-Berechnung — Trägerprofile (DE)

_Wie das Tool ermittelt, wie viele 6-m-Träger auf Basis der Schnittliste bestellt werden müssen._

---

## Überblick

Die Stückliste enthält einzelne Schnittlängen, z. B.:

| Menge | Schnittlänge |
|-------|-------------|
| 4 | 900 mm |
| 2 | 1 800 mm |
| 7 | 125 mm |
| 1 | 2 400 mm |
| 3 | 450 mm |

Das einfache Minimum (`ceil( Gesamtschnittlänge ÷ 6 000 )`) ist immer zu niedrig — es ignoriert Sägeverlust und Endverlust. Der **First-Fit-Decreasing (FFD)**-Algorithmus packt die Teile optimal in echte Stäbe und liefert die tatsächliche Mindestanzahl.

---

## Parameter

| | Wert |
|---|---|
| Physische Stablänge | 6 000 mm |
| Endverlust (nicht sägebares Stabende) | 100 mm |
| **Nutzbare Stablänge** | **5 900 mm** |
| Sägeblattbreite (Kerf) je Schnitt | 3 mm |
| Werkstattpuffer | 10 % |

---

## Schritt-für-Schritt-Beispiel

**Schritt 1 — Netto-Schnittlänge**

`4×900 + 2×1800 + 7×125 + 1×2400 + 3×450 = 11 825 mm = 11,825 m`

**Schritt 2 — Naives Minimum** _(nur als Referenz)_

`ceil(11 825 ÷ 6 000) = 2 Stäbe` — in der Praxis nicht realisierbar (Verluste werden ignoriert).

**Schritt 3 — Teile absteigend sortieren**

`[2400, 1800, 1800, 900, 900, 900, 900, 450, 450, 450, 125×7]`

Große Teile zuerst → Restzuschnitte passen zu den kleinen Teilen — das macht FFD effizient.

**Schritt 4 — FFD-Platzierung** _(nutzbare Stablänge = 5 900 mm, Kerf = 3 mm)_

Jedes Teil benötigt `Teilgröße + 3 mm` Restplatz im Stab. Zuerst in den ersten passenden Stab legen; neuen Stab öffnen, falls keiner passt.

| Teil | Platzbedarf | Stab 1 Rest | Stab 2 Rest | Stab 3 Rest |
|------|------------|------------|------------|------------|
| 2 400 | 2 403 | 5900→**3497** | — | — |
| 1 800 | 1 803 | 3497→**1694** | — | — |
| 1 800 | 1 803 | 1694 zu klein | 5900→**4097** | — |
| 900 | 903 | 1694→**791** | — | — |
| 900 | 903 | 791 zu klein | 4097→**3194** | — |
| 900 | 903 | — | 3194→**2291** | — |
| 900 | 903 | — | 2291→**1388** | — |
| 450 | 453 | 791→**338** | — | — |
| 450 | 453 | 338 zu klein | 1388→**935** | — |
| 450 | 453 | — | 935→**482** | — |
| 125 | 128 | 338→**210** | — | — |
| 125 | 128 | 210→**82** | — | — |
| 125 | 128 | 82 zu klein | 482→**354** | — |
| 125 | 128 | — | 354→**226** | — |
| 125 | 128 | — | 226→**98** | — |
| 125 | 128 | 82 zu klein | 98 zu klein | 5900→**5772** |
| 125 | 128 | — | — | 5772→**5644** |

**FFD-Ergebnis: 3 Stäbe**

**Schritt 5 — 10 % Werkstattpuffer anwenden**

`ceil(3 × 1,10) = ceil(3,30) = `**4 Stäbe bestellt**

**Schritt 6 — Endergebnisse**

| Ausgabe | Wert | Berechnung |
|---------|------|-----------|
| `Qty_pcs` | 2 | naives Minimum |
| `Order_Qty_pcs` | 4 | FFD → Puffer |
| `Spare_pcs` | 2 | 4 − 2 |
| `Total_Order_Length_m` | 24,0 m | 4 × 6 m |
| `Utilisation_pct` | 49,3 % | 11,825 ÷ 24,0 |
| `Order_Weight_kg` | skaliert | Nettogewicht × (24,0 ÷ 11,825) |

---

## Wie die Verluste sich addieren

```
Physische Stablänge  6 000 mm
  − 100 mm Endverlust  →  Nutzbare Länge  5 900 mm
     jedes Teil + 3 mm Kerf  →  FFD-Stabanzahl (Minimum)
        × 1,10, aufrunden    →  bestellte Stäbe
```

# Invirtus — Bottoms-Up Market Sizing

A working Excel model that derives **TAM / SAM / SOM** for Invirtus from the
count of target buyers × realistic usage-based revenue (tests run × net price
per test) — **not** a top-down slice of a published HR-tech number.

## The answer (base case, USD)

| Ring | Value | Basis |
|------|-------|-------|
| **TAM** | **~$1.20B** | 3 buyer segments across all geos where the product works, full ACV (focus-6 grossed up by focus-6 share of the global universe). |
| **SAM** | **~$144M** | Focus-6 geos (France, UK, UAE, Morocco, KSA, Egypt), segments serviceable today. |
| **SOM (Yr 3)** | **~$2.2M** | Reachable in 3 years. **Channel-led (~56%)**: Manatal embedded attach, then direct self-serve/sales (~44%). |

**Lead with SOM.** The single biggest assumption is the **Year-3 channel attach
rate** on the ICP-relevant Manatal base (Assumptions §7); tests/account and the
$15 PAYG price are the next two swing factors (see Sensitivity).

## Workbook tabs

1. **Cover** — objective, one-line answer, the three rings, SOM split, how-to-read.
2. **Assumptions** — every driver as a single editable input (blue), grouped and
   colour-coded; researched inputs carry a source key (col E) tied to Sources.
3. **Build** — bottoms-up SAM (entities × %fit × tests/yr × net $/test), grossed
   up to TAM, and SOM via two routes side by side (channel + direct), plus a
   value-delivered read (tests → hires → $150k cost-of-bad-hire exposure).
4. **Sensitivity** — two two-way tables and a tornado on the three inputs that
   move SOM most: channel attach rate, tests/account, PAYG price.
5. **Sources** — one row per researched input with source, URL, date, confidence.

## Method

- **Segments (MECE):** (A) tech-focused recruiting agencies, (B) engineering-heavy
  in-house firms, (C) mid-size firms (50–250) with structured hiring. Agencies and
  the firms they serve are distinct buyers — no double count.
- **Revenue engine:** usage-based. `tests/account/yr = roles hired × candidates
  tested per role`. Net price per test comes from a pack mix (PAYG $15, Small −10%,
  Medium −15%, Large −20%) blended, plus an enterprise flat-fee tier.
- **SOM:** Channel = Manatal 10,000+ clients × ICP-relevant share × attach rate ×
  ACV × retention. Direct = penetration of the serviceable SAM (captures that share
  of segment-weighted SAM value), cross-checked against a traffic→trial→paid funnel
  with deliberately low trial-to-paid.

## Conventions (financial-modeling skill)

- **Single input colour:** blue font + pale fill = editable input. Black = formula.
- **No hardcodes in formula cells** — every number references an assumption cell.
- Researched figures flagged with a source key; assumptions state their logic.

## Open decisions (per brief — defaults applied)

- **Horizon:** SOM = 3 years. **Currency:** USD throughout.
- **Second ATS (Socium):** held as **upside only** (`SOCIUM_ON = 0`); set to `1`
  on the Assumptions tab to include it in the base.

## Rebuild

```bash
pip install openpyxl
python3 build/build_model.py     # writes Invirtus_Market_Sizing.xlsx
```

Formulas were validated end-to-end with the `formulas` library (0 error cells).
The softest inputs — use-case-fit % and the engineer-hiring share behind the
mid-size counts — are the ones to pressure-test before locking TAM.

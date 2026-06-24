# Invirtus — Bottoms-Up Revenue Potential

A working Excel model that derives Invirtus's **revenue potential** (TAM / SAM /
SOM equivalents) from the count of target buyers × realistic usage-based revenue
— **not** a top-down slice of a published HR-tech number. Two revenue streams:
**assessment packs** + **talent-DB monetization**.

## The answer (base case = breadth mode 1 "Engineers", USD)

| Tier | Assessments | + Talent-DB | **Total revenue potential** |
|------|-------------|-------------|------------------------------|
| **Full opportunity** (TAM) | ~$640M | ~$471M | **~$1.11B** |
| **Serviceable today** (SAM) | ~$77M | ~$56M | **~$133M** |
| **3-year reachable** (SOM) | ~$2.29M | ~$1.69M | **~$3.97M** |

- **Full opportunity** — 3 buyer segments, all geos where the product works, full ACV.
- **Serviceable today** — focus-6 geos (France, UK, UAE, Morocco, KSA, Egypt).
- **3-year reachable** — **assessment line is channel-led (~76% partner channel / ~24% direct)**;
  ties to the finalized BP. Talent-DB adds ~43% on top, consistent with the BP reaching
  ~45% DB revenue by 2030.

### Market-breadth toggle (Assumptions §1b)

The addressable universe scales with a single toggle, so the same model serves the
tech-only and broader-market views (addresses the "TAM seems low" feedback):

| Mode | Full opportunity | Serviceable | 3-yr reachable |
|------|------------------|-------------|----------------|
| **1 — Engineers only** (current) | ~$1.11B | ~$133M | ~$3.97M |
| **2 — White-collar hiring** | ~$4.98B | ~$597M | ~$7.31M |
| **3 — Narrower core** | ~$0.64B | ~$76M | ~$3.56M |

### Distribution = partner portfolio (Assumptions §7)

Channel is no longer Manatal-only. It's a toggleable portfolio — Manatal (live),
Socium, other HR-tech/ATS players, staffing-firm networks, HR consultancies, system
integrators — each with its own client base, ICP-relevant share, and attach rate
(~789 base accounts at year 3). A global attach scalar drives sensitivity.

> Pricing aligned to Aida's **finalized, pack-led business plan**: $14 base/test;
> packs PAYG / Small (20, −10%) / Medium (100, −15%) / **Enterprise (1,000, −20% = $11,200)**;
> blended net **$12.22/test**; no flat enterprise fee; 90% retention (BP base 93%).

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
  tested per role`. Net price per test comes from a pack mix (PAYG $14, Small −10%,
  Medium −15%, Enterprise 1,000-pack −20%) blended to ~$12.22/test. All tiers are
  per-test — no flat enterprise fee (matches the finalized BP).
- **Breadth toggle:** a 1/2/3 mode scales each segment's universe by a documented
  multiplier (engineers → white-collar → narrower core).
- **SOM:** Channel = Σ over partner portfolio (clients × ICP-relevant share × attach ×
  scalar) × ACV × retention. Direct = penetration of the serviceable SAM (captures that
  share of segment-weighted SAM value), cross-checked against a traffic→trial→paid funnel
  with deliberately low trial-to-paid.
- **Second stream — talent-DB monetization:** for each tier,
  `DB revenue = tests run × % profiles sourced from own DB × $30/profile`
  (mature 30% used; finalized BP ramps 1%→50%). Total revenue potential =
  assessments + talent-DB.

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

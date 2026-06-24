# Invirtus — Bottoms-Up Revenue Potential

A working Excel model that maps Invirtus's **revenue potential** from the count of
target buyers × realistic usage-based revenue
— **not** a top-down slice of a published HR-tech number. Two revenue streams:
**assessment packs** + **talent-DB monetization**.

## The answer (base case = breadth mode 1 "Engineers", USD)

| Tier | Assessments | + Talent-DB | **Total revenue potential** |
|------|-------------|-------------|------------------------------|
| **Full opportunity** | ~$640M | ~$471M | **~$1.11B** |
| **Serviceable today** | ~$77M | ~$56M | **~$133M** |
| **3-year reachable** | ~$2.29M | ~$1.69M | **~$3.97M** |

- **Full opportunity** — 3 buyer segments, all geos where the product works, full ACV.
- **Serviceable today** — focus-6 geos (France, UK, UAE, Morocco, KSA, Egypt).
- **3-year reachable** — **assessment line is channel-led (~76% partner channel / ~24% direct)**;
  ties to the finalized BP. Talent-DB adds ~43% on top, consistent with the BP reaching
  ~45% DB revenue by 2030.

### Market-breadth toggle (Assumptions §2)

The addressable universe scales with a single toggle, so the same model serves the
tech-only and broader-market views (addresses the "market seems too narrow" feedback):

| Mode | Full opportunity | Serviceable | 3-yr reachable |
|------|------------------|-------------|----------------|
| **1 — Engineers only** (current) | ~$1.11B | ~$133M | ~$3.97M |
| **2 — Some white-collar** | ~$2.41B | ~$289M | ~$5.09M |
| **3 — All white-collar** | ~$4.98B | ~$597M | ~$7.31M |

### Distribution = partner portfolio (Assumptions §4)

Channel is no longer Manatal-only. It's a toggleable portfolio — Manatal (live),
Socium, other HR-tech/ATS players, staffing-firm networks, HR consultancies, system
integrators — each with its own client base, ICP-relevant share, and attach rate
(~789 base accounts at year 3). A global attach scalar drives sensitivity.

> Pricing aligned to Aida's **finalized, pack-led business plan**: $14 base/test;
> packs PAYG / Small (20, −10%) / Medium (100, −15%) / **Enterprise (1,000, −20% = $11,200)**;
> blended net **$12.22/test**; no flat enterprise fee; 90% retention (BP base 93%).

**Lead with the 3-year reachable number.** The single biggest assumption is the
**Year-3 channel attach scalar** across the partner portfolio (Assumptions §4);
tests/account and the $14 base price are the next two swing factors (see Sensitivity).

## Workbook tabs

1. **Cover** — the three tiers, the 3-year breakdown, scenario levers, how-to-read.
2. **Assumptions** — six short sections: buyers, breadth toggle, value per account,
   channel portfolio, direct/reachability, talent-DB. Blue = input, black = formula.
3. **Build** — serviceable today (buyers × %addressable × ACV), grossed up to the full
   opportunity, the 3-year reachable tier (channel + direct), and the two-stream totals.
4. **Sensitivity** — one tornado on the 3-year number: attach scalar, price, tests/account, direct penetration.
5. **Sources** — one row per researched input with source, URL, date, confidence.

## Method

- **Segments (MECE):** (A) tech-focused recruiting agencies, (B) engineering-heavy
  in-house firms, (C) mid-size firms (50–250) with structured hiring. Agencies and
  the firms they serve are distinct buyers — no double count.
- **Value per account:** `ACV = tests/account/yr × net price/test`. One blended price
  (~$12.22 = $14 PAYG less 10–20% pack discounts at the BP mix); tests/account is a single
  input per segment (≈ roles × candidates).
- **Breadth toggle:** a 1/2/3 mode scales each segment's universe by a documented
  multiplier (engineers → some white-collar → all white-collar).
- **3-year reachable:** Channel = Σ over partner portfolio (clients × ICP-relevant share ×
  attach × scalar) × ACV × retention. Direct = penetration of the serviceable tier
  (captures that share of segment-weighted serviceable value).
- **Second stream — talent-DB monetization:** for each tier,
  `DB revenue = tests run × % profiles sourced from own DB × $30/profile`
  (mature 30% used; finalized BP ramps 1%→50%). Total revenue potential =
  assessments + talent-DB.

## Conventions (financial-modeling skill)

- **Single input colour:** blue font + pale fill = editable input. Black = formula.
- **No hardcodes in formula cells** — every number references an assumption cell.
- Researched figures flagged with a source key; assumptions state their logic.

## Open decisions (per brief — defaults applied)

- **Horizon:** reachable tier = 3 years. **Currency:** USD throughout.
- **Distribution partners:** each partner row on the Assumptions tab has an On/Off
  toggle (1/0); switch any partner in or out of the base.

## Rebuild

```bash
pip install openpyxl
python3 build/build_model.py     # writes Invirtus_Market_Sizing.xlsx
```

Formulas were validated end-to-end with the `formulas` library (0 error cells).
The softest inputs — use-case-fit % and the engineer-hiring share behind the
mid-size counts — are the ones to pressure-test before widening breadth.

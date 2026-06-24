#!/usr/bin/env python3
"""
Invirtus - Bottoms-Up Market Sizing Model builder.

Conventions (financial-modeling skill):
  - Single INPUT color: blue font (#0000FF) = hardcoded, editable assumption.
  - Formula cells: black font, never contain a hardcoded number (everything
    references an assumption cell on the Assumptions tab).
  - Researched inputs are flagged and carry a source key tying to the Sources tab.

Run:  python3 build/build_model.py
Out:  Invirtus_Market_Sizing.xlsx
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

# ----------------------------------------------------------------------------
# STYLE PALETTE
# ----------------------------------------------------------------------------
BLUE   = "FF0000FF"      # input font
BLACK  = "FF000000"      # formula font
GREEN  = "FF006100"      # cross-sheet link (used sparingly)
WHITE  = "FFFFFFFF"
NAVY   = "FF1F2A44"      # header band
LILAC  = "FFE7DDEA"      # section fill (matches deck)
CREAM  = "FFF3EFE9"      # cover bg accent
GREY   = "FF7F7F7F"

f_title   = Font(name="Calibri", size=22, bold=True, color="FF1F2A44")
f_sub     = Font(name="Calibri", size=11, italic=True, color=GREY)
f_h1      = Font(name="Calibri", size=13, bold=True, color=WHITE)
f_h2      = Font(name="Calibri", size=11, bold=True, color="FF1F2A44")
f_input   = Font(name="Calibri", size=10, color=BLUE)
f_calc    = Font(name="Calibri", size=10, color=BLACK)
f_calcb   = Font(name="Calibri", size=10, bold=True, color=BLACK)
f_label   = Font(name="Calibri", size=10, color=BLACK)
f_note    = Font(name="Calibri", size=9, italic=True, color=GREY)
f_big      = Font(name="Calibri", size=28, bold=True, color="FF1F2A44")
f_biglbl   = Font(name="Calibri", size=11, bold=True, color=GREY)

fill_h1   = PatternFill("solid", fgColor=NAVY)
fill_sec  = PatternFill("solid", fgColor=LILAC)
fill_cov  = PatternFill("solid", fgColor=CREAM)
fill_inp  = PatternFill("solid", fgColor="FFFFFDE7")   # pale yellow = editable
fill_card = PatternFill("solid", fgColor="FFFFFFFF")

thin = Side(style="thin", color="FFD9D2DE")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
btop = Border(top=Side(style="thin", color=GREY))

center = Alignment(horizontal="center", vertical="center")
left   = Alignment(horizontal="left", vertical="center", wrap_text=True)
right  = Alignment(horizontal="right", vertical="center")

PCT = '0.0%'
PCT0 = '0%'
USD = '#,##0'
USD2 = '$#,##0.00'
USDM = '$#,##0.0,,"M"'
USDB = '$#,##0.00,,,"B"'
NUM = '#,##0'

# ============================================================================
# DATA  --  every figure here is an INPUT (blue). Researched ones carry src key.
# ============================================================================

GEOS = ["France", "UK", "UAE", "Morocco", "KSA", "Egypt"]

# Segment A: tech-focused recruiting agencies (count of addressable entities)
AGENCIES = {
    "France": 500,    # src: INSEE SIRENE 7810Z/7820Z, tech subset ~500
    "UK":     3500,   # src: REC/IBISWorld ~30k total x ~10-15% tech
    "UAE":    100,    # src: SIA/Lusha ~500 prof. x ~20% eng/IT
    "Morocco":45,     # src: EcoActu / MIEPEEC triangulation
    "KSA":    50,     # src: GoodFirms/Clutch prof. segment subset
    "Egypt":  60,     # src: Egypt YellowPages/Clutch IT recruiters
}
# Segment B: engineering-heavy / tech-software-ICT companies (in-house)
TECHFIRMS = {
    "France": 95000,  # src: INSEE - software/IT-services working base
    "UK":     107000, # src: GOV.UK Digital & Tech Sector Plan 2026 (107,082)
    "UAE":    5000,   # src: DIC 4,000 / MAGNiTT 4,576
    "Morocco":2000,   # src: UM6P startups ~1,500 + APEBI ~1,000 offshoring
    "KSA":    1800,   # src: CST Manassa 1,000+ floor
    "Egypt":  2500,   # src: ITIDA 240 offshoring + MAGNiTT ~2,059 startups
}
# Segment C: mid-size firms (50-250) running structured tech/eng hiring
MIDSIZE = {
    "France": 7300,   # src: EC SME factsheet 20,896 medium x ~35%
    "UK":     13200,  # src: UK BPE 2024 37,800 medium x ~35%
    "UAE":    2800,   # src: ~20k medium x ~14% eng-hiring
    "Morocco":1250,   # src: OMTPME ~3,800 medium x ~30%
    "KSA":    2600,   # src: Monsha'at 18,723 medium x ~13%
    "Egypt":  2000,   # src: CAPMAS ~10k 50+ emp x ~20%
}

# Source keys per geo input (for Sources tab cross-ref)
SRC = {
    "France": "S1",  "UK": "S2", "UAE": "S3", "Morocco": "S4", "KSA": "S5", "Egypt": "S6",
}

wb = openpyxl.Workbook()

# ============================================================================
# SHEET: Assumptions  (built first so Build can reference its cells)
# ============================================================================
A = wb.active
A.title = "Assumptions"
A.sheet_view.showGridLines = False
A.column_dimensions["A"].width = 3
A.column_dimensions["B"].width = 46
A.column_dimensions["C"].width = 14
A.column_dimensions["D"].width = 12
A.column_dimensions["E"].width = 10
A.column_dimensions["F"].width = 52
A.column_dimensions["G"].width = 12
A.column_dimensions["H"].width = 7

ref = {}   # name -> "Assumptions!$C$n"
row = 1

def band(ws, r, text, span_to="F"):
    ws.cell(r, 2, text).font = f_h1
    for c in range(2, 7):
        ws.cell(r, c).fill = fill_h1
    ws.row_dimensions[r].height = 20

def secrow(ws, r, text):
    ws.cell(r, 2, text).font = f_h2
    for c in range(2, 7):
        ws.cell(r, c).fill = fill_sec

def put(name, label, value, unit, fmt, note="", is_input=True, src=""):
    """Write an assumption row; register cell ref under `name`."""
    global row
    r = row
    A.cell(r, 2, label).font = f_label
    A.cell(r, 2).alignment = left
    c = A.cell(r, 3, value)
    c.number_format = fmt
    if is_input:
        c.font = f_input
        c.fill = fill_inp
    else:
        c.font = f_calc
    c.alignment = right
    A.cell(r, 4, unit).font = f_note
    A.cell(r, 4).alignment = left
    A.cell(r, 5, src).font = f_note
    A.cell(r, 5).alignment = center
    A.cell(r, 6, note).font = f_note
    A.cell(r, 6).alignment = left
    if name:
        ref[name] = f"Assumptions!$C${r}"
    row += 1
    return f"Assumptions!$C${r}" if False else f"Assumptions!$C${r-1+1-0}"

# header
A.cell(row, 2, "INVIRTUS  -  Assumptions").font = f_title
row += 1
A.cell(row, 2, "All blue cells are editable inputs. Black cells are formulas. "
               "Col E = source key (see Sources tab). Currency = USD.").font = f_sub
A.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
row += 2

# column captions
for cidx, t in [(2,"Driver"),(3,"Value"),(4,"Unit"),(5,"Src"),(6,"Logic / flag")]:
    A.cell(row, cidx, t).font = f_h2
    A.cell(row, cidx).fill = fill_sec
row += 1

# ---- Section 1: addressable entities by geo & segment ----
band(A, row, "1.  Addressable entities  (researched)"); row += 1

for seg, label, data, note in [
    ("AG", "Tech-focused recruiting agencies", AGENCIES,
     "Tech/IT recruitment agencies."),
    ("TF", "Engineering-heavy firms (tech/software/ICT cos.)", TECHFIRMS,
     "Tech/software/ICT companies."),
    ("MS", "Mid-size firms w/ structured tech hiring (50-250)", MIDSIZE,
     "Mid-size firms (50-250) hiring engineers."),
]:
    secrow(A, row, f"{label}"); row += 1
    for g in GEOS:
        put(f"{seg}_{g}", f"   {g}", data[g], "entities", NUM, note, True, SRC[g])

# focus-6 subtotals (formulas)
secrow(A, row, "Focus-6 subtotals (calc)"); row += 1
for seg, lbl in [("AG","Agencies"),("TF","Eng-heavy firms"),("MS","Mid-size firms")]:
    cells = "+".join(ref[f"{seg}_{g}"] for g in GEOS)
    put(f"SUM_{seg}", f"   {lbl} - focus-6 total", f"=({cells})", "entities", NUM,
        "Sum across France, UK, UAE, Morocco, KSA, Egypt.", is_input=False)

# ---- Market breadth toggle (scales the addressable universe) ----
band(A, row, "1b. Market breadth toggle"); row += 1
put("BREADTH_MODE", "Active breadth mode (1 / 2 / 3)", 1, "mode", '0',
    "1 engineers only, 2 engineers + select white-collar, 3 all white-collar hiring.")
put("BREADTH_NAME", "   Active breadth",
    f'=CHOOSE({ref["BREADTH_MODE"]},"1 Engineers only","2 Some white-collar","3 All white-collar")',
    "", '@', "", is_input=False)
# 3x3 multiplier matrix (segment x mode)
A.cell(row,2,"Breadth multiplier by segment").font=f_h2
for j,t in [(3,"1 Eng"),(4,"2 Some WC"),(5,"3 All WC")]:
    A.cell(row,j,t).font=f_note; A.cell(row,j).alignment=center
A.cell(row,6,"Notes").font=f_note
row+=1
breadth_rows={}
for seg,lbl,m1,m2,m3,logic in [
  ("AG","Agencies",1,3,7,"Tech agencies are ~10-15% of all recruitment agencies."),
  ("TF","Eng-heavy firms",1,2,4,"Tech firms to all firms with structured hiring."),
  ("MS","Mid-size firms",1,1.8,3,"Engineer-hiring to all mid-size hirers."),
]:
    A.cell(row,2,f"   {lbl}").font=f_label
    for j,val in [(3,m1),(4,m2),(5,m3)]:
        c=A.cell(row,j,val); c.font=f_input; c.fill=fill_inp; c.number_format='0.0'; c.alignment=right
    A.cell(row,6,logic).font=f_note
    breadth_rows[seg]=row
    row+=1
secrow(A, row, "Effective addressable entities (calc = subtotal x active breadth mult)"); row += 1
for seg,lbl in [("AG","Agencies"),("TF","Eng-heavy firms"),("MS","Mid-size firms")]:
    r=breadth_rows[seg]
    mult=f"CHOOSE({ref['BREADTH_MODE']},$C${r},$D${r},$E${r})"
    put(f"EFF_{seg}", f"   {lbl} - effective", f"={ref['SUM_'+seg]}*{mult}", "entities", NUM,
        "Focus-6 subtotal x active breadth multiplier.", is_input=False)

# ---- Section 2: use-case fit ----
band(A, row, "2.  Share that fits the structured-evaluation use case"); row += 1
put("FIT_AG", "Agencies - % fit use case", 0.55, "%", PCT,
    "Assessment is core to agency workflow.")
put("FIT_TF", "Eng-heavy firms - % fit use case", 0.15, "%", PCT,
    "Base skews micro; only volume hirers fit.")
put("FIT_MS", "Mid-size firms - % fit use case", 0.25, "%", PCT,
    "Lower adoption; hire sporadically.")

# ---- Section 3: pricing & packs ----
band(A, row, "3.  Pricing  (usage-based, per test)"); row += 1
put("P_PAYG", "PAYG base price per test", 14.00, "$/test", USD2,
    "From BP. Premium vs TestGorilla (~$4/candidate); near TestDome/HackerRank.", src="S7")
put("D_S", "Small pack discount (20 tests)", 0.10, "%", PCT0, "")
put("D_M", "Medium pack discount (100 tests)", 0.15, "%", PCT0, "")
put("D_L", "Enterprise pack discount (1,000 tests)", 0.20, "%", PCT0,
    "1,000-test pack, $11,200 ($11.20/test).")
# net prices (calc)
put("NP_PAYG", "   Net price - PAYG", f"={ref['P_PAYG']}", "$/test", USD2, "", is_input=False)
put("NP_S", "   Net price - Small", f"={ref['P_PAYG']}*(1-{ref['D_S']})", "$/test", USD2, "", is_input=False)
put("NP_M", "   Net price - Medium", f"={ref['P_PAYG']}*(1-{ref['D_M']})", "$/test", USD2, "", is_input=False)
put("NP_L", "   Net price - Enterprise (1,000-pack)", f"={ref['P_PAYG']}*(1-{ref['D_L']})", "$/test", USD2, "", is_input=False)

# ---- Section 4: pack mix ----
band(A, row, "4.  Pack mix"); row += 1
put("MIX_PAYG", "PAYG share", 0.10, "%", PCT0, "Pack-led mix; PAYG is a thin trial tier.")
put("MIX_S", "Small pack share", 0.35, "%", PCT0, "")
put("MIX_M", "Medium pack share", 0.35, "%", PCT0, "")
put("MIX_L", "Enterprise (1,000-pack) share", 0.20, "%", PCT0, "")
put("MIX_CHK", "   Mix check (=100%)",
    f"={ref['MIX_PAYG']}+{ref['MIX_S']}+{ref['MIX_M']}+{ref['MIX_L']}",
    "%", PCT0, "Must equal 100%.", is_input=False)
# blended net price across the four per-test packs (all per-test; no flat enterprise fee)
put("BLEND", "   Blended net price per test",
    f"=({ref['MIX_PAYG']}*{ref['NP_PAYG']}+{ref['MIX_S']}*{ref['NP_S']}+{ref['MIX_M']}*{ref['NP_M']}+{ref['MIX_L']}*{ref['NP_L']})"
    f"/({ref['MIX_PAYG']}+{ref['MIX_S']}+{ref['MIX_M']}+{ref['MIX_L']})",
    "$/test", USD2, "Mix-weighted net price across all four packs.", is_input=False)

# ---- Section 5: tests-per-account engine ----
band(A, row, "5.  Tests per account / year  (roles x candidates per role)"); row += 1
put("TSCALE", "Tests-per-account scalar (sensitivity handle)", 1.00, "x", '0.00',
    "Multiplier on tests/account.")
for seg, lbl, roles, cand, note in [
    ("AG","Agencies", 80, 6, "High volume across client roles."),
    ("TF","Eng-heavy firms", 25, 6, "Steady in-house hiring."),
    ("MS","Mid-size firms", 10, 5, "Lower-volume hiring."),
]:
    put(f"ROLE_{seg}", f"   {lbl} - roles hired / yr", roles, "roles", NUM, note)
    put(f"CAND_{seg}", f"   {lbl} - candidates tested / role", cand, "cand", NUM, "")
    put(f"TPA_{seg}", f"   {lbl} - tests / account / yr",
        f"={ref[f'ROLE_{seg}']}*{ref[f'CAND_{seg}']}*{ref['TSCALE']}", "tests", NUM,
        "Roles x candidates x scalar.", is_input=False)

# channel-specific tests/account
put("ROLE_CH", "   Channel acct - roles hired / yr", 40, "roles", NUM,
    "Mix of agencies and in-house recruiters.")
put("CAND_CH", "   Channel acct - candidates tested / role", 5, "cand", NUM, "")
put("TPA_CH", "   Channel acct - tests / account / yr",
    f"={ref['ROLE_CH']}*{ref['CAND_CH']}*{ref['TSCALE']}", "tests", NUM, "", is_input=False)

# ---- Section 6: ACV per account (calc) ----
band(A, row, "6.  ACV per account  (revenue/yr)"); row += 1
for seg, lbl in [("AG","Agencies"),("TF","Eng-heavy firms"),("MS","Mid-size firms"),("CH","Channel acct")]:
    tpa = ref[f"TPA_{seg}"]
    put(f"ACV_{seg}", f"   {lbl} - ACV / account / yr",
        f"={tpa}*{ref['BLEND']}",
        "$/yr", USD, "tests/account x blended net price per test.", is_input=False)

# ---- Section 7: channel route (partner distribution portfolio) ----
band(A, row, "7.  Channel: partner distribution portfolio"); row += 1
A.cell(row,2,"Partner  (set On=1 to include)").font=f_h2
for j,t in [(3,"On 1/0"),(4,"Client base"),(5,"ICP-rel%"),(6,"Attach y3"),(7,"Accounts")]:
    A.cell(row,j,t).font=f_note; A.cell(row,j).alignment=center
row+=1
partner_rows=[]
PARTNERS=[
  ("Manatal (ATS, live)",        1,10000,0.45,0.08,"S8"),
  ("Socium (2nd ATS, exploring)",1, 3000,0.45,0.05,"S9"),
  ("Other HR-tech / ATS players",1,15000,0.35,0.03,""),
  ("Staffing-firm networks",     1, 5000,0.60,0.04,""),
  ("HR consultancies",           1, 3000,0.50,0.04,""),
  ("System integrators",         1, 2000,0.40,0.03,""),
]
for name,on,clients,rel,att,src in PARTNERS:
    A.cell(row,2,name).font=f_label
    for j,val,fmt in [(3,on,'0'),(4,clients,NUM),(5,rel,PCT0),(6,att,PCT)]:
        c=A.cell(row,j,val); c.font=f_input; c.fill=fill_inp; c.number_format=fmt
        c.alignment=center if j==3 else right
    ac=A.cell(row,7,f"=C{row}*D{row}*E{row}*F{row}"); ac.font=f_calc; ac.number_format=NUM
    if src: A.cell(row,5).value=A.cell(row,5).value  # keep
    A.cell(row,8,src).font=f_note; A.cell(row,8).alignment=center
    partner_rows.append(row)
    row+=1
A.cell(row,2,"Channel base accounts (portfolio, y3)").font=f_calcb
ct=A.cell(row,7,f"=SUM(G{partner_rows[0]}:G{partner_rows[-1]})")
ct.font=f_calcb; ct.number_format=NUM; ct.border=btop
ref["CH_BASE_ACCTS"]=f"Assumptions!$G${row}"
row+=1
put("ATTACH_SCALAR", "Channel attach scalar (sensitivity handle)", 1.00, "x", '0.00',
    "Multiplier on all partner attach rates.")

# ---- Section 8: direct route (self-serve + sales) ----
band(A, row, "8.  SOM - Direct route (self-serve + sales-led)"); row += 1
put("DIR_PEN_Y3", "3-yr penetration of serviceable SAM accounts", 0.008, "%", PCT,
    "Thin slice; gated by low trial-to-paid.")
put("TRAFFIC", "Monthly website traffic", 8000, "visits/mo", NUM,
    "Funnel cross-check (year-3 run-rate).")
put("TRIAL_CV", "Visit -> free-trial conversion", 0.03, "%", PCT, "")
put("T2P", "Trial -> paid conversion", 0.05, "%", PCT,
    "Low today; not best-in-class.", src="S8")

# ---- Section 9: retention & TAM scope ----
band(A, row, "9.  Retention & global scope"); row += 1
put("RETENTION", "Steady-state annual retention", 0.90, "%", PCT,
    "BP uses 7% churn; 90% here. 100% to date on a tiny base.", src="S8")
put("GLOB_SHARE", "Focus-6 share of global addressable universe", 0.12, "%", PCT,
    "Focus-6 as % of global; grosses up to full opportunity.")

# ---- Section 10: value metric ----
band(A, row, "10. Value delivered  (cost-of-bad-hire anchor)"); row += 1
put("BADHIRE", "Cost of a bad hire", 150000, "$", USD,
    "Value anchor.", src="S8")
put("TESTS_PER_HIRE", "Tests run per hire placed", 6, "tests", NUM,
    "Candidates per hire.")
put("MITIGATION", "Bad-hire rate reduction from structured eval", 0.20, "%", PCT0,
    "Share of bad-hire cost avoided.")

# ---- Section 11: talent-DB monetization (2nd revenue stream) ----
band(A, row, "11. Talent-DB monetization  (2nd stream)"); row += 1
put("PCT_FROM_DB", "Profiles sourced from own DB (mature)", 0.30, "%", PCT,
    "BP ramps 1% to 50%; 30% mature. Profiles = tests x this %.", src="S7")
put("PRICE_PROFILE", "Avg price per sourced profile", 30, "$", USD,
    "$30 per profile (BP).", src="S7")

ASSUM_LAST = row

# ============================================================================
# SHEET: Build  (TAM / SAM / SOM)
# ============================================================================
B = wb.create_sheet("Build")
B.sheet_view.showGridLines = False
for col, w in [("A",3),("B",40),("C",14),("D",14),("E",14),("F",14),("G",14),("H",16)]:
    B.column_dimensions[col].width = w
br = 1

def bband(r, text, c1=2, c2=8):
    B.cell(r, c1, text).font = f_h1
    for c in range(c1, c2+1):
        B.cell(r, c).fill = fill_h1
    B.row_dimensions[r].height = 18

def bsec(r, text, c1=2, c2=8):
    B.cell(r, c1, text).font = f_h2
    for c in range(c1, c2+1):
        B.cell(r, c).fill = fill_sec

B.cell(br,2,"INVIRTUS  -  Revenue potential build  (TAM / SAM / SOM)").font = f_title; br += 1
B.cell(br,2,"Bottoms-up: entities x %fit x tests/yr x net $/test, plus talent-DB stream. USD. All drivers live on Assumptions.").font = f_sub
B.merge_cells(start_row=br,start_column=2,end_row=br,end_column=8); br += 2

# ---------- Serviceable accounts by segment (focus-6) ----------
bband(br, "A.  Serviceable accounts  (focus-6 = SAM geography)"); br += 1
hdr = ["Segment","Entities (focus-6)","% fit","Serviceable accts","Tests/acct/yr","ACV/acct ($)","SAM value ($/yr)"]
for i,h in enumerate(hdr):
    cell=B.cell(br,2+i,h); cell.font=f_h2; cell.fill=fill_sec; cell.alignment=center
br += 1
seg_rows = {}
for seg,lbl in [("AG","Tech recruiting agencies"),("TF","Engineering-heavy firms"),("MS","Mid-size structured-hiring firms")]:
    B.cell(br,2,lbl).font=f_label
    B.cell(br,3,f"={ref['EFF_'+seg]}").number_format=NUM
    B.cell(br,4,f"={ref['FIT_'+seg]}").number_format=PCT
    B.cell(br,5,f"=C{br}*D{br}").number_format=NUM
    B.cell(br,6,f"={ref['TPA_'+seg]}").number_format=NUM
    B.cell(br,7,f"={ref['ACV_'+seg]}").number_format=USD
    B.cell(br,8,f"=E{br}*G{br}").number_format=USD
    for c in range(3,9): B.cell(br,c).font=f_calc
    seg_rows[seg]=br
    br += 1
# SAM total row
B.cell(br,2,"SAM  (focus-6, serviceable today)").font=f_calcb
B.cell(br,5,f"=SUM(E{seg_rows['AG']}:E{seg_rows['MS']})").number_format=NUM
B.cell(br,8,f"=SUM(H{seg_rows['AG']}:H{seg_rows['MS']})").number_format=USD
for c in [5,8]:
    B.cell(br,c).font=f_calcb; B.cell(br,c).border=btop
SAM_ROW=br; SAM_VAL=f"Build!$H${br}"; SAM_ACCTS=f"Build!$E${br}"
br += 2

# ---------- TAM ----------
bband(br, "B.  TAM  (all geos where product works, at full ACV)"); br += 1
B.cell(br,2,"Global gross-up = focus-6 / (focus-6 share of global)").font=f_note; br+=1
for i,h in enumerate(["Segment","Global entities","% fit","Serviceable accts","ACV/acct ($)","TAM value ($/yr)"]):
    cell=B.cell(br,2+i,h); cell.font=f_h2; cell.fill=fill_sec; cell.alignment=center
br+=1
tam_rows={}
for seg,lbl in [("AG","Tech recruiting agencies"),("TF","Engineering-heavy firms"),("MS","Mid-size structured-hiring firms")]:
    B.cell(br,2,lbl).font=f_label
    B.cell(br,3,f"={ref['EFF_'+seg]}/{ref['GLOB_SHARE']}").number_format=NUM
    B.cell(br,4,f"={ref['FIT_'+seg]}").number_format=PCT
    B.cell(br,5,f"=C{br}*D{br}").number_format=NUM
    B.cell(br,6,f"={ref['ACV_'+seg]}").number_format=USD
    B.cell(br,7,f"=E{br}*F{br}").number_format=USD
    for c in range(3,8): B.cell(br,c).font=f_calc
    tam_rows[seg]=br; br+=1
B.cell(br,2,"TAM  (global prize, full ACV)").font=f_calcb
B.cell(br,7,f"=SUM(G{tam_rows['AG']}:G{tam_rows['MS']})").number_format=USD
B.cell(br,7).font=f_calcb; B.cell(br,7).border=btop
TAM_VAL=f"Build!$G${br}"
br+=2

# ---------- SOM channel + direct ----------
bband(br, "C.  SOM  (3-year reachable)  -  two routes side by side"); br += 1
c_start=br
B.cell(br,2,"CHANNEL ROUTE  (partner portfolio - see Assumptions sec.7)").font=f_h2
B.cell(br,2).fill=fill_sec
for c in [3,4,5]: B.cell(br,c).fill=fill_sec
br+=1
B.cell(br,2,"Channel base accounts (portfolio, y3)").font=f_label
B.cell(br,5,f"={ref['CH_BASE_ACCTS']}").number_format=NUM; B.cell(br,5).font=f_calc
CH_BASE=br; br+=1
B.cell(br,2,"Attach scalar").font=f_label
B.cell(br,5,f"={ref['ATTACH_SCALAR']}").number_format='0.00'; B.cell(br,5).font=f_calc
CH_SC=br; br+=1
B.cell(br,2,"Active channel accounts").font=f_label
B.cell(br,5,f"=E{CH_BASE}*E{CH_SC}").number_format=NUM; B.cell(br,5).font=f_calc
CH_ACCT=br; br+=1
B.cell(br,2,"Channel revenue ($/yr, retention-adj)").font=f_calcb
B.cell(br,5,f"=E{CH_ACCT}*{ref['ACV_CH']}*{ref['RETENTION']}").number_format=USD
B.cell(br,5).font=f_calcb
CH_REV=br; br+=2

# Direct route block
B.cell(br,2,"DIRECT ROUTE  (self-serve + sales-led)").font=f_h2; B.cell(br,2).fill=fill_sec
B.cell(br,5,"Year 3").font=f_h2; B.cell(br,5).fill=fill_sec; B.cell(br,5).alignment=center
for c in [3,4]: B.cell(br,c).fill=fill_sec
br+=1
B.cell(br,2,"Serviceable SAM accounts").font=f_label
B.cell(br,5,f"={SAM_ACCTS}").number_format=NUM; B.cell(br,5).font=f_calc
D_SAM=br; br+=1
B.cell(br,2,"3-yr penetration").font=f_label
B.cell(br,5,f"={ref['DIR_PEN_Y3']}").number_format=PCT; B.cell(br,5).font=f_calc
D_PEN=br; br+=1
B.cell(br,2,"Active direct accounts").font=f_label
B.cell(br,5,f"=E{D_SAM}*E{D_PEN}").number_format=NUM; B.cell(br,5).font=f_calc
D_ACCT=br; br+=1
B.cell(br,2,"Direct revenue ($/yr, retention-adj)").font=f_calcb
# Penetrating x% of the serviceable market captures x% of SAM value (segment-weighted ACV).
B.cell(br,5,f"={SAM_VAL}*E{D_PEN}*{ref['RETENTION']}").number_format=USD
B.cell(br,5).font=f_calcb
D_REV=br; br+=1
B.cell(br,2,"Funnel cross-check: implied paid/yr").font=f_note
B.cell(br,5,f"={ref['TRAFFIC']}*12*{ref['TRIAL_CV']}*{ref['T2P']}").number_format=NUM
B.cell(br,5).font=f_note
br+=2

# SOM total
B.cell(br,2,"SOM  (Year 3 = Channel + Direct)").font=f_calcb
B.cell(br,2).fill=fill_sec
B.cell(br,5,f"=E{CH_REV}+E{D_REV}").number_format=USD
B.cell(br,5).font=Font(name="Calibri",size=12,bold=True,color="FF1F2A44")
B.cell(br,5).fill=fill_sec
SOM_VAL=f"Build!$E${br}"
SOM_CH=f"Build!$E${CH_REV}"; SOM_DIR=f"Build!$E${D_REV}"
SOM_ACCT_TOT=f"(Build!$E${CH_ACCT}+Build!$E${D_ACCT})"
br+=2

# ---------- two revenue streams & total revenue potential ----------
bband(br, "D.  Revenue potential by stream  (assessments + talent-DB)"); br+=1
for i,h in enumerate(["Tier","Assessment rev ($/yr)","Tests run / yr","Talent-DB rev ($/yr)","TOTAL revenue potential ($/yr)"]):
    cell=B.cell(br,2+i,h); cell.font=f_h2; cell.fill=fill_sec; cell.alignment=center
br+=1
tot_cells={}
for tier,assess_ref in [("Full opportunity (TAM)",TAM_VAL),("Serviceable today (SAM)",SAM_VAL),("3-yr reachable (SOM)",SOM_VAL)]:
    B.cell(br,2,tier).font=f_label
    B.cell(br,3,f"={assess_ref}").number_format=USD
    B.cell(br,4,f"=C{br}/{ref['BLEND']}").number_format=NUM                       # tests = assess rev / blended price
    B.cell(br,5,f"=D{br}*{ref['PCT_FROM_DB']}*{ref['PRICE_PROFILE']}").number_format=USD  # DB = tests x %fromDB x price
    B.cell(br,6,f"=C{br}+E{br}").number_format=USD
    for c in range(3,7): B.cell(br,c).font=f_calc
    B.cell(br,6).font=f_calcb
    tot_cells[tier]=br
    br+=1
TOT_TAM=f"Build!$F${tot_cells['Full opportunity (TAM)']}"
TOT_SAM=f"Build!$F${tot_cells['Serviceable today (SAM)']}"
TOT_SOM=f"Build!$F${tot_cells['3-yr reachable (SOM)']}"
ASSESS_SOM=f"Build!$C${tot_cells['3-yr reachable (SOM)']}"
DB_SOM=f"Build!$E${tot_cells['3-yr reachable (SOM)']}"
br+=1

# ---------- value-delivered parallel read ----------
bband(br, "E.  Value delivered  (tests run -> hires -> bad-hire cost addressed)"); br+=1
for i,h in enumerate(["Ring","Accounts","Tests run / yr","Hires supported / yr","Bad-hire $ exposure influenced","Mitigable value to buyers ($)"]):
    cell=B.cell(br,2+i,h); cell.font=f_h2; cell.fill=fill_sec; cell.alignment=center
br+=1
# SOM value row
B.cell(br,2,"SOM (yr 3)").font=f_label
B.cell(br,3,f"={SOM_ACCT_TOT}").number_format=NUM
B.cell(br,4,f"=C{br}*{ref['TPA_CH']}").number_format=NUM
B.cell(br,5,f"=D{br}/{ref['TESTS_PER_HIRE']}").number_format=NUM
B.cell(br,6,f"=E{br}*{ref['BADHIRE']}").number_format=USD
B.cell(br,7,f"=F{br}*{ref['MITIGATION']}").number_format=USD
for c in range(3,8): B.cell(br,c).font=f_calc
br+=2

B.cell(br,2,"Notes: Agencies and the firms they serve are counted separately, so there is no double count. Channel and direct may overlap a little; treated as additive. The value-delivered figures are buyer-side, not market size.").font=f_note
B.merge_cells(start_row=br,start_column=2,end_row=br+1,end_column=8)
B.cell(br,2).alignment=left

# ============================================================================
# SHEET: Cover
# ============================================================================
C = wb.create_sheet("Cover")
wb.move_sheet("Cover", -(wb.sheetnames.index("Cover")))  # move to front
C.sheet_view.showGridLines = False
for col,w in [("A",3),("B",30),("C",22),("D",22),("E",22),("F",6)]:
    C.column_dimensions[col].width=w
cr=2
C.cell(cr,2,"INVIRTUS").font=Font(name="Calibri",size=30,bold=True,color="FF1F2A44"); cr+=1
C.cell(cr,2,"Bottoms-Up Revenue Potential  (two streams; TAM / SAM / SOM equivalents)").font=Font(name="Calibri",size=14,color=GREY); cr+=1
C.cell(cr,2,"Candidate assessment platform  |  USD  |  June 2026").font=f_note; cr+=2

# one-line answer band
C.cell(cr,2,"THE ANSWER").font=f_h1
for c in range(2,6): C.cell(cr,c).fill=fill_h1
cr+=1
C.merge_cells(start_row=cr,start_column=2,end_row=cr+2,end_column=5)
C.cell(cr,2,(
 "The 3-year reachable number runs on two streams: assessment packs and talent-DB. "
 "The assessment line is channel-led (partner portfolio ~76%, direct ~24%). The biggest "
 "swing is the channel attach scalar, then tests/account and price (see Sensitivity).")).font=Font(name="Calibri",size=11,color="FF1F2A44")
C.cell(cr,2).alignment=left
C.cell(cr,2).fill=fill_cov
for r in range(cr,cr+3):
    for c in range(2,6): C.cell(r,c).fill=fill_cov
cr+=4

# the three cards = TOTAL revenue potential (assessments + talent-DB)
C.cell(cr,2,"FULL OPPORTUNITY").font=f_biglbl
C.cell(cr,3,"SERVICEABLE TODAY").font=f_biglbl
C.cell(cr,4,"3-YEAR REACHABLE").font=f_biglbl
cr+=1
C.cell(cr,2,f"={TOT_TAM}").number_format=USDB; C.cell(cr,2).font=f_big
C.cell(cr,3,f"={TOT_SAM}").number_format=USDM; C.cell(cr,3).font=f_big
C.cell(cr,4,f"={TOT_SOM}").number_format=USDM; C.cell(cr,4).font=f_big
cr+=1
C.cell(cr,2,"(TAM) all geos | 2 streams").font=f_note
C.cell(cr,3,"(SAM) focus-6 | serviceable").font=f_note
C.cell(cr,4,"(SOM) channel+direct | 3-yr").font=f_note
cr+=2

# 3-yr reachable by stream
C.cell(cr,2,"3-yr reachable - by revenue stream").font=f_h2; cr+=1
C.cell(cr,2,"Assessment packs").font=f_label
C.cell(cr,3,f"={ASSESS_SOM}").number_format=USDM; C.cell(cr,3).font=f_calcb; cr+=1
C.cell(cr,2,"Talent-DB monetization").font=f_label
C.cell(cr,3,f"={DB_SOM}").number_format=USDM; C.cell(cr,3).font=f_calcb; cr+=1
C.cell(cr,2,"Total revenue potential").font=f_calcb
C.cell(cr,3,f"={TOT_SOM}").number_format=USDM; C.cell(cr,3).font=f_calcb; cr+=2

# 3-yr reachable assessments by route
C.cell(cr,2,"3-yr reachable - assessments by route").font=f_h2; cr+=1
C.cell(cr,2,"Channel (partner portfolio)").font=f_label
C.cell(cr,3,f"={SOM_CH}").number_format=USDM; C.cell(cr,3).font=f_calcb; cr+=1
C.cell(cr,2,"Direct (self-serve + sales)").font=f_label
C.cell(cr,3,f"={SOM_DIR}").number_format=USDM; C.cell(cr,3).font=f_calcb; cr+=2

# scenario levers (live indicators)
C.cell(cr,2,"Scenario levers (edit on Assumptions)").font=f_h2; cr+=1
C.cell(cr,2,"Market breadth mode").font=f_label
C.cell(cr,3,f"={ref['BREADTH_NAME']}").font=f_calcb; cr+=1
C.cell(cr,2,"Channel attach scalar").font=f_label
C.cell(cr,3,f"={ref['ATTACH_SCALAR']}").number_format='0.00'; C.cell(cr,3).font=f_calcb; cr+=2

# read-me / legend
C.cell(cr,2,"How to read this model").font=f_h2; cr+=1
for line in [
 "1.  Blue cells are inputs; black cells are formulas.",
 "2.  Breadth toggle (1 engineers / 2 some white-collar / 3 all white-collar) scales the universe.",
 "3.  Channel is a partner portfolio (Assumptions sec.7); each partner can be switched on or off.",
 "4.  Two streams: assessment packs + talent-DB. Sources lists the researched inputs.",
]:
    C.cell(cr,2,line).font=f_calc
    C.merge_cells(start_row=cr,start_column=2,end_row=cr,end_column=5)
    C.cell(cr,2).alignment=left
    cr+=1
cr+=1
# legend
C.cell(cr,2,"Input (editable)").font=f_input; C.cell(cr,2).fill=fill_inp
C.cell(cr,3,"Formula (calc)").font=f_calc
cr+=1

# ============================================================================
# SHEET: Sensitivity
# ============================================================================
S = wb.create_sheet("Sensitivity")
S.sheet_view.showGridLines=False
for col,w in [("A",3),("B",30),("C",14),("D",14),("E",14),("F",14),("G",14)]:
    S.column_dimensions[col].width=w
sr=2
S.cell(sr,2,"INVIRTUS  -  Sensitivity").font=f_title; sr+=1
S.cell(sr,2,"SOM (Yr3) response to the three swing inputs. Tables recompute the SOM chain live from the axis values; non-swung drivers pull from Assumptions.").font=f_sub
S.merge_cells(start_row=sr,start_column=2,end_row=sr,end_column=7); sr+=2

# --- helper: build a closed-form SOM (yr3) formula given symbolic tests-per-acct & payg price & attach ---
# blended price as fn of payg p (packs/mix fixed):
def blended_of(p):
    return (f"(({ref['MIX_PAYG']}*{p}+{ref['MIX_S']}*{p}*(1-{ref['D_S']})+"
            f"{ref['MIX_M']}*{p}*(1-{ref['D_M']})+{ref['MIX_L']}*{p}*(1-{ref['D_L']}))"
            f"/({ref['MIX_PAYG']}+{ref['MIX_S']}+{ref['MIX_M']}+{ref['MIX_L']}))")

def acv_of(tpa_expr, p):
    # ACV per acct: tests x blended net price (all per-test packs; no flat fee)
    return f"({tpa_expr}*{blended_of(p)})"

def som_formula(p_cell, tpa_ch_cell, scalar_cell):
    """SOM yr3 = channel + direct, parameterised by payg price, channel tests/acct, attach scalar."""
    chan=f"{ref['CH_BASE_ACCTS']}*{scalar_cell}*{acv_of(tpa_ch_cell,p_cell)}*{ref['RETENTION']}"
    # direct: penetrate x% of the serviceable market => x% of SAM value (segment-weighted ACV under price p)
    parts=[]
    for s in ["AG","TF","MS"]:
        tpa_s = "(" + ref['ROLE_'+s] + "*" + ref['CAND_'+s] + "*" + ref['TSCALE'] + ")"
        parts.append("(" + ref['EFF_'+s] + "*" + ref['FIT_'+s] + "*" + acv_of(tpa_s, p_cell) + ")")
    sam_p="+".join(parts)
    direct=f"(({sam_p})*{ref['DIR_PEN_Y3']}*{ref['RETENTION']})"
    return f"={chan}+{direct}"

# ---- Table 1: Channel tests/acct (rows) x PAYG price (cols) ----
S.cell(sr,2,"Table 1:  SOM (Yr3)  -  Channel tests/account  vs  PAYG price/test").font=f_h2; sr+=1
price_axis=[10,12.5,15,17.5,20]
tpa_axis=[100,150,200,250,300]
hdr_r=sr
S.cell(sr,2,"tests/acct  \\  $/test").font=f_note
for j,p in enumerate(price_axis):
    cell=S.cell(sr,3+j,p); cell.font=f_input; cell.fill=fill_inp; cell.number_format=USD2; cell.alignment=center
sr+=1
t1_top=sr
for i,t in enumerate(tpa_axis):
    S.cell(sr,2,t).font=f_input; S.cell(sr,2).fill=fill_inp; S.cell(sr,2).number_format=NUM
    for j,p in enumerate(price_axis):
        p_cell=f"{get_column_letter(3+j)}${hdr_r}"
        t_cell=f"$B{sr}"
        S.cell(sr,3+j, som_formula(p_cell, t_cell, ref['ATTACH_SCALAR'])).number_format=USDM
        S.cell(sr,3+j).font=f_calc
    sr+=1
sr+=1

# ---- Table 2: Attach rate (rows) x Channel tests/acct (cols) ----
S.cell(sr,2,"Table 2:  SOM (Yr3)  -  Channel attach scalar  vs  Channel tests/account").font=f_h2; sr+=1
att_axis=[0.5,0.75,1.0,1.25,1.5]
hdr2=sr
S.cell(sr,2,"attach x  \\  tests/acct").font=f_note
for j,t in enumerate(tpa_axis):
    cell=S.cell(sr,3+j,t); cell.font=f_input; cell.fill=fill_inp; cell.number_format=NUM; cell.alignment=center
sr+=1
for i,a in enumerate(att_axis):
    S.cell(sr,2,a).font=f_input; S.cell(sr,2).fill=fill_inp; S.cell(sr,2).number_format='0.00'
    for j,t in enumerate(tpa_axis):
        t_cell=f"{get_column_letter(3+j)}${hdr2}"
        a_cell=f"$B{sr}"
        S.cell(sr,3+j, som_formula(ref['P_PAYG'], t_cell, a_cell)).number_format=USDM
        S.cell(sr,3+j).font=f_calc
    sr+=1
sr+=2

# ---- Tornado data (low/base/high) ----
S.cell(sr,2,"Tornado:  SOM swing from +/- moves in each driver").font=f_h2; sr+=1
for i,h in enumerate(["Driver","Low case","Base","High case"]):
    cell=S.cell(sr,2+i,h); cell.font=f_h2; cell.fill=fill_sec; cell.alignment=center
sr+=1
# base attach/tests/price
base_p=ref['P_PAYG']; base_att=ref['ATTACH_SCALAR']
# Channel tests base = TPA_CH
base_tpa=ref['TPA_CH']
tornado=[
 ("Channel attach scalar (0.5 / 1.0 / 1.5)",
   som_formula(base_p, base_tpa, "0.5"), som_formula(base_p, base_tpa, base_att), som_formula(base_p, base_tpa, "1.5")),
 ("Channel tests/account (100 / base / 300)",
   som_formula(base_p, "100", base_att), som_formula(base_p, base_tpa, base_att), som_formula(base_p, "300", base_att)),
 ("PAYG price/test ($10 / base $14 / $20)",
   som_formula("10", base_tpa, base_att), som_formula(base_p, base_tpa, base_att), som_formula("20", base_tpa, base_att)),
]
for lbl,lo,ba,hi in tornado:
    S.cell(sr,2,lbl).font=f_label
    S.cell(sr,3,lo).number_format=USDM; S.cell(sr,3).font=f_calc
    S.cell(sr,4,ba).number_format=USDM; S.cell(sr,4).font=f_calc
    S.cell(sr,5,hi).number_format=USDM; S.cell(sr,5).font=f_calc
    sr+=1

# ============================================================================
# SHEET: Sources
# ============================================================================
SO=wb.create_sheet("Sources")
SO.sheet_view.showGridLines=False
for col,w in [("A",3),("B",6),("C",34),("D",30),("E",58),("F",12),("G",10)]:
    SO.column_dimensions[col].width=w
sor=2
SO.cell(sor,2,"INVIRTUS  -  Sources").font=f_title; sor+=1
SO.cell(sor,2,"One row per researched input. Figures triangulated where official counts are absent; confidence noted.").font=f_sub
SO.merge_cells(start_row=sor,start_column=2,end_row=sor,end_column=7); sor+=2
for i,h in enumerate(["Key","Input","Source","URL","Date","Conf."]):
    cell=SO.cell(sor,2+i,h); cell.font=f_h2; cell.fill=fill_sec; cell.alignment=center
sor+=1

SOURCES=[
 ("S1","France: agencies & firm counts","INSEE SIRENE (7810Z/7820Z); Modeles de Business Plan",
   "public.opendatasoft.com/explore/dataset/economicref-france-sirene-v3 ; modelesdebusinessplan.com/blogs/infos/nombre-cabinets-recrutement-france","2026-06-16","High/Low"),
 ("S2","UK: agencies & firm counts","REC / ONS / Dealroom (triangulated)",
   "rec.uk.com ; ons.gov.uk","2026-06-16","Med"),
 ("S3","UAE: tech firms & mid-size","DIC; MAGNiTT; DIFC 2024; UAE Min. of Economy (557k SMEs)",
   "magnitt.com/en-ae/startups ; thenationalnews.com/business/economy/2023/04/06","2026-06-16","Med"),
 ("S4","Morocco: agencies, tech & mid-size","EcoActu/MIEPEEC; OMTPME 2025; UM6P; APEBI",
   "ecoactu.ma/lintermediation-emploi ; lematin.ma (OMTPME 2025)","2026-06-16","Med/Low"),
 ("S5","KSA: tech firms & mid-size","Monsha'at SME Monitor Q4'23 (18,723 medium); CST Manassa (1,000+)",
   "argaam.com/en/article/articledetail/id/1709526 ; my.gov.sa/en/news/10102","2026-06-16","Med/High"),
 ("S6","Egypt: tech firms & mid-size","CAPMAS 5th Economic Census; ITIDA Outlook 2026; MAGNiTT",
   "censusinfo.capmas.gov.eg ; itida.gov.eg ; magnitt.com/en-eg/startups","2026-06-16","Med/Low"),
 ("S7","Pricing: $14/test + packs (finalized BP); benchmarked","Invirtus finalized BP (Aida); TestGorilla; TestDome; Adaface; HackerRank pricing",
   "Invirtus_business_plan.xlsx ; testgorilla.com/pricing ; testdome.com/pricing ; adaface.com/pricing","2026-06-17","High"),
 ("S8","Traction: Manatal, retention, value anchor, funnel","Invirtus April 2026 IC Update; GTM notes (34V Notion)",
   "Notion: Invirtus - April 2026 IC Update ; Invirtus - GTM","2026-06-16","Co. data"),
 ("S9","Second ATS (Socium) - upside","Invirtus GTM notes (integration being explored)",
   "Notion: Invirtus - GTM","2026-06-16","Co. data"),
]
for k,inp,src,url,dt,conf in SOURCES:
    SO.cell(sor,2,k).font=f_calcb
    SO.cell(sor,3,inp).font=f_calc; SO.cell(sor,3).alignment=left
    SO.cell(sor,4,src).font=f_calc; SO.cell(sor,4).alignment=left
    SO.cell(sor,5,url).font=f_note; SO.cell(sor,5).alignment=left
    SO.cell(sor,6,dt).font=f_calc; SO.cell(sor,6).alignment=center
    SO.cell(sor,7,conf).font=f_calc; SO.cell(sor,7).alignment=center
    for c in range(2,8): SO.cell(sor,c).border=border
    sor+=1
sor+=1
SO.cell(sor,2,"Methodology note: 'medium enterprise' bands differ by country (KSA 50-249 emp; Morocco turnover-based 50-175M MAD; UAE <200 emp; Egypt 50+ emp census band). Mid-size counts are the closest official proxy x an engineer-hiring share. The engineer-hiring share and use-case-fit % are the softest inputs - pressure-test before locking TAM.").font=f_note
SO.merge_cells(start_row=sor,start_column=2,end_row=sor+2,end_column=7)
SO.cell(sor,2).alignment=left

# order sheets: Cover, Assumptions, Build, Sensitivity, Sources
wb.move_sheet("Cover", -wb.sheetnames.index("Cover"))
desired=["Cover","Assumptions","Build","Sensitivity","Sources"]
wb._sheets.sort(key=lambda s: desired.index(s.title))

wb.save("Invirtus_Market_Sizing.xlsx")
print("Saved Invirtus_Market_Sizing.xlsx")
print("Assumptions last row:", ASSUM_LAST)

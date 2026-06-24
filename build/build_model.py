#!/usr/bin/env python3
"""
Invirtus - Bottoms-Up Revenue Potential builder (lean).

Spine: buyers x % addressable x ACV per account = revenue, then three tiers
(Full opportunity / Serviceable today / 3-year reachable) and two streams
(assessment packs + talent-DB).

Conventions: blue = editable input, black = formula (no hardcodes in formulas).
Run:  python3 build/build_model.py   ->  Invirtus_Market_Sizing.xlsx
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---- styles ----
BLUE="FF0000FF"; BLACK="FF000000"; WHITE="FFFFFFFF"; NAVY="FF1F2A44"
LILAC="FFE7DDEA"; CREAM="FFF3EFE9"; GREY="FF7F7F7F"
f_title=Font(name="Calibri",size=22,bold=True,color=NAVY)
f_sub  =Font(name="Calibri",size=11,italic=True,color=GREY)
f_h1   =Font(name="Calibri",size=13,bold=True,color=WHITE)
f_h2   =Font(name="Calibri",size=11,bold=True,color=NAVY)
f_input=Font(name="Calibri",size=10,color=BLUE)
f_calc =Font(name="Calibri",size=10,color=BLACK)
f_calcb=Font(name="Calibri",size=10,bold=True,color=BLACK)
f_label=Font(name="Calibri",size=10,color=BLACK)
f_note =Font(name="Calibri",size=9,italic=True,color=GREY)
f_big  =Font(name="Calibri",size=28,bold=True,color=NAVY)
f_biglbl=Font(name="Calibri",size=11,bold=True,color=GREY)
fill_h1=PatternFill("solid",fgColor=NAVY)
fill_sec=PatternFill("solid",fgColor=LILAC)
fill_cov=PatternFill("solid",fgColor=CREAM)
fill_inp=PatternFill("solid",fgColor="FFFFFDE7")
btop=Border(top=Side(style="thin",color=GREY))
border=Border(left=Side(style="thin",color="FFD9D2DE"),right=Side(style="thin",color="FFD9D2DE"),
              top=Side(style="thin",color="FFD9D2DE"),bottom=Side(style="thin",color="FFD9D2DE"))
center=Alignment(horizontal="center",vertical="center")
left=Alignment(horizontal="left",vertical="center",wrap_text=True)
right=Alignment(horizontal="right",vertical="center")
PCT='0.0%'; PCT0='0%'; USD='#,##0'; USD2='$#,##0.00'; USDM='$#,##0.0,,"M"'; USDB='$#,##0.00,,,"B"'; NUM='#,##0'

# ---- researched data ----
GEOS=["France","UK","UAE","Morocco","KSA","Egypt"]
AGENCIES={"France":500,"UK":3500,"UAE":100,"Morocco":45,"KSA":50,"Egypt":60}
TECHFIRMS={"France":95000,"UK":107000,"UAE":5000,"Morocco":2000,"KSA":1800,"Egypt":2500}
MIDSIZE={"France":7300,"UK":13200,"UAE":2800,"Morocco":1250,"KSA":2600,"Egypt":2000}
SRC={"France":"S1","UK":"S2","UAE":"S3","Morocco":"S4","KSA":"S5","Egypt":"S6"}
PARTNERS=[
 ("Manatal (ATS, live)",        1,10000,0.45,0.08,"S8"),
 ("Socium (2nd ATS, exploring)",1, 3000,0.45,0.05,"S9"),
 ("Other HR-tech / ATS players",1,15000,0.35,0.03,""),
 ("Staffing-firm networks",     1, 5000,0.60,0.04,""),
 ("HR consultancies",           1, 3000,0.50,0.04,""),
 ("System integrators",         1, 2000,0.40,0.03,""),
]

wb=openpyxl.Workbook()
ref={}

# ============================================================ ASSUMPTIONS
A=wb.active; A.title="Assumptions"; A.sheet_view.showGridLines=False
for col,w in [("A",3),("B",44),("C",13),("D",11),("E",9),("F",50),("G",12),("H",7)]:
    A.column_dimensions[col].width=w
row=1
def band(t):
    global row
    A.cell(row,2,t).font=f_h1
    for c in range(2,7): A.cell(row,c).fill=fill_h1
    A.row_dimensions[row].height=18; row+=1
def secrow(t):
    global row
    A.cell(row,2,t).font=f_h2
    for c in range(2,7): A.cell(row,c).fill=fill_sec
    row+=1
def put(name,label,value,unit,fmt,note="",is_input=True,src=""):
    global row
    A.cell(row,2,label).font=f_label; A.cell(row,2).alignment=left
    c=A.cell(row,3,value); c.number_format=fmt; c.alignment=right
    if is_input: c.font=f_input; c.fill=fill_inp
    else: c.font=f_calc
    A.cell(row,4,unit).font=f_note; A.cell(row,4).alignment=left
    A.cell(row,5,src).font=f_note; A.cell(row,5).alignment=center
    A.cell(row,6,note).font=f_note; A.cell(row,6).alignment=left
    if name: ref[name]=f"Assumptions!$C${row}"
    row+=1

A.cell(row,2,"INVIRTUS  -  Assumptions").font=f_title; row+=1
A.cell(row,2,"Blue = editable input. Black = formula. Col E = source key (Sources tab). USD.").font=f_sub
A.merge_cells(start_row=row,start_column=2,end_row=row,end_column=6); row+=2
for c,t in [(2,"Driver"),(3,"Value"),(4,"Unit"),(5,"Src"),(6,"Note")]:
    A.cell(row,c,t).font=f_h2; A.cell(row,c).fill=fill_sec
row+=1

# 1. buyers
band("1.  Addressable buyers  (researched)")
for seg,label,data in [("AG","Tech recruiting agencies",AGENCIES),
                       ("TF","Engineering-heavy firms",TECHFIRMS),
                       ("MS","Mid-size firms (50-250)",MIDSIZE)]:
    secrow(label)
    for g in GEOS: put(f"{seg}_{g}",f"   {g}",data[g],"firms",NUM,"",True,SRC[g])
secrow("Focus-6 subtotal (calc)")
for seg,lbl in [("AG","Agencies"),("TF","Eng-heavy firms"),("MS","Mid-size firms")]:
    cells="+".join(ref[f"{seg}_{g}"] for g in GEOS)
    put(f"SUM_{seg}",f"   {lbl}",f"=({cells})","firms",NUM,"",is_input=False)

# 2. breadth toggle
band("2.  Market breadth toggle")
put("BREADTH_MODE","Active mode (1 / 2 / 3)",1,"mode",'0',
    "1 engineers only, 2 some white-collar, 3 all white-collar.")
put("BREADTH_NAME","   Active breadth",
    f'=CHOOSE({ref["BREADTH_MODE"]},"1 Engineers only","2 Some white-collar","3 All white-collar")',
    "",'@',"",is_input=False)
A.cell(row,2,"Multiplier by segment").font=f_h2
for j,t in [(3,"1 Eng"),(4,"2 Some WC"),(5,"3 All WC")]: A.cell(row,j,t).font=f_note; A.cell(row,j).alignment=center
A.cell(row,6,"Notes").font=f_note
row+=1
brow={}
for seg,lbl,m1,m2,m3,logic in [
  ("AG","Agencies",1,3,7,"Tech agencies are ~10-15% of all recruitment agencies."),
  ("TF","Eng-heavy firms",1,2,4,"Tech firms to all firms with structured hiring."),
  ("MS","Mid-size firms",1,1.8,3,"Engineer-hiring to all mid-size hirers."),
]:
    A.cell(row,2,f"   {lbl}").font=f_label
    for j,v in [(3,m1),(4,m2),(5,m3)]:
        c=A.cell(row,j,v); c.font=f_input; c.fill=fill_inp; c.number_format='0.0'; c.alignment=right
    A.cell(row,6,logic).font=f_note
    brow[seg]=row; row+=1
secrow("Effective buyers (calc = subtotal x active multiplier)")
for seg,lbl in [("AG","Agencies"),("TF","Eng-heavy firms"),("MS","Mid-size firms")]:
    r=brow[seg]
    put(f"EFF_{seg}",f"   {lbl}",
        f"={ref['SUM_'+seg]}*CHOOSE({ref['BREADTH_MODE']},$C${r},$D${r},$E${r})","firms",NUM,"",is_input=False)

# 3. value per account
band("3.  Value per account")
put("PRICE","Net price per test (blended)",12.22,"$/test",USD2,
    "$14 PAYG less 10-20% pack discounts at the BP pack mix.",src="S7")
put("FIT_AG","Agencies - % addressable",0.55,"%",PCT,"Core to agency workflow.")
put("FIT_TF","Eng-heavy firms - % addressable",0.15,"%",PCT,"Base skews micro; only volume hirers fit.")
put("FIT_MS","Mid-size firms - % addressable",0.25,"%",PCT,"Lower adoption; hire sporadically.")
for seg,lbl,t,note in [("AG","Agencies",480,"~80 roles x 6 candidates."),
                       ("TF","Eng-heavy firms",150,"~25 roles x 6 candidates."),
                       ("MS","Mid-size firms",50,"~10 roles x 5 candidates."),
                       ("CH","Channel account",200,"~40 roles x 5 candidates.")]:
    put(f"TPA_{seg}",f"   {lbl} - tests/account/yr",t,"tests",NUM,note)
for seg,lbl in [("AG","Agencies"),("TF","Eng-heavy firms"),("MS","Mid-size firms"),("CH","Channel account")]:
    put(f"ACV_{seg}",f"   {lbl} - ACV/account/yr",f"={ref['TPA_'+seg]}*{ref['PRICE']}","$/yr",USD,"",is_input=False)

# 4. channel partner portfolio
band("4.  Channel: partner distribution portfolio")
A.cell(row,2,"Partner  (On=1 to include)").font=f_h2
for j,t in [(3,"On"),(4,"Clients"),(5,"ICP%"),(6,"Attach"),(7,"Accounts")]:
    A.cell(row,j,t).font=f_note; A.cell(row,j).alignment=center
row+=1
prow=[]
for name,on,clients,rel,att,src in PARTNERS:
    A.cell(row,2,name).font=f_label
    for j,v,fmt in [(3,on,'0'),(4,clients,NUM),(5,rel,PCT0),(6,att,PCT)]:
        c=A.cell(row,j,v); c.font=f_input; c.fill=fill_inp; c.number_format=fmt
        c.alignment=center if j==3 else right
    a=A.cell(row,7,f"=C{row}*D{row}*E{row}*F{row}"); a.font=f_calc; a.number_format=NUM
    A.cell(row,8,src).font=f_note; A.cell(row,8).alignment=center
    prow.append(row); row+=1
A.cell(row,2,"Channel base accounts (yr 3)").font=f_calcb
t=A.cell(row,7,f"=SUM(G{prow[0]}:G{prow[-1]})"); t.font=f_calcb; t.number_format=NUM; t.border=btop
ref["CH_ACCTS"]=f"Assumptions!$G${row}"; row+=1
put("ATTACH_SCALAR","Attach scalar (sensitivity handle)",1.00,"x",'0.00',"Multiplier on all partner attach rates.")

# 5. direct & reachability
band("5.  Direct capture & reachability")
put("DIR_PEN","Direct 3-yr penetration of serviceable",0.008,"%",PCT,"Thin slice; low trial-to-paid.")
put("RETENTION","Steady-state retention",0.90,"%",PCT,"BP uses 7% churn; 100% to date on a tiny base.",src="S8")
put("GLOB_SHARE","Focus-6 share of global",0.12,"%",PCT,"Grosses serviceable up to the full opportunity.")

# 6. talent-DB stream
band("6.  Talent-DB stream  (2nd revenue stream)")
put("PCT_DB","Profiles sourced from own DB (mature)",0.30,"%",PCT,"BP ramps 1% to 50%. Profiles = tests x this %.",src="S7")
put("PRICE_DB","Price per sourced profile",30,"$",USD,"$30 per profile (BP).",src="S7")

# ============================================================ BUILD
B=wb.create_sheet("Build"); B.sheet_view.showGridLines=False
for col,w in [("A",3),("B",34),("C",15),("D",13),("E",13),("F",16)]: B.column_dimensions[col].width=w
r=1
def bband(t,c2=6):
    global r
    B.cell(r,2,t).font=f_h1
    for c in range(2,c2+1): B.cell(r,c).fill=fill_h1
    r+=1
def bhead(cols):
    global r
    for i,h in enumerate(cols):
        c=B.cell(r,2+i,h); c.font=f_h2; c.fill=fill_sec; c.alignment=center
    r+=1
B.cell(r,2,"INVIRTUS  -  Revenue potential build").font=f_title; r+=1
B.cell(r,2,"Buyers x % addressable x ACV. Two streams. USD. Drivers live on Assumptions.").font=f_sub
B.merge_cells(start_row=r,start_column=2,end_row=r,end_column=6); r+=2

# A. serviceable today
bband("A.  Serviceable today  (focus-6)")
bhead(["Segment","Buyers (eff.)","% addressable","ACV/acct","Revenue ($/yr)"])
seg_first=r
for seg,lbl in [("AG","Tech recruiting agencies"),("TF","Engineering-heavy firms"),("MS","Mid-size firms")]:
    B.cell(r,2,lbl).font=f_label
    B.cell(r,3,f"={ref['EFF_'+seg]}").number_format=NUM
    B.cell(r,4,f"={ref['FIT_'+seg]}").number_format=PCT
    B.cell(r,5,f"={ref['ACV_'+seg]}").number_format=USD
    B.cell(r,6,f"=C{r}*D{r}*E{r}").number_format=USD
    for c in range(3,7): B.cell(r,c).font=f_calc
    r+=1
B.cell(r,2,"Serviceable today (assessments)").font=f_calcb
B.cell(r,6,f"=SUM(F{seg_first}:F{r-1})").number_format=USD; B.cell(r,6).font=f_calcb; B.cell(r,6).border=btop
SERV=f"Build!$F${r}"; r+=2

# B. full opportunity
bband("B.  Full opportunity  (all geos)")
B.cell(r,2,"Serviceable / focus-6 share of global").font=f_label
B.cell(r,6,f"={SERV}/{ref['GLOB_SHARE']}").number_format=USD; B.cell(r,6).font=f_calcb
FULL=f"Build!$F${r}"; r+=2

# C. 3-year reachable
bband("C.  3-year reachable  (channel + direct)")
B.cell(r,2,"Channel: accounts x attach scalar").font=f_label
B.cell(r,6,f"={ref['CH_ACCTS']}*{ref['ATTACH_SCALAR']}").number_format=NUM; B.cell(r,6).font=f_calc
CH_N=r; r+=1
B.cell(r,2,"Channel revenue (ret-adj)").font=f_label
B.cell(r,6,f"=F{CH_N}*{ref['ACV_CH']}*{ref['RETENTION']}").number_format=USD; B.cell(r,6).font=f_calc
CH_REV=r; r+=1
B.cell(r,2,"Direct revenue (pen. of serviceable, ret-adj)").font=f_label
B.cell(r,6,f"={SERV}*{ref['DIR_PEN']}*{ref['RETENTION']}").number_format=USD; B.cell(r,6).font=f_calc
D_REV=r; r+=1
B.cell(r,2,"3-year reachable (assessments)").font=f_calcb
B.cell(r,6,f"=F{CH_REV}+F{D_REV}").number_format=USD; B.cell(r,6).font=f_calcb; B.cell(r,6).border=btop
YR3=f"Build!$F${r}"; CH_REVc=f"Build!$F${CH_REV}"; D_REVc=f"Build!$F${D_REV}"; r+=2

# D. two streams
bband("D.  Revenue potential by stream")
bhead(["Tier","Assessments","Tests/yr","Talent-DB","Total ($/yr)"])
tier_rows={}
for tier,assess in [("Full opportunity",FULL),("Serviceable today",SERV),("3-year reachable",YR3)]:
    B.cell(r,2,tier).font=f_label
    B.cell(r,3,f"={assess}").number_format=USD
    B.cell(r,4,f"=C{r}/{ref['PRICE']}").number_format=NUM
    B.cell(r,5,f"=D{r}*{ref['PCT_DB']}*{ref['PRICE_DB']}").number_format=USD
    B.cell(r,6,f"=C{r}+E{r}").number_format=USD
    for c in range(3,7): B.cell(r,c).font=f_calc
    B.cell(r,6).font=f_calcb
    tier_rows[tier]=r; r+=1
TOT_FULL=f"Build!$F${tier_rows['Full opportunity']}"
TOT_SERV=f"Build!$F${tier_rows['Serviceable today']}"
TOT_YR3=f"Build!$F${tier_rows['3-year reachable']}"
ASSESS_YR3=f"Build!$C${tier_rows['3-year reachable']}"
DB_YR3=f"Build!$E${tier_rows['3-year reachable']}"
r+=1
B.cell(r,2,"Agencies and the firms they serve are counted separately (no double count). Channel and direct may overlap a little; treated as additive.").font=f_note
B.merge_cells(start_row=r,start_column=2,end_row=r,end_column=6); B.cell(r,2).alignment=left

# ============================================================ COVER
C=wb.create_sheet("Cover"); C.sheet_view.showGridLines=False
for col,w in [("A",3),("B",28),("C",22),("D",22),("E",6)]: C.column_dimensions[col].width=w
cr=2
C.cell(cr,2,"INVIRTUS").font=Font(name="Calibri",size=30,bold=True,color=NAVY); cr+=1
C.cell(cr,2,"Bottoms-Up Revenue Potential").font=Font(name="Calibri",size=14,color=GREY); cr+=1
C.cell(cr,2,"Candidate assessment platform  |  USD  |  June 2026").font=f_note; cr+=2
C.cell(cr,2,"THE ANSWER").font=f_h1
for c in range(2,6): C.cell(cr,c).fill=fill_h1
cr+=1
C.merge_cells(start_row=cr,start_column=2,end_row=cr+1,end_column=5)
C.cell(cr,2,("Two streams: assessment packs + talent-DB. The 3-year number is channel-led "
             "(partner portfolio ~76%, direct ~24%). Biggest swing: the channel attach scalar, "
             "then price and tests/account.")).font=Font(name="Calibri",size=11,color=NAVY)
C.cell(cr,2).alignment=left
for rr in range(cr,cr+2):
    for c in range(2,6): C.cell(rr,c).fill=fill_cov
cr+=3
C.cell(cr,2,"FULL OPPORTUNITY").font=f_biglbl
C.cell(cr,3,"SERVICEABLE TODAY").font=f_biglbl
C.cell(cr,4,"3-YEAR REACHABLE").font=f_biglbl
cr+=1
C.cell(cr,2,f"={TOT_FULL}").number_format=USDB; C.cell(cr,2).font=f_big
C.cell(cr,3,f"={TOT_SERV}").number_format=USDM; C.cell(cr,3).font=f_big
C.cell(cr,4,f"={TOT_YR3}").number_format=USDM; C.cell(cr,4).font=f_big
cr+=1
C.cell(cr,2,"All geos, full ACV").font=f_note
C.cell(cr,3,"Focus-6, addressable").font=f_note
C.cell(cr,4,"Channel + direct, 3-yr").font=f_note
cr+=2
C.cell(cr,2,"3-year reachable - detail").font=f_h2; cr+=1
for lbl,val,fmt in [("Assessment packs",ASSESS_YR3,USDM),("Talent-DB",DB_YR3,USDM),
                    ("Total",TOT_YR3,USDM),("  Channel",CH_REVc,USDM),("  Direct",D_REVc,USDM)]:
    C.cell(cr,2,lbl).font=f_label if not lbl.startswith("  ") else f_note
    c=C.cell(cr,3,f"={val}"); c.number_format=fmt; c.font=f_calcb; cr+=1
cr+=1
C.cell(cr,2,"Scenario levers (edit on Assumptions)").font=f_h2; cr+=1
C.cell(cr,2,"Market breadth").font=f_label
C.cell(cr,3,f"={ref['BREADTH_NAME']}").font=f_calcb; cr+=1
C.cell(cr,2,"Attach scalar").font=f_label
C.cell(cr,3,f"={ref['ATTACH_SCALAR']}").number_format='0.00'; C.cell(cr,3).font=f_calcb; cr+=2
for line in [
 "Blue cells are inputs; black cells are formulas.",
 "Breadth toggle (1 engineers / 2 some white-collar / 3 all white-collar) scales the buyer universe.",
 "Channel is a partner portfolio (Assumptions sec.4); each partner can be switched on or off.",
]:
    C.cell(cr,2,line).font=f_calc; C.merge_cells(start_row=cr,start_column=2,end_row=cr,end_column=5)
    C.cell(cr,2).alignment=left; cr+=1
cr+=1
C.cell(cr,2,"Input (editable)").font=f_input; C.cell(cr,2).fill=fill_inp
C.cell(cr,3,"Formula").font=f_calc

# ============================================================ SENSITIVITY
S=wb.create_sheet("Sensitivity"); S.sheet_view.showGridLines=False
for col,w in [("A",3),("B",34),("C",14),("D",14),("E",14)]: S.column_dimensions[col].width=w
sr=2
S.cell(sr,2,"INVIRTUS  -  Sensitivity").font=f_title; sr+=1
S.cell(sr,2,"3-year reachable (total) under low / base / high moves in each driver. Recomputed live from Assumptions.").font=f_sub
S.merge_cells(start_row=sr,start_column=2,end_row=sr,end_column=5); sr+=2

# closed-form total 3-yr = ret*p*(CH*sc*tch + K*pen)*(1+pctDB*priceDB/p),  K=sum EFF*FIT*TPA
K="("+"+".join(f"{ref['EFF_'+s]}*{ref['FIT_'+s]}*{ref['TPA_'+s]}" for s in ["AG","TF","MS"])+")"
def yr3(p,tch,sc,pen):
    assess=f"{ref['RETENTION']}*{p}*({ref['CH_ACCTS']}*{sc}*{tch}+{K}*{pen})"
    return f"=({assess})*(1+{ref['PCT_DB']}*{ref['PRICE_DB']}/{p})"
P=ref['PRICE']; TCH=ref['TPA_CH']; SC=ref['ATTACH_SCALAR']; PEN=ref['DIR_PEN']
for i,h in enumerate(["Driver","Low","Base","High"]):
    c=S.cell(sr,2+i,h); c.font=f_h2; c.fill=fill_sec; c.alignment=center
sr+=1
rows=[
 ("Channel attach scalar (0.5 / base / 1.5)", yr3(P,TCH,"0.5",PEN), yr3(P,TCH,SC,PEN), yr3(P,TCH,"1.5",PEN)),
 ("Net price per test ($9 / base / $15)",     yr3("9",TCH,SC,PEN),  yr3(P,TCH,SC,PEN), yr3("15",TCH,SC,PEN)),
 ("Channel tests/account (100 / base / 300)", yr3(P,"100",SC,PEN),  yr3(P,TCH,SC,PEN), yr3(P,"300",SC,PEN)),
 ("Direct penetration (0.4% / base / 1.2%)",  yr3(P,TCH,SC,"0.004"),yr3(P,TCH,SC,PEN), yr3(P,TCH,SC,"0.012")),
]
for lbl,lo,ba,hi in rows:
    S.cell(sr,2,lbl).font=f_label
    for c,v in [(3,lo),(4,ba),(5,hi)]:
        S.cell(sr,c,v).number_format=USDM; S.cell(sr,c).font=f_calc
    sr+=1

# ============================================================ SOURCES
SO=wb.create_sheet("Sources"); SO.sheet_view.showGridLines=False
for col,w in [("A",3),("B",6),("C",32),("D",30),("E",54),("F",12),("G",9)]: SO.column_dimensions[col].width=w
so=2
SO.cell(so,2,"INVIRTUS  -  Sources").font=f_title; so+=1
SO.cell(so,2,"One row per researched input. Triangulated where official counts are absent.").font=f_sub
SO.merge_cells(start_row=so,start_column=2,end_row=so,end_column=7); so+=2
for i,h in enumerate(["Key","Input","Source","URL","Date","Conf."]):
    c=SO.cell(so,2+i,h); c.font=f_h2; c.fill=fill_sec; c.alignment=center
so+=1
SOURCES=[
 ("S1","France: agencies & firms","INSEE SIRENE; Modeles de Business Plan",
   "public.opendatasoft.com/.../sirene-v3 ; modelesdebusinessplan.com","2026-06-16","High/Low"),
 ("S2","UK: agencies & firms","REC / IBISWorld; GOV.UK Digital & Tech 2026",
   "rec.uk.com ; gov.uk/.../digital-and-technologies-sector-plan","2026-06-16","Med/High"),
 ("S3","UAE: agencies & firms","SIA; Lusha; DIC; MAGNiTT; Min. of Economy",
   "magnitt.com/en-ae/startups ; thenationalnews.com","2026-06-16","Med"),
 ("S4","Morocco: agencies & firms","EcoActu/MIEPEEC; OMTPME 2025; UM6P; APEBI",
   "ecoactu.ma ; lematin.ma","2026-06-16","Med/Low"),
 ("S5","KSA: agencies & firms","Monsha'at SME Monitor Q4'23; CST Manassa",
   "argaam.com/.../1709526 ; my.gov.sa/en/news/10102","2026-06-16","Med"),
 ("S6","Egypt: agencies & firms","CAPMAS Economic Census; ITIDA; MAGNiTT",
   "censusinfo.capmas.gov.eg ; itida.gov.eg","2026-06-16","Med/Low"),
 ("S7","Pricing & DB ($14/test, packs, $30/profile)","Invirtus finalized BP; TestGorilla/TestDome/Adaface",
   "Invirtus_business_plan.xlsx ; testgorilla.com/pricing","2026-06-17","High"),
 ("S8","Traction, retention, channel anchor","Invirtus April 2026 IC Update; GTM notes",
   "Notion: Invirtus - April 2026 IC Update","2026-06-16","Co. data"),
 ("S9","Socium (2nd ATS)","Invirtus GTM notes",
   "Notion: Invirtus - GTM","2026-06-16","Co. data"),
]
for k,inp,src,url,dt,conf in SOURCES:
    SO.cell(so,2,k).font=f_calcb
    SO.cell(so,3,inp).font=f_calc; SO.cell(so,3).alignment=left
    SO.cell(so,4,src).font=f_calc; SO.cell(so,4).alignment=left
    SO.cell(so,5,url).font=f_note; SO.cell(so,5).alignment=left
    SO.cell(so,6,dt).font=f_calc; SO.cell(so,6).alignment=center
    SO.cell(so,7,conf).font=f_calc; SO.cell(so,7).alignment=center
    for c in range(2,8): SO.cell(so,c).border=border
    so+=1
so+=1
SO.cell(so,2,"Note: 'medium enterprise' bands differ by country; mid-size counts are the closest official proxy x an engineer-hiring share. The %-addressable and engineer-hiring share are the softest inputs - pressure-test before widening breadth.").font=f_note
SO.merge_cells(start_row=so,start_column=2,end_row=so+1,end_column=7); SO.cell(so,2).alignment=left

# order
desired=["Cover","Assumptions","Build","Sensitivity","Sources"]
wb._sheets.sort(key=lambda s:desired.index(s.title))
wb.save("Invirtus_Market_Sizing.xlsx")
print("saved; assumptions rows:",row)

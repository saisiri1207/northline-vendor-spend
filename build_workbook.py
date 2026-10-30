"""Build Northline_Vendor_Spend.xlsx (fictional sample data). Yellow cells are inputs, black cells are formulas."""
import random
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter as L
random.seed(11)
Y = PatternFill("solid", fgColor="FFF2CC"); H = PatternFill("solid", fgColor="1F3864"); HF = Font(bold=True, color="FFFFFF")
months = ["Jul 2025","Aug 2025","Sep 2025","Oct 2025","Nov 2025","Dec 2025","Jan 2026","Feb 2026","Mar 2026","Apr 2026","May 2026","Jun 2026"]
# vendor, category, cost center, annual base spend, lever, current rate, target rate
vendors = [
 ("V001 Harbor Card Processing","Processing fees","Finance",2400000,"Rate benchmark",0.0290,0.0268),
 ("V002 Summit Payments","Processing fees","Finance",900000,"Rate benchmark",0.0310,0.0290),
 ("V003 Atlas Freight","Logistics","Operations",1650000,"Volume tier",0.00,0.03),
 ("V004 Atlas Freight Inc","Logistics","Operations",420000,"Duplicate vendor",0.00,0.05),
 ("V005 BrightPath Staffing","Contract labor","Operations",1100000,"Rate benchmark",0.180,0.170),
 ("V006 Cloudway Hosting","Software","IT",780000,"Volume tier",0.00,0.03),
 ("V007 Cloudway Hosting LLC","Software","IT",210000,"Duplicate vendor",0.00,0.08),
 ("V008 Pinecrest Office Supply","Office","G&A",260000,"Volume tier",0.00,0.06),
 ("V009 Metro Print Co","Marketing","Marketing",340000,"Duplicate vendor",0.00,0.04),
 ("V010 Keystone Telecom","Telecom","IT",390000,"Rate benchmark",0.00,0.04),
 ("V011 Lakeside Facilities","Facilities","G&A",620000,"None",0.00,0.00),
 ("V012 Northgate Legal","Professional services","G&A",480000,"None",0.00,0.00),
]
wb = Workbook()
ws = wb.active; ws.title = "README"
for i, t in enumerate(["Northline Vendor Spend and Savings Model","Northline Consumer Products is a fictional company. All figures are sample data.",
  "Yellow cells are inputs. Black cells are formulas.","Sheets: Vendors, Monthly_Spend, Savings, Summary",
  "Savings logic: Rate benchmark processing fees = spend x (current rate minus target rate) / current rate.",
  "Rate benchmark other = spend x target reduction %. Volume tier and duplicate vendor = spend x tier discount %."],1):
    ws.cell(i,1,t)
ws["A1"].font = Font(bold=True, size=14)
v = wb.create_sheet("Vendors")
hdr = ["Vendor","Category","Cost center","Lever","Current fee rate","Target rate or discount"]
for j,h in enumerate(hdr,1): c=v.cell(1,j,h); c.fill=H; c.font=HF
for i,(n,cat,cc,_,lev,cur,tgt) in enumerate(vendors,2):
    for j,val in enumerate([n,cat,cc,lev,cur,tgt],1):
        c=v.cell(i,j,val); c.fill=Y
    v.cell(i,5).number_format='0.00%'; v.cell(i,6).number_format='0.00%'
m = wb.create_sheet("Monthly_Spend")
m.cell(1,1,"Vendor").fill=H; m.cell(1,1).font=HF
for j,mo in enumerate(months,2): c=m.cell(1,j,mo); c.fill=H; c.font=HF
c=m.cell(1,14,"Annual total"); c.fill=H; c.font=HF
season=[0.95,0.97,1.0,1.03,1.08,1.15,0.92,0.93,0.98,1.0,1.0,0.99]; s=sum(season)
spend={}
for i,vd in enumerate(vendors,2):
    m.cell(i,1,f"=Vendors!A{i}")
    row=[round(vd[3]*season[k]/s*random.uniform(0.97,1.03),0) for k in range(12)]
    spend[vd[0]]=sum(row)
    for k,x in enumerate(row,2): c=m.cell(i,k,x); c.fill=Y; c.number_format='#,##0'
    m.cell(i,14,f"=SUM(B{i}:M{i})").number_format='#,##0'
n=len(vendors)+1
m.cell(n+1,1,"Total"); 
for k in range(2,15): m.cell(n+1,k,f"=SUM({L(k)}2:{L(k)}{n})").number_format='#,##0'
sv = wb.create_sheet("Savings")
for j,h in enumerate(["Vendor","Lever","Annual spend","Current rate","Target","Annual savings","Rank"],1):
    c=sv.cell(1,j,h); c.fill=H; c.font=HF
calc=[]
for i,vd in enumerate(vendors,2):
    sv.cell(i,1,f"=Vendors!A{i}"); sv.cell(i,2,f"=Vendors!D{i}")
    sv.cell(i,3,f"=Monthly_Spend!N{i}").number_format='#,##0'
    sv.cell(i,4,f"=Vendors!E{i}").number_format='0.00%'; sv.cell(i,5,f"=Vendors!F{i}").number_format='0.00%'
    sv.cell(i,6,f'=IF(B{i}="None",0,IF(D{i}>0,C{i}*(D{i}-E{i})/D{i},C{i}*E{i}))').number_format='#,##0'
    sv.cell(i,7,f"=IF(F{i}=0,\"\",RANK(F{i},$F$2:$F${n}))")
    a=spend[vd[0]]; cur,tgt=vd[5],vd[6]
    calc.append((vd[0],vd[4],a, 0 if vd[4]=="None" else (a*(cur-tgt)/cur if cur>0 else a*tgt)))
sv.cell(n+1,1,"Total"); sv.cell(n+1,3,f"=SUM(C2:C{n})").number_format='#,##0'; sv.cell(n+1,6,f"=SUM(F2:F{n})").number_format='#,##0'
su = wb.create_sheet("Summary")
su.append(["Lever","Annual savings","Share of spend"])
for c in su[1]: c.fill=H; c.font=HF
for r,lev in enumerate(["Rate benchmark","Volume tier","Duplicate vendor"],2):
    su.cell(r,1,lev); su.cell(r,2,f'=SUMIF(Savings!B2:B{n},A{r},Savings!F2:F{n})').number_format='#,##0'
    su.cell(r,3,f"=B{r}/Savings!C{n+1}").number_format='0.0%'
su.cell(5,1,"Total"); su.cell(5,2,"=SUM(B2:B4)").number_format='#,##0'; su.cell(5,3,f"=B5/Savings!C{n+1}").number_format='0.0%'
for w_ in wb.worksheets:
    w_.column_dimensions['A'].width=32
    for k in range(2,15): w_.column_dimensions[L(k)].width=14
wb.save("Northline_Vendor_Spend.xlsx")
tot=sum(x[2] for x in calc); sav=sum(x[3] for x in calc)
print(f"Total annual spend {tot:,.0f}")
for lev in ["Rate benchmark","Volume tier","Duplicate vendor"]:
    print(lev, f"{sum(x[3] for x in calc if x[1]==lev):,.0f}")
print(f"Total savings {sav:,.0f} ({sav/tot:.1%} of spend)")
for x in sorted(calc,key=lambda x:-x[3])[:6]: print(f"{x[0]} | {x[1]} | {x[2]:,.0f} | {x[3]:,.0f}")

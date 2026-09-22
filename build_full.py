#!/usr/bin/env python3
"""Comprehensive, editable wedding budget workbook — plain Arial, neutral theme.
Sheets: Ringkasan (dashboard + charts) · Budget (flat editable table) · Timeline (savings runway)."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.utils import get_column_letter

AR="Arial"
INK="222222"; MUT="808080"; NAVY="2F5496"; SLATE="44546A"
GREEN="2E7D32"; RED="C0392B"; AMBER="B7791F"
HEADFILL=PatternFill("solid",fgColor="44546A")
BAND    =PatternFill("solid",fgColor="EDEDED")
TOTFILL =PatternFill("solid",fgColor="DDE3EC")
KPIFILL =PatternFill("solid",fgColor="F2F5FA")
thin=Side(style="thin",color="D0D0D0"); medb=Side(style="medium",color="9AA5B5")
ROW=Border(bottom=thin)
BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
L=Alignment(horizontal="left",vertical="center",wrap_text=True)
LT=Alignment(horizontal="left",vertical="top",wrap_text=True)
C=Alignment(horizontal="center",vertical="center",wrap_text=True)
R=Alignment(horizontal="right",vertical="center")
RP='"Rp"#,##0'
def f(sz=10,b=False,color=INK,it=False): return Font(name=AR,size=sz,bold=b,color=color,italic=it)

# ---------------------------------------------------------------- data
# (Kategori, Item, Qty/Pax, Biaya, Status, Dibayar oleh, Catatan)
items=[
 ("Lamaran","Tenda + kursi (sewa)","",1500000,"Belum","Blugether","Di rumah mempelai wanita"),
 ("Lamaran","Dekorasi & bunga","",1500000,"Belum","Blugether",""),
 ("Lamaran","Makanan & minuman","",1000000,"Belum","Blugether","Masak sendiri di rumah"),
 ("Lamaran","Seserahan / hantaran","",500000,"Belum","Blugether","Nampan + wrapping"),
 ("Lamaran","Dokumentasi lamaran","½ hari",300000,"Belum","Blugether",""),
 ("Lamaran","Makeup mempelai wanita (lamaran)","",2000000,"DP dibayar","Blugether","DP sudah dibayar"),
 ("Lamaran","Buffer lamaran","",200000,"Belum","Blugether",""),

 ("Siraman","Upacara siraman","",3000000,"Belum","Blugether","Bunga setaman, air 7 sumber, kendi, dawet, tumpeng"),
 ("Siraman","Katering keluarga (siraman)","",1500000,"Belum","Ortu Wanita","Makan keluarga di rumah"),
 ("Siraman","Lulur & henna pengantin","",1500000,"Belum","Blugether",""),

 ("Akad & Adat","Perlengkapan panggih (temu manten)","",2000000,"Belum","Blugether","Kembar mayang ×2, suruh, telur, kacar-kucur, sindur"),
 ("Akad & Adat","Hadroh / rebana (akad)","",3000000,"Belum","Ortu Pria",""),
 ("Akad & Adat","Penghulu / KUA (akad)","",600000,"Belum","Blugether","Di luar jam kantor"),
 ("Akad & Adat","Mahar + replika mahar","",1200000,"Belum","Blugether","Simbolis"),
 ("Akad & Adat","Medical check-up (syarat nikah)","2 org",1000000,"Belum","Blugether",""),

 ("Venue & F&B","The Bimasena (Rooang) — minimum spending","100+ pax",150000000,"Belum","Blugether","Min. spend; sudah termasuk F&B 100+ pax"),
 ("Venue & F&B","Tambahan pax / upgrade menu","",0,"Belum","Blugether","Isi bila melebihi minimum"),
 ("Venue & F&B","Makan vendor & panitia","",3000000,"Belum","Blugether",""),

 ("Resepsi & Dekor","Dekorasi & bunga resepsi","",12000000,"Belum","Blugether","Melati putih: pelaminan, entrance, aisle, meja"),
 ("Resepsi & Dekor","Bridal bouquet + melati groom & bride","",2000000,"Belum","Blugether",""),
 ("Resepsi & Dekor","Karpet (bila perlu)","",0,"Belum","Blugether",""),

 ("Hiburan & Teknis","Live music / band","",8000000,"Belum","Blugether","Akustik → band"),
 ("Hiburan & Teknis","MC / pranatacara","",4000000,"Belum","Blugether","Bilingual"),
 ("Hiburan & Teknis","Koordinator hari-H (WO on-the-day)","",5000000,"Belum","Blugether","Jaga rundown, cue vendor"),
 ("Hiburan & Teknis","Sound system & lighting","",5000000,"Belum","Blugether","Bila belum termasuk venue"),
 ("Hiburan & Teknis","Live streaming","",3000000,"Belum","Blugether","Opsional"),

 ("Busana, Rias & Rambut","Busana & rias mempelai wanita","",18000000,"Belum","Blugether","Paes ageng: siraman, akad, resepsi"),
 ("Busana, Rias & Rambut","Hairdo mempelai wanita","",2000000,"Belum","Blugether",""),
 ("Busana, Rias & Rambut","Sewa beskap mempelai pria (akad)","",4000000,"Belum","Blugether","Hanya beskap akad; tak perlu jas lamaran/resepsi"),
 ("Busana, Rias & Rambut","Busana adat orang tua","",4000000,"Belum","Ortu","Kebaya / beskap"),
 ("Busana, Rias & Rambut","Makeup orang tua & keluarga inti","",4000000,"Belum","Ortu",""),
 ("Busana, Rias & Rambut","Seragam panitia / keluarga","± 20 org",5000000,"Belum","Blugether","± 20 × 250rb"),
 ("Busana, Rias & Rambut","Accessories (mahkota, printilan)","",3000000,"Belum","Blugether",""),

 ("Dokumentasi","Foto & video (2 hari)","siraman + hari-H",12000000,"Belum","Blugether",""),
 ("Dokumentasi","Foto prewedding","",3000000,"Belum","Blugether","Opsional"),

 ("Cincin & Pribadi","Cincin tunangan + cincin nikah","",23000000,"Lunas","Blugether","Total kedua cincin — sudah dibeli"),
 ("Cincin & Pribadi","Hair care / perawatan pengantin","",1500000,"Belum","Blugether",""),

 ("Souvenir & Tamu","Souvenir tamu","100 pcs",2500000,"Belum","Blugether","± 100 × 25rb"),
 ("Souvenir & Tamu","Souvenir VIP / besan","",1500000,"Belum","Blugether",""),
 ("Souvenir & Tamu","Usher & buku tamu","",1500000,"Belum","Blugether",""),

 ("Logistik & Keluarga","Undangan digital","",300000,"Belum","Blugether",""),
 ("Logistik & Keluarga","Undangan cetak","100 pcs",1500000,"Belum","Ortu",""),
 ("Logistik & Keluarga","Kamar hotel — bridal suite (1 malam)","",5000000,"Belum","Blugether","Malam pernikahan"),
 ("Logistik & Keluarga","Kamar hotel tambahan (MUA/keluarga)","",3000000,"Belum","Blugether",""),
 ("Logistik & Keluarga","Transportasi & mobil pengantin","",3500000,"Belum","Blugether",""),
 ("Logistik & Keluarga","Family pocket money","",5000000,"Belum","Ortu",""),
 ("Logistik & Keluarga","Wedding cake & dessert","",2500000,"Belum","Blugether",""),
]
cats=[]
for it in items:
    if it[0] not in cats: cats.append(it[0])
payers=["Blugether","Ortu Wanita","Ortu Pria","Ortu","Lainnya"]
statuses=["Lunas","DP dibayar","Belum"]

wb=Workbook()

# ================================================================ BUDGET
b=wb.active; b.title="Budget"; b.sheet_view.showGridLines=False
for col,w in {"A":5,"B":20,"C":40,"D":13,"E":16,"F":12,"G":14,"H":34}.items():
    b.column_dimensions[col].width=w
b["A1"]="BUDGET PERNIKAHAN — [ Mempelai Wanita ] & [ Mempelai Pria ]"
b["A1"].font=f(15,True,NAVY); b.merge_cells("A1:H1"); b.row_dimensions[1].height=24
b["A2"]="Adat Jawa · The Bimasena (Rooang), Dharmawangsa · semua angka Rupiah · sel putih bisa diedit"
b["A2"].font=f(9,color=MUT,it=True); b.merge_cells("A2:H2")
HDR=3
for i,h in enumerate(["No","Kategori","Item","Qty / Pax","Biaya","Status","Dibayar oleh","Catatan"]):
    c=b.cell(HDR,i+1,h); c.font=f(10,True,"FFFFFF"); c.fill=HEADFILL
    c.alignment=C if i in(0,3,4,5,6) else L; c.border=BOX
b.row_dimensions[HDR].height=20
first=HDR+1; r=first; n=0; venue_row=rings_row=None
for kat,item,qty,amt,status,payer,note in items:
    n+=1
    b.cell(r,1,n).font=f(9,color=MUT); b.cell(r,1).alignment=C
    b.cell(r,2,kat).font=f(9,color=SLATE); b.cell(r,2).alignment=LT
    b.cell(r,3,item).font=f(10); b.cell(r,3).alignment=LT
    b.cell(r,4,qty).font=f(9,color=MUT); b.cell(r,4).alignment=C
    e=b.cell(r,5,amt); e.font=f(10); e.number_format=RP; e.alignment=R
    s=b.cell(r,6,status); s.font=f(9,True,GREEN if status=="Lunas" else AMBER if "DP" in status else RED); s.alignment=C
    b.cell(r,7,payer).font=f(9); b.cell(r,7).alignment=C
    b.cell(r,8,note).font=f(9,color=MUT); b.cell(r,8).alignment=LT
    for cc in range(1,9): b.cell(r,cc).border=ROW
    if "Bimasena" in item: venue_row=r
    if item.startswith("Cincin"): rings_row=r
    b.row_dimensions[r].height=26; r+=1
# contingency row
n+=1
b.cell(r,1,n).font=f(9,color=MUT); b.cell(r,1).alignment=C
b.cell(r,2,"Cadangan").font=f(9,color=SLATE); b.cell(r,2).alignment=LT
b.cell(r,3,"Biaya tak terduga (10%)").font=f(10); b.cell(r,3).alignment=LT
e=b.cell(r,5,f"=ROUND(0.1*(SUM(E{first}:E{r-1})-E{venue_row}-E{rings_row}),-3)")
e.font=f(10); e.number_format=RP; e.alignment=R
b.cell(r,6,"Belum").font=f(9,True,RED); b.cell(r,6).alignment=C
b.cell(r,7,"Blugether").font=f(9); b.cell(r,7).alignment=C
b.cell(r,8,"10% pos non-fix (di luar venue & cincin)").font=f(9,color=MUT); b.cell(r,8).alignment=LT
for cc in range(1,9): b.cell(r,cc).border=ROW
last=r; r+=2
# totals
def brow(r,lab,formula,color=INK):
    b.cell(r,3,lab).font=f(11,True,color); b.cell(r,3).alignment=R
    t=b.cell(r,5,formula); t.font=f(12,True,color); t.number_format=RP; t.alignment=R
    for cc in range(3,6): b.cell(r,cc).fill=TOTFILL; b.cell(r,cc).border=Border(top=medb,bottom=medb)
    b.row_dimensions[r].height=22
brow(r,"TOTAL ANGGARAN",f"=SUM(E{first}:E{last})"); tot_row=r; r+=1
brow(r,"Sudah dibayar (Lunas)",f'=SUMIF(F{first}:F{last},"Lunas",E{first}:E{last})',GREEN); paid_row=r; r+=1
brow(r,"Sisa yang belum dibayar",f"=E{tot_row}-E{paid_row}",RED); r+=1
b.auto_filter.ref=f"A{HDR}:H{last}"
b.freeze_panes="A4"
dv_s=DataValidation(type="list",formula1='"'+",".join(statuses)+'"',allow_blank=True)
dv_p=DataValidation(type="list",formula1='"'+",".join(payers)+'"',allow_blank=True)
b.add_data_validation(dv_s); b.add_data_validation(dv_p)
dv_s.add(f"F{first}:F{last}"); dv_p.add(f"G{first}:G{last}")
ER=f"E{first}:E{last}"; FR=f"F{first}:F{last}"; GR=f"G{first}:G{last}"; BR=f"B{first}:B{last}"

# ================================================================ RINGKASAN
s=wb.create_sheet("Ringkasan",0); s.sheet_view.showGridLines=False
for col,w in {"A":3,"B":26,"C":18,"D":3,"E":18,"F":16,"G":16}.items():
    s.column_dimensions[col].width=w
s["B1"]="RINGKASAN — RENCANA PERNIKAHAN"; s["B1"].font=f(16,True,NAVY); s.merge_cells("B1:G1"); s.row_dimensions[1].height=26
s["B2"]="Semua angka Rupiah. Ubah di sheet Budget → ringkasan & grafik ikut berubah."; s["B2"].font=f(9,color=MUT,it=True); s.merge_cells("B2:G2")
# KPI cards
kpis=[("Total anggaran",f"=Budget!E{tot_row}"),
      ("Sudah dibayar",f"=Budget!E{paid_row}"),
      ("Sisa",f"=Budget!E{tot_row}-Budget!E{paid_row}"),
      ("Beban Blugether",f'=SUMIF(Budget!{GR},"Blugether",Budget!{ER})'),
      ("Beban Orang tua",f'=SUMIF(Budget!{GR},"Ortu*",Budget!{ER})+SUMIF(Budget!{GR},"Lainnya",Budget!{ER})')]
kr=4
for lab,form in kpis:
    s.cell(kr,2,lab).font=f(10,color=SLATE); s.cell(kr,2).fill=KPIFILL; s.cell(kr,2).border=BOX; s.cell(kr,2).alignment=L
    v=s.cell(kr,3,form); v.font=f(12,True,NAVY); v.number_format=RP; v.fill=KPIFILL; v.border=BOX; v.alignment=R
    kr+=1
# funding inputs (kuning = bisa diedit)
YEL=PatternFill("solid",fgColor="FFF7DA")
def lab(cell,txt): s[cell]=txt; s[cell].font=f(10,color=SLATE); s[cell].alignment=L
def inp(cell,val,fmt=RP): s[cell]=val; s[cell].font=f(11,True); s[cell].number_format=fmt; s[cell].alignment=R; s[cell].fill=YEL; s[cell].border=BOX
def out(cell,form,color=NAVY,fill=None):
    s[cell]=form; s[cell].font=f(11,True,color); s[cell].number_format=RP; s[cell].alignment=R; s[cell].border=BOX
    if fill: s[cell].fill=fill
lab("E4","Saldo Blugether (Sep 2026)"); inp("F4",50000000)
lab("E5","Menabung / bulan"); inp("F5",5000000)
lab("E6","Bulan s/d nikah (Syawal ±Apr 2028)"); inp("F6",18,'0" bln"')
lab("E7","Kapasitas Blugether s/d nikah"); out("F7","=F4+F5*F6")
lab("E8","Saran kontribusi orang tua (min)"); out("F8","=MAX(0,C6-F7)",RED,PatternFill("solid",fgColor="FBE9E7"))
# category summary table
ct=10
s.cell(ct,2,"Anggaran per kategori").font=f(11,True,SLATE); s.merge_cells(f"B{ct}:C{ct}")
s.cell(ct+1,2,"Kategori").font=f(9,True,"FFFFFF"); s.cell(ct+1,2).fill=HEADFILL; s.cell(ct+1,2).border=BOX
s.cell(ct+1,3,"Total").font=f(9,True,"FFFFFF"); s.cell(ct+1,3).fill=HEADFILL; s.cell(ct+1,3).border=BOX
cr=ct+2
for cat in cats+["Cadangan"]:
    s.cell(cr,2,cat).font=f(10); s.cell(cr,2).border=ROW; s.cell(cr,2).alignment=L
    v=s.cell(cr,3,f'=SUMIF(Budget!{BR},B{cr},Budget!{ER})'); v.font=f(10); v.number_format=RP; v.border=ROW; v.alignment=R
    cr+=1
cat_first=ct+2; cat_last=cr-1
# status summary
st=ct
s.cell(st,5,"Status pembayaran").font=f(11,True,SLATE); s.merge_cells(f"E{st}:F{st}")
s.cell(st+1,5,"Status").font=f(9,True,"FFFFFF"); s.cell(st+1,5).fill=HEADFILL; s.cell(st+1,5).border=BOX
s.cell(st+1,6,"Total").font=f(9,True,"FFFFFF"); s.cell(st+1,6).fill=HEADFILL; s.cell(st+1,6).border=BOX
sr=st+2
for stt in statuses:
    s.cell(sr,5,stt).font=f(10); s.cell(sr,5).border=ROW; s.cell(sr,5).alignment=L
    v=s.cell(sr,6,f'=SUMIF(Budget!{FR},E{sr},Budget!{ER})'); v.font=f(10); v.number_format=RP; v.border=ROW; v.alignment=R
    sr+=1
st_first=st+2; st_last=sr-1
# payer summary
pt=sr+1
s.cell(pt,5,"Siapa membayar").font=f(11,True,SLATE); s.merge_cells(f"E{pt}:F{pt}")
s.cell(pt+1,5,"Pembayar").font=f(9,True,"FFFFFF"); s.cell(pt+1,5).fill=HEADFILL; s.cell(pt+1,5).border=BOX
s.cell(pt+1,6,"Total").font=f(9,True,"FFFFFF"); s.cell(pt+1,6).fill=HEADFILL; s.cell(pt+1,6).border=BOX
pr=pt+2
pay_groups=[("Blugether",f'=SUMIF(Budget!{GR},"Blugether",Budget!{ER})'),
            ("Orang tua",f'=SUMIF(Budget!{GR},"Ortu*",Budget!{ER})'),
            ("Lainnya",f'=SUMIF(Budget!{GR},"Lainnya",Budget!{ER})')]
for lab,form in pay_groups:
    s.cell(pr,5,lab).font=f(10); s.cell(pr,5).border=ROW; s.cell(pr,5).alignment=L
    v=s.cell(pr,6,form); v.font=f(10); v.number_format=RP; v.border=ROW; v.alignment=R
    pr+=1
pay_first=pt+2; pay_last=pr-1
# charts
barc=BarChart(); barc.type="bar"; barc.title="Anggaran per kategori"; barc.height=8.5; barc.width=15; barc.legend=None
data=Reference(s,min_col=3,min_row=ct+1,max_row=cat_last); cat_ref=Reference(s,min_col=2,min_row=cat_first,max_row=cat_last)
barc.add_data(data,titles_from_data=True); barc.set_categories(cat_ref)
s.add_chart(barc,"B"+str(cat_last+3))

pie=PieChart(); pie.title="Status pembayaran"; pie.height=6.5; pie.width=8
d2=Reference(s,min_col=6,min_row=st+1,max_row=st_last); c2=Reference(s,min_col=5,min_row=st_first,max_row=st_last)
pie.add_data(d2,titles_from_data=True); pie.set_categories(c2)
s.add_chart(pie,"E"+str(pay_last+3))

pie2=PieChart(); pie2.title="Siapa membayar"; pie2.height=6.5; pie2.width=8
d3=Reference(s,min_col=6,min_row=pt+1,max_row=pay_last); c3=Reference(s,min_col=5,min_row=pay_first,max_row=pay_last)
pie2.add_data(d3,titles_from_data=True); pie2.set_categories(c3)
s.add_chart(pie2,"E"+str(pay_last+18))

# ================================================================ TIMELINE
t=wb.create_sheet("Timeline"); t.sheet_view.showGridLines=False
for col,w in {"A":3,"B":14,"C":20,"D":20,"E":20,"F":30}.items(): t.column_dimensions[col].width=w
t["B1"]="TIMELINE MENABUNG"; t["B1"].font=f(15,True,NAVY); t.merge_cells("B1:F1"); t.row_dimensions[1].height=24
t["B2"]="Tabungan kumulatif vs beban Blugether. Ubah saldo & tabungan di sheet Ringkasan."; t["B2"].font=f(9,color=MUT,it=True); t.merge_cells("B2:F2")
hdr=4
for i,h in enumerate(["Bulan","Tabungan kumulatif","Beban Blugether","Total anggaran"]):
    c=t.cell(hdr,i+2,h); c.font=f(10,True,"FFFFFF"); c.fill=HEADFILL; c.border=BOX; c.alignment=C
months=["Sep 2026","Okt 2026","Nov 2026","Des 2026"]+[f"{m} 2027" for m in ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"]]+[f"{m} 2028" for m in ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"]]
rr=hdr+1
for i,m in enumerate(months):
    t.cell(rr,2,m).font=f(10); t.cell(rr,2).alignment=C; t.cell(rr,2).border=ROW
    v=t.cell(rr,3,f"=Ringkasan!$F$4+Ringkasan!$F$5*{i}"); v.font=f(10); v.number_format=RP; v.alignment=R; v.border=ROW
    bb=t.cell(rr,4,"=Ringkasan!$C$7"); bb.font=f(10,color=MUT); bb.number_format=RP; bb.alignment=R; bb.border=ROW
    tt=t.cell(rr,5,f"=Budget!$E${tot_row}"); tt.font=f(10,color=MUT); tt.number_format=RP; tt.alignment=R; tt.border=ROW
    rr+=1
tlast=rr-1
line=LineChart(); line.title="Runway menabung (Rp)"; line.height=9; line.width=20; line.style=2
dref=Reference(t,min_col=3,min_row=hdr,max_col=5,max_row=tlast)
cref=Reference(t,min_col=2,min_row=hdr+1,max_row=tlast)
line.add_data(dref,titles_from_data=True); line.set_categories(cref)
line.y_axis.numFmt='#,##0'; line.y_axis.title="Rupiah"; line.x_axis.title="Bulan"
t.add_chart(line,"B"+str(tlast+3))
# milestones
mr=tlast+3
t.cell(mr,5,"Milestone").font=f(11,True,SLATE)
for mlab,dt in [("Lamaran / tunangan","November 2026"),("Siraman","H-1 pernikahan"),("Akad & resepsi — Syawal 1449 H","± April 2028")]:
    mr+=1
    t.cell(mr,5,dt).font=f(9,True,SLATE); t.cell(mr,5).alignment=L
    t.cell(mr,6,mlab).font=f(10); t.cell(mr,6).alignment=L

out="/Users/gomobile/Documents/Project/Wedding-Budget/Budget_Pernikahan_Bimasena.xlsx"
wb.save(out); print("saved",out,"| items:",n)

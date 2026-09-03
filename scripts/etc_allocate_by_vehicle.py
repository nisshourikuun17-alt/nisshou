import re, collections, pdfplumber, copy, openpyxl
from openpyxl.styles import Font, Alignment

PDF='/root/.claude/uploads/e46d2994-6f12-5b86-9532-97b1f22855c3/6be19bcb-____.pdf'
XLSX='/root/.claude/uploads/e46d2994-6f12-5b86-9532-97b1f22855c3/e7bb02bd-____.xlsx'
OUT='/home/user/nisshou/8月分_高速料金_車番別.xlsx'

time_re=re.compile(r'^\d{2}:\d{2}(\s+\d{2}:\d{2})?\b')
card_re=re.compile(r'\*{3,}\d+$')

lines=[]
with pdfplumber.open(PDF) as pdf:
    for p in pdf.pages:
        lines += (p.extract_text() or '').split('\n')

totals=collections.Counter(); counts=collections.Counter(); cur=None; bad=[]
for ln in lines:
    s=ln.strip()
    if time_re.match(s):
        last=s.split()[-1]
        cur=int(last) if re.fullmatch(r'\d+',last) else None
        if cur is None: bad.append(s)
    elif card_re.search(s):
        toks=s.split(); fee=toks[-2].replace(',','')
        if not re.fullmatch(r'\d+',fee) or cur is None:
            bad.append(s); continue
        totals[cur]+=int(fee); counts[cur]+=1; cur=None
assert not bad, bad
print('明細件数:',sum(counts.values()),'通行料金合計:',sum(totals.values()))

wb=openpyxl.load_workbook(XLSX); ws=wb['高速料金']
YEN='"￥"#,##0'
used=set()
for r in range(2, ws.max_row+1):
    car=ws.cell(r,1).value
    c=ws.cell(r,3)
    if isinstance(car,int):
        c.value=totals.get(car,0); c.number_format=YEN; used.add(car)
        ws.cell(r,4).value=counts.get(car,0)
    else:
        c.value=None

# 明細にあるがリストに無い車番を追記
extra=sorted(set(totals)-used)
row=ws.max_row+1
for car in extra:
    ws.cell(row,1).value=car; ws.cell(row,2).value='（一覧に無い車番）'
    ws.cell(row,3).value=totals[car]; ws.cell(row,3).number_format=YEN
    ws.cell(row,4).value=counts[car]; row+=1

hdr=ws.cell(2,4); hdr2=ws.cell(2,2)
ws.cell(2,4).value=None
ws['D2']=None
tr=row+1
ws.cell(tr,2).value='合計'
ws.cell(tr,3).value=sum(totals.values()); ws.cell(tr,3).number_format=YEN
ws.cell(tr,4).value=sum(counts.values())
for col in (2,3,4):
    ws.cell(tr,col).font=Font(name='游ゴシック',sz=12,b=True)
for r in range(2, tr+1):
    for col in (1,2,3,4):
        cell=ws.cell(r,col)
        if cell.font.name is None: cell.font=Font(name='游ゴシック',sz=12)
ws.column_dimensions['C'].width=13
ws.column_dimensions['D'].width=9
wb.save(OUT)
print('未使用（明細に出現しない車番）:', sorted(c for c in used if c not in totals))
print('追記した車番:', extra)
print('saved', OUT)

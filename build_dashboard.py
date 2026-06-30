#!/usr/bin/env python3
import sys, os, json, datetime, openpyxl
from collections import defaultdict
def ni(x):
    return int(x) if isinstance(x,float) and x.is_integer() else x
XLSX = sys.argv[1] if len(sys.argv)>1 else 'live_sheet.xlsx'
OUT  = sys.argv[2] if len(sys.argv)>2 else 'index.html'
TPL  = os.path.join(os.path.dirname(os.path.abspath(__file__)),'template.html')
TAB  = '게시판제작주문진행현황의 (25년6월-)'
wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb[TAB] if TAB in wb.sheetnames else wb.active
wbs = openpyxl.load_workbook(XLSX)
wss = wbs[TAB] if TAB in wbs.sheetnames else wbs.active
def is_done(r):
    try:
        f=wss.cell(r,C_V).fill
        if not f or not f.patternType: return False
        rgb=f.fgColor.rgb
        if not isinstance(rgb,str) or len(rgb)<6: return False
        h=rgb[-6:]; R,G,B=int(h[0:2],16),int(h[2:4],16),int(h[4:6],16)
        return abs(R-G)<16 and abs(G-B)<16 and abs(R-B)<16 and R<=0xB0
    except Exception: return False
hdr={}
for c in range(1,ws.max_column+1):
    v=ws.cell(2,c).value
    if v: hdr[str(v).replace('\n','').strip()]=c
def col(*ns):
    for n in ns:
        if n in hdr: return hdr[n]
C_V,C_P,C_S,C_Q=col('업체'),col('상품'),col('사이즈'),col('수량')
C_D=col('출고일(디자이너)','출고일'); C_SH,C_T,C_A=col('배송'),col('연락처'),col('주소/송장','주소')
today=datetime.date.today()
months=[(today.year,today.month)]; y,m=today.year,today.month
for _ in range(2):
    m+=1
    if m>12: m=1;y+=1
    months.append((y,m))
def inwin(d): return (d.year,d.month) in months
agg=defaultdict(lambda:defaultdict(lambda:defaultdict(int))); detail=defaultdict(list); vdone={}
for r in range(4,ws.max_row+1):
    v=ws.cell(r,C_V).value; k=ws.cell(r,C_D).value
    if not v or not hasattr(k,'year'): continue
    d=k.date() if hasattr(k,'date') else k
    if not inwin(d): continue
    key=d.strftime('%Y-%m-%d')
    p=ws.cell(r,C_P).value or ''; sz=ws.cell(r,C_S).value or ''
    qv=ws.cell(r,C_Q).value; qty=qv if isinstance(qv,(int,float)) else (qv or 0)
    sh=(ws.cell(r,C_SH).value or '') if C_SH else ''
    tel=(ws.cell(r,C_T).value or '') if C_T else ''
    ad=(ws.cell(r,C_A).value or '') if C_A else ''
    done=is_done(r)
    agg[key][v][p]+= qty if isinstance(qty,(int,float)) else 0
    vdone.setdefault((key,v),[]).append(done)
    detail[key].append({'vendor':v,'product':p,'size':sz,'qty':ni(qty),'ship':sh,'tel':tel,'addr':ad,'done':done})
AGG={}
for dt in sorted(agg):
    AGG[dt]={'total':0,'vendors':[],'done':0,'all_done':False}
    for v in agg[dt]:
        prods=[{'name':p,'qty':ni(q)} for p,q in agg[dt][v].items()]
        vt=ni(sum(x['qty'] for x in prods)); AGG[dt]['total']+=vt
        flags=vdone.get((dt,v),[]); vdoneflag=bool(flags) and all(flags)
        if vdoneflag: AGG[dt]['done']+=vt
        AGG[dt]['vendors'].append({'vendor':v,'vtotal':vt,'products':prods,'done':vdoneflag})
    AGG[dt]['all_done']=AGG[dt]['done']>=AGG[dt]['total'] and AGG[dt]['total']>0
DET={k:detail[k] for k in sorted(detail)}
tpl=open(TPL,encoding='utf-8').read()
html=tpl.replace('__AGG__',json.dumps(AGG,ensure_ascii=False)).replace('__DET__',json.dumps(DET,ensure_ascii=False)).replace('__GEN__',today.strftime('%Y-%m-%d'))
open(OUT,'w',encoding='utf-8').write(html)
print('built',OUT,'| dates',len(AGG),'| today rows',len(DET.get(today.strftime('%Y-%m-%d'),[])))

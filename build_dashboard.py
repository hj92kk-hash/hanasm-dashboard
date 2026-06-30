#!/usr/bin/env python3
import sys, os, json, datetime, openpyxl
from collections import defaultdict
XLSX = sys.argv[1] if len(sys.argv)>1 else 'live_sheet.xlsx'
OUT  = sys.argv[2] if len(sys.argv)>2 else 'index.html'
TPL  = os.path.join(os.path.dirname(os.path.abspath(__file__)),'template.html')

# 소스 정의: (key, 화면라벨, 탭이름)
SOURCES_DEF = [
    ('board',  '게시판', '게시판제작주문진행현황의 (25년6월-)'),
    ('ottogi', '오뚜기', '오뚜기진행현황'),
]
def ni(x):
    return int(x) if isinstance(x,float) and x.is_integer() else x

today=datetime.date.today()
months=[(today.year,today.month)]; y,m=today.year,today.month
for _ in range(2):
    m+=1
    if m>12: m=1;y+=1
    months.append((y,m))
def inwin(d): return (d.year,d.month) in months

wb  = openpyxl.load_workbook(XLSX, data_only=True)
wbs = openpyxl.load_workbook(XLSX)

def norm(v): return ''.join(str(v).split()) if v is not None else ''

def build_source(tab):
    ws=wb[tab]; wss=wbs[tab]
    hdr={}
    for c in range(1,ws.max_column+1):
        k=norm(ws.cell(2,c).value)
        if k: hdr.setdefault(k,c)
    def col(*ns):
        for n in ns:
            if n in hdr: return hdr[n]
    C_V=col('업체'); C_P=col('상품'); C_S=col('사이즈')
    C_Q=col('수량','주문수량')
    C_D=col('출고일(디자이너)','출고일','고객납기일')
    C_SH=col('배송'); C_T=col('연락처'); C_A=col('주소/송장','주소')
    def is_done(r):
        try:
            f=wss.cell(r,C_V).fill
            if not f or not f.patternType: return False
            rgb=f.fgColor.rgb
            if not isinstance(rgb,str) or len(rgb)<6: return False
            h=rgb[-6:]; R,G,B=int(h[0:2],16),int(h[2:4],16),int(h[4:6],16)
            return abs(R-G)<16 and abs(G-B)<16 and abs(R-B)<16 and R<=0xB0
        except Exception: return False
    agg=defaultdict(lambda:defaultdict(lambda:defaultdict(int))); detail=defaultdict(list); vdone={}
    for r in range(4,ws.max_row+1):
        v=ws.cell(r,C_V).value if C_V else None
        k=ws.cell(r,C_D).value if C_D else None
        if not v or not hasattr(k,'year'): continue
        d=k.date() if hasattr(k,'date') else k
        if not inwin(d): continue
        key=d.strftime('%Y-%m-%d')
        p=(ws.cell(r,C_P).value or '') if C_P else ''
        sz=(ws.cell(r,C_S).value or '') if C_S else ''
        qv=ws.cell(r,C_Q).value if C_Q else 0
        qty=qv if isinstance(qv,(int,float)) else (qv or 0)
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
            flags=vdone.get((dt,v),[]); vf=bool(flags) and all(flags)
            if vf: AGG[dt]['done']+=vt
            AGG[dt]['vendors'].append({'vendor':v,'vtotal':vt,'products':prods,'done':vf})
        AGG[dt]['all_done']=AGG[dt]['done']>=AGG[dt]['total'] and AGG[dt]['total']>0
    DET={kk:detail[kk] for kk in sorted(detail)}
    return AGG,DET

SOURCES={}
order=[]
for key,label,tab in SOURCES_DEF:
    if tab in wb.sheetnames:
        a,d=build_source(tab)
        SOURCES[key]={'label':label,'agg':a,'det':d}
        order.append(key)
        print('source',key,label,'dates',len(a))

tpl=open(TPL,encoding='utf-8').read()
html=(tpl.replace('__SOURCES__',json.dumps(SOURCES,ensure_ascii=False))
         .replace('__ORDER__',json.dumps(order,ensure_ascii=False))
         .replace('__GEN__',today.strftime('%Y-%m-%d')))
open(OUT,'w',encoding='utf-8').write(html)
print('built',OUT)

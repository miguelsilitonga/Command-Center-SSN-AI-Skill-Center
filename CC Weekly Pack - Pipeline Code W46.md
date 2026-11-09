# CC Weekly Command Pack — pipeline code (W46 final, 9 Nov 2026 — interim file 6 + kiwi SP1 lines)

Working folder per week (e.g. scratchpad/w46b). Run order:
1. `load.py <interim.xlsx>` → s.pkl (BI sales), x.pkl (ScyllaX secondary), masters.pkl
2. `prep.py` → roster.pkl, target.pkl · `stockparse.py <SAP stock xlsx>` → stock.pkl · `stockgt.py` → stockgt9.pkl (GT Tangerang share per SKU: gt / gt_cond)
3. `engine.py` → D.pkl, jaga.pkl (signals, routes A/B) · `review.py` → rev45.pkl (last week's verdict inputs)
4. `w46data2.py` → P.pkl — THE WEEKLY FILE: verdicts, lines, Lampiran, supply commands, B/C sections, antre. Week-specific text lives here; rewrite it every week.
5. `price.py` (recommended price per line) → `risk.py` (A2 perish-risk + LA2 outlets) → `pricing.py`, `pricing2.py` (pricing deck data)
6. `build_sheet / build_review / build_wcr / build_pricing / build_snap / build_reg / build_v31 / build_wa` — refill last week's decks/workbooks (templates). Builders also hold hard-coded narrative: grep them for last week's numbers before building.
7. QA: `stripfonts.py` on the Review deck, `fillcheck.py *.pptx` (must be 0 bad tcPr), export PDFs with the pptx skill's soffice.py, render with pdftoppm and look at every page.
Template paths: common.py `A=` points at the folder holding last week's pack (attachments/uploads). tpl_sheet5.pptx = W45 sheet with the Lampiran slide duplicated (pptx skill add_slide.py).

## common.py
```python
import pickle, pandas as pd
from tfill import n0, n1
P=pickle.load(open('/tmp/claude-0/-home-claude/364a3463-b6f3-544b-bfc5-810208a0f40e/scratchpad/w46b/P.pkl','rb'))
A='/mnt/user-data/attachments/364a3463-b6f3-544b-bfc5-810208a0f40e/'
OUT='/tmp/claude-0/-home-claude/364a3463-b6f3-544b-bfc5-810208a0f40e/scratchpad/w46b/out/'
SHORT={'Achmad Farchan Syahrul':'Achmad Farchan','Bahar Ashary':'Bahar Ashary','Moreno Ananto':'Moreno Ananto','Ni DG Putri Mega Gita':'Putri Mega Gita','Reza Adiputranto':'Reza Adiputranto','Rizqi Akbar Nabawi':'Rizqi Akbar','Ryan Enrico':'Ryan Enrico','Sahrin Sahrin':'Sahrin','Bahar Ashary / Ryan Enrico (konfirmasi)':'Bahar / Ryan (konfirmasi)'}
def sid(i): return i.split('-')[-1]
def jaga(l): return 'Jaga' if 'Jaga' in l['rung'] else 'Perdalam'
def bnum(v):
    if v < 1: return '<1'
    return n0(v) if abs(v-round(v))<0.05 else n1(v)
def delta(b,w45,days=7):
    w=w45*7/days
    if b < 1: return f'{w/b:.0f}×' if w>0 else '—'
    if w>=2*b: return f'{w/b:.0f}×'
    return f'{(w/b-1)*100:+.0f}%'.replace('+-','-')
MON={10:'Okt',11:'Nov',9:'Sep',8:'Agu'}
def dmy(d): return f'{d.day} {MON[d.month]}'
```

## tfill.py
```python
"""Template refill helpers: keep W45 formatting, swap content."""
import copy
from pptx.util import Emu
from pptx.dml.color import RGBColor
from lxml import etree
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

def _set_runs(p, texts):
    runs = p.runs
    if not runs:
        r = p.add_run(); runs = p.runs
    for i, t in enumerate(texts):
        if i < len(runs): runs[i].text = t
        else:
            nr = copy.deepcopy(runs[-1]._r); runs[-1]._r.addnext(nr)
            runs = p.runs; runs[-1].text = t
    for r in p.runs[len(texts):]: r._r.getparent().remove(r._r)

def set_text(sh, paras):
    """paras: list of str | list[str] (runs). Reuses paragraph i's format; extra paras clone the last one."""
    tf = sh.text_frame if hasattr(sh, 'text_frame') else sh
    if isinstance(paras, str): paras = [paras]
    ps = tf.paragraphs
    for i, item in enumerate(paras):
        texts = ([item[:3], item[3:]] if item.startswith('•  ') else [item]) if isinstance(item, str) else list(item)
        if i < len(ps): p = ps[i]
        else:
            src = ps[-1]._p if len(ps) > 1 else ps[0]._p
            np_ = copy.deepcopy(src); tf._txBody.append(np_)
            p = tf.paragraphs[-1]
        _set_runs(p, texts)
        ps = tf.paragraphs
    for p in tf.paragraphs[len(paras):]: p._p.getparent().remove(p._p)

def find(slide, startswith):
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip().startswith(startswith): return sh
    raise KeyError(startswith)

def cell(c, text, bold=None, color=None, fill=None, size=None, align=None):
    tf = c.text_frame
    ps = tf.paragraphs
    for p in ps[1:]: p._p.getparent().remove(p._p)
    p = tf.paragraphs[0]
    lines = str(text).split('\n')
    _set_runs(p, [lines[0]])
    for ln in lines[1:]:
        np_ = copy.deepcopy(p._p); tf._txBody.append(np_); _set_runs(tf.paragraphs[-1], [ln])
    for p in tf.paragraphs:
        for r in p.runs:
            if bold is not None: r.font.bold = bold
            if color is not None: r.font.color.rgb = RGBColor.from_string(color)
            if size is not None: r.font.size = size
        if align is not None: p.alignment = align
    if fill is not None: set_fill(c, fill)

def set_fill(c, hexcol):
    """Single valid fill: drop any existing fill (noFill/solidFill/...), insert solidFill after borders/cell3D."""
    tcPr = c._tc.get_or_add_tcPr()
    for e in list(tcPr):
        if e.tag.split('}')[1] in ('noFill','solidFill','gradFill','blipFill','pattFill','grpFill'): tcPr.remove(e)
    sf = etree.Element(A + 'solidFill'); clr = etree.SubElement(sf, A + 'srgbClr'); clr.set('val', hexcol)
    before = [e for e in tcPr if e.tag.split('}')[1] in ('lnL','lnR','lnT','lnB','lnTlToBr','lnBlToTr','cell3D')]
    if before: before[-1].addnext(sf)
    else: tcPr.insert(0, sf)

def get_fill(c):
    f = c._tc.get_or_add_tcPr().find(A + 'solidFill/' + A + 'srgbClr')
    return f.get('val') if f is not None else None

def resize_rows(tbl, n_data, header=1, tmpl=None):
    """Make table have header + n_data rows. New rows clone row `tmpl` (default first data row)."""
    trs = tbl._tbl.tr_lst
    tmpl = header if tmpl is None else tmpl
    src = copy.deepcopy(trs[min(tmpl, len(trs) - 1)])
    cur = len(trs) - header
    while cur > n_data:
        tbl._tbl.remove(tbl._tbl.tr_lst[-1]); cur -= 1
    while cur < n_data:
        tbl._tbl.tr_lst[-1].addnext(copy.deepcopy(src)); cur += 1

def fill_table(tbl, rows, header=1, styles=None):
    """rows: list of lists. styles: optional fn(r_idx, c_idx) -> dict for cell()."""
    resize_rows(tbl, len(rows), header)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            kw = styles(i, j, row) if styles else {}
            cell(tbl.cell(header + i, j), '' if v is None else v, **(kw or {}))

def table_height(tbl):
    return sum(r.height for r in tbl.rows)

def set_row_h(tbl, h_emu, start=1):
    for r in list(tbl.rows)[start:]: r.height = h_emu

def shift(shapes, dy):
    for s in shapes: s.top = s.top + int(dy)

def chart_data(chart, cats, series):
    from pptx.chart.data import CategoryChartData
    cd = CategoryChartData(); cd.categories = cats
    for name, vals in series: cd.add_series(name, vals)
    chart.replace_data(cd)

def delete_slide(prs, idx):
    sldIdLst = prs.slides._sldIdLst; sid = sldIdLst[idx]
    prs.part.drop_rel(sid.rId); sldIdLst.remove(sid)

def fmt_id(c):  # pptx number formatting helper
    return c

def n0(v):  # Indonesian thousand separator
    return f'{v:,.0f}'.replace(',', '.')
def n1(v):
    return f'{v:,.1f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
def jt(v):  # rupiah to 'Rp x,y jt'
    return n1(v / 1e6)
def M(v):
    return n1(v / 1e9) + ' M'
def pct(v):
    return f'{v*100:.0f}%'

import math
def _cell_pt(c):
    for p in c.text_frame.paragraphs:
        for r in p.runs:
            if r.font.size: return r.font.size.pt
    return 7.0
def est_heights(tbl, k=0.56, minh=0.17, lead=1.22):
    """Estimate rendered row heights (inches) from text length; set them; return total EMU."""
    widths=[c.width/914400 for c in tbl.columns]
    tot=0
    for ri,row in enumerate(tbl.rows):
        need=minh
        for ci,cel in enumerate(row.cells):
            pt=_cell_pt(cel); w=max(0.2,widths[ci]-0.09)
            lines=0
            for p in cel.text_frame.paragraphs:
                txt=''.join(r.text for r in p.runs)
                lines+=max(1,math.ceil(len(txt)*k*pt/72/w))
            need=max(need, lines*pt*lead/72+0.06)
        row.height=int(need*914400); tot+=row.height
    return tot
def below(shape, ref, gap=0.06):
    h = ref.height if not getattr(ref,'has_table',False) else sum(r.height for r in ref.table.rows)
    shape.top = ref.top + h + int(gap*914400)
def color_points(chart, base='0B5A2A', last='E8A317'):
    from pptx.dml.color import RGBColor
    ser=chart.plots[0].series[0]; n=len(list(ser.values))
    for i in range(n):
        pt=ser.points[i]; pt.format.fill.solid(); pt.format.fill.fore_color.rgb=RGBColor.from_string(last if i==n-1 else base)

def add_col(tbl, after, width_emu):
    """Insert a column after index `after`, cloning that column's cells/format."""
    grid = tbl._tbl.tblGrid; gcs = grid.findall(A + 'gridCol')
    ng = copy.deepcopy(gcs[after]); ng.set('w', str(int(width_emu))); gcs[after].addnext(ng)
    for tr in tbl._tbl.tr_lst:
        tcs = tr.findall(A + 'tc'); nt = copy.deepcopy(tcs[after]); tcs[after].addnext(nt)
def del_cols(tbl, idxs):
    grid = tbl._tbl.tblGrid; gcs = grid.findall(A + 'gridCol')
    for i in sorted(idxs, reverse=True):
        grid.remove(gcs[i])
        for tr in tbl._tbl.tr_lst: tr.remove(tr.findall(A + 'tc')[i])
def set_widths(tbl, inches):
    for c, w in zip(tbl.columns, inches): c.width = int(w * 914400)
```

## load.py
```python
"""Load interim dashboard (latest) → s.pkl (BI sales 2026, enriched), x.pkl (ScyllaX secondary), masters.pkl"""
import sys, pickle, pandas as pd, numpy as np
from python_calamine import CalamineWorkbook
F = sys.argv[1] if len(sys.argv) > 1 else '/mnt/user-data/uploads/CC_Interim_Dashboard_v1_4_2__6_.xlsx'
wb = CalamineWorkbook.from_path(F)
def sheet(n):
    r = wb.get_sheet_by_name(n).to_python(); return pd.DataFrame(r[1:], columns=[str(c) for c in r[0]])
mc = sheet('MASTER_CUSTOMER'); msku = sheet('MASTER_SKU'); msub = sheet('MASTER_SUBDIST'); mo = sheet('MASTER_OUTLET')
s = sheet('RAW_SALES_2026')
s = s.rename(columns={'View Dim Time All NEW': 'dt', 'Sales Office': 'br', 'Channel Desc': 'ch', 'View Dim Customer New': 'code', 'ITEMID': 'itemid', 'PRODH2': 'prodh2'})
s['dt'] = pd.to_datetime(s.dt, errors='coerce'); s = s[s.dt.notna()]
for c in ['Net Sales BE', 'Net Sales Amt', 'So BE', 'Delivery BE', 'Cancel So BE', 'Cancel So Amt', 'Returned BE', 'Returned Amt']: s[c] = pd.to_numeric(s[c], errors='coerce').fillna(0.0)
s['code'] = s.code.astype(str).str.strip()
cn = mc.drop_duplicates('View Dim Customer New').set_index('View Dim Customer New')
s['cname'] = s.code.map(cn['Customer Account']).fillna(s.code)
s['fmt'] = s.code.map(cn['Format Desc']).fillna('')
s['sku'] = s.itemid.map(msku.drop_duplicates('ITEMID').set_index('ITEMID')['ITEMNAME']).fillna(s.itemid)
K = {'Banana': 'C', 'Pineapple': 'P', 'Guava': 'G'}
s['kat'] = s.prodh2.map(K)
loc = s.itemid.str.startswith('FF-') & s.kat.isna()
# Local vs Import: local fruit codes (melon, papaya, watermelon, orange local, mango, etc.) — PRODH2 tells group
print(sorted(s.prodh2.dropna().unique()))
s.to_pickle('s_raw.pkl')
x = sheet('RAW_SECONDARY')
x = x.rename(columns={'DistributorCode': 'dist', 'DocumentDate': 'dt'})
x['dt'] = pd.to_datetime(x.dt, errors='coerce'); x['dist'] = x.dist.astype(str).str.strip()
for c in ['Qty3Final', 'Qty2Final', 'ConvUnit3', 'GrossSales']: x[c] = pd.to_numeric(x[c], errors='coerce').fillna(0.0)
x['box'] = x.Qty3Final + np.where(x.ConvUnit3 > 0, x.Qty2Final / x.ConvUnit3.replace(0, np.nan), 0).astype(float)
x['box'] = x.box.fillna(x.Qty3Final)
print(x.columns.tolist()[30:])
x.to_pickle('x.pkl'); pickle.dump(dict(mc=mc, msku=msku, msub=msub, mo=mo), open('masters.pkl', 'wb'))
print(s.dt.min(), s.dt.max(), len(s), '| x', x.dt.min(), x.dt.max(), len(x))
```

## prep.py
```python
import pandas as pd, numpy as np
LOCAL = {'Melon', 'Papaya', 'Watermelon', 'Mango', 'Durian', 'Salacca', 'Jack Fruit', 'Coconut', 'Naga', 'Passion Fruit', 'Avocado'}
FMCG = {'Chips', 'Cut Fruit', 'Fresh Milk', 'Fruit Fillings', 'Juice', 'Yogurt Drink', 'Yogurt Spoonable', 'Yogurt Squeeze', 'Preserved Fruits', 'Mix Fruit', 'Unknown'}
def kat(p):
    if p == 'Banana': return 'C'
    if p == 'Pineapple': return 'P'
    if p == 'Guava': return 'G'
    if p in LOCAL: return 'L'
    if p in FMCG: return 'F'
    return 'I'
s = pd.read_pickle('s_raw.pkl')
u = s.prodh2.drop_duplicates(); m = dict(zip(u, u.map(kat))); s['kat'] = s.prodh2.map(m)
s['k2'] = np.where(s.prodh2 == 'Melon', 'M', s.kat)
s = s[s.kat != 'F']
s.to_pickle('s.pkl')
x = pd.read_pickle('x.pkl'); pc = x.ProCode.astype(str)
x['kat'] = np.select([pc.str.startswith('FF-01'), pc.str.startswith('FF-03'), pc.str.startswith('FF-02'), pc.str.startswith('FF-')], ['C', 'P', 'G', 'I'], 'F')
x.to_pickle('x.pkl')
g = s[(s.br == 'Tangerang') & (s.ch == 'General Trade')]
w = g[g.dt >= '2026-10-26']; print(w.groupby(w.dt.dt.date)[['So BE', 'Delivery BE', 'Net Sales BE']].sum().round(0).to_string())
print(x.groupby('dist').dt.max().value_counts().sort_index().tail(8))
```

## stockparse.py
```python
"""Parse ZSAP Stock Sales Automatic Report (All Tangerang): banana / import / local sheets → one table (desc, banner, qty, aged4)."""
import re, pandas as pd
from python_calamine import CalamineWorkbook
def _num(v):
    try: return float(v)
    except (TypeError, ValueError): return 0.0
def parse(path):
    wb = CalamineWorkbook.from_path(path); out = []; date = None
    for sn in wb.sheet_names:
        rows = wb.get_sheet_by_name(sn).to_python()
        if not rows: continue
        if date is None and len(rows[0]) > 2: date = rows[0][2]
        banner = next((str(r[0]).strip() for r in rows[:6] if str(r[0]).strip() and 'Stock Sales' not in str(r[0]) and 'Dokumen' not in str(r[0])), '')
        if banner not in ('BANANA 1S', 'IMPORT', 'LOCAL'): continue
        hi = next(i for i, r in enumerate(rows) if str(r[0]).startswith(('Product Short', 'Mtrl Group', 'Product Description')))
        hdr = [str(h).strip() for h in rows[hi]]
        age_cols = [j for j, h in enumerate(hdr) if j >= 6 and (h == 'Hari ini' or re.fullmatch(r'\d+', h) or h == '>14')]
        tot = next((j for j, h in enumerate(hdr) if h.lower() == 'total'), None)
        for r in rows[hi + 1:]:
            d = str(r[0]).strip()
            if not d: continue
            ages = {hdr[j]: _num(r[j]) for j in age_cols if j < len(r)}
            age_total = sum(ages.values())
            qty = _num(r[tot]) if tot is not None and tot < len(r) else age_total
            aged4 = sum(v for h, v in ages.items() if h == '>14' or (h.isdigit() and int(h) >= 4))
            out.append(dict(desc=d, banner=banner, qty=qty, aged4=aged4, uom=r[5] if len(r) > 5 else ''))
    T = pd.DataFrame(out); T.attrs['date'] = date
    return T
def stk(T, pat): return float(T[T.desc.str.contains(pat, case=False, regex=True)].qty.sum())
def aged(T, pat): return float(T[T.desc.str.contains(pat, case=False, regex=True)].aged4.sum())
```

## stockgt.py
```python
"""GT Tangerang-eligible stock (Migs 8-9 Nov): ORI / fresh 0-3 days -> MT & other GT branches; any grade aged >=4 days, and SP2/SP3, -> GT TGR.
SP1 is GT TGR only when aged >3 days. Import sheet has no age: gt = SP2+SP3, gt_cond = SP1 (confirm age at warehouse)."""
import re, pandas as pd
from python_calamine import CalamineWorkbook
def num(v):
    try: return float(v)
    except: return 0.0
def parse(path):
    wb=CalamineWorkbook.from_path(path); out=[]
    for sn in wb.sheet_names[:3]:
        rows=wb.get_sheet_by_name(sn).to_python()
        hi=next(i for i,r in enumerate(rows) if str(r[0]).startswith(('Product Short','Mtrl Group')))
        h=[str(x).strip() for x in rows[hi]]; ix={k:j for j,k in enumerate(h) if k}
        ages=[j for j,k in enumerate(h) if j>=6 and (k=='Hari ini' or re.fullmatch(r'\d+',k) or k=='>14')]
        for r in rows[hi+1:]:
            d=str(r[0]).strip()
            if not d: continue
            g=lambda k: num(r[ix[k]]) if k in ix and ix[k]<len(r) else 0.0
            a4=sum(num(r[j]) for j in ages if h[j]=='>14' or (h[j].isdigit() and int(h[j])>=4))
            fresh=sum(num(r[j]) for j in ages if h[j]=='Hari ini' or (h[j].isdigit() and int(h[j])<=3))
            tot=g('Total') if 'Total' in ix else a4+fresh
            ori=g('ORI'); sp1=g('SP1') if 'SP1' in ix else g('SP'); sp2=g('SP2'); sp3=g('SP3')
            if sn==wb.sheet_names[0]: gt=a4                      # banana: age only
            elif 'SP1' in ix: gt=a4+sp2+sp3                      # local: age columns track ORI/SP1
            else: gt=sp2+sp3                                     # import: grade only (SP1 conditional)
            out.append(dict(desc=d,sheet=sn,total=tot,fresh=fresh,aged4=a4,ori=ori,sp1=sp1,sp2=sp2,sp3=sp3,gt=gt,gt_cond=(sp1 if ('SP1' not in ix and sn!=wb.sheet_names[0]) else 0.0)))
    return pd.DataFrame(out)
```

## engine.py
```python
import pandas as pd, numpy as np, stockparse as sp
pd.set_option('display.width',250); pd.set_option('display.max_columns',30)
s=pd.read_pickle('s.pkl'); x=pd.read_pickle('x.pkl'); R=pd.read_pickle('roster.pkl')
g=s[(s.br=='Tangerang')&(s.ch=='General Trade')].copy()
T=sp.parse('/mnt/user-data/attachments/364a3463-b6f3-544b-bfc5-810208a0f40e/016d8c57-ZSAP_Stock_Sales_Automatic_Report_All_Tangerang_2.xlsx')
def sales(code,name):
    r=R[R.code==code]
    if len(r): return r.sales.iloc[0]
    n=name.lower()
    if 'kebun buah' in n: return 'Rizqi Akbar Nabawi'
    if 'kopkar' in n or 'koperasi' in n: return 'Bahar Ashary / Ryan Enrico (konfirmasi)'
    return None
WK=(g.dt>='2026-11-02')&(g.dt<='2026-11-08'); B4=(g.dt>='2026-10-05')&(g.dt<='2026-11-01'); FL=(g.dt>='2026-11-02')&(g.dt<='2026-11-05')
key=['code','cname','fmt','k2']
wk=g[WK].groupby(key)['So BE'].sum().rename('wk')
b4=(g[B4].groupby(key)['So BE'].sum()/4).rename('b4')
fl=g[FL].groupby(key)[['So BE','Delivery BE']].sum(); fl=(fl['Delivery BE']/fl['So BE']).rename('fill')
rp=g[B4|FL].groupby(key).apply(lambda d: d['Net Sales Amt'].sum()/max(d['Net Sales BE'].sum(),1e-9)).rename('rp')
top=g[B4].groupby(key+['sku'])['So BE'].sum().reset_index().sort_values('So BE').groupby(key).tail(1).set_index(key)['sku'].rename('topsku')
D=pd.concat([b4,wk,fl,rp,top],axis=1).fillna({'wk':0,'b4':0}).reset_index()
brfill=g[FL].groupby('k2')[['So BE','Delivery BE']].sum(); brfill=brfill['Delivery BE']/brfill['So BE']
D['brfill']=D.k2.map(brfill)
J=D[(D.b4>=10)&(D.wk<0.7*D.b4)].copy()
J['gap']=J.b4-J.wk; J['val']=J.gap*J.rp*4.3/1e6; J['score']=J.val*0.6
J=J[J.score>=0.3]
J['sales']=[sales(c,n) for c,n in zip(J.code,J.cname)]
def route(r):
    if r.k2=='G': return 'B: Guava'
    if r.brfill<0.6: return f'B: fill cabang {r.brfill:.0%}'
    if pd.notna(r.fill) and r.fill<0.9: return f'B: fill cust {r.fill:.0%}'
    return 'A'
J['route']=J.apply(route,axis=1)
J=J.sort_values('score',ascending=False)
print(brfill.round(2).to_dict())
print(J[['code','cname','fmt','k2','b4','wk','fill','rp','val','score','topsku','sales','route']].head(45).round(2).to_string())
J.to_pickle('jaga.pkl'); D.to_pickle('D.pkl'); T.to_pickle('stock.pkl')
```

## review.py
```python
import pandas as pd, numpy as np, json
s=pd.read_pickle('s.pkl'); x=pd.read_pickle('x.pkl')
g=s[(s.br=='Tangerang')&(s.ch=='General Trade')].copy()
D=lambda a,b:(g.dt>=a)&(g.dt<=b)
CMD=[
 dict(id='CC-2610-129',cust='0003100403',dist='3100403',pat=r'Pine',k='P',target=115,typ='Jaga',L=['FF260073','FF260069','FF260044','FF260058','FF260088'],xp='Pine'),
 dict(id='CC-2610-134',cust='0003101297',pat=r'Honey Emerald',k='M',target=150,typ='Jaga'),
 dict(id='CC-2610-138',cust='0003101242',dist='3101242',pat=r'Pine',k='P',target=410,typ='Jaga',L=['RNM0120','RNM0002','RNM0128','RNM0008','RNM0157'],xp='Pine'),
 dict(id='CC-2610-139',cust='0003104911',pat=r'Honey Emerald',k='M',target=110,typ='Jaga'),
 dict(id='CC-2610-136',cust='0003101202',dist='3101202',pat=r'Honey Emerald',k='M',target=40,typ='Jaga',L=['RSN260002','RSN260003','RSN062','RSN260010','RSN003'],xp='Melon'),
 dict(id='CC-2610-147',cust='0003100518',dist='3100518',pat=r'Honey Emerald',k='M',target=90,typ='Perdalam',L=['HRM260033','HRM0038','HRM260005','HRM0045','HRM260070','HRM260044','HRM0039','HRM0149'],xp='Melon'),
 dict(id='CC-2610-145',cust='0003100295',pat=r'Honey Emerald',k='M',target=60,typ='Perdalam'),
 dict(id='CC-2610-135',cust='0003101297',pat=r'Kiwi',k='I',target=120,typ='Perdalam'),
 dict(id='CC-2610-146',cust='0003101607',pat=r'Honey Emerald',k='M',target=30,typ='Perdalam'),
 dict(id='CC-2610-140',cust='0003101242',dist='3101242',pat=r'Dazzle',k='I',target=40,typ='Perdalam',L=['SSN260002','RNM0197','RNM0065','RNM0032','RNM0025','RNM0122','RNM0031','RNM0121'],xp='Dazzle'),
]
res=[]
for c in CMD:
    a=g[(g.code==c['cust'])&g.sku.str.contains(c['pat'],case=False)]
    cat=g[(g.code==c['cust'])&(g.k2==c['k'])]
    wk=a[(a.dt>='2026-11-02')&(a.dt<='2026-11-08')]['So BE'].sum()
    wkc=cat[(cat.dt>='2026-11-02')&(cat.dt<='2026-11-08')]['So BE'].sum()
    w44=a[(a.dt>='2026-10-26')&(a.dt<='2026-11-01')]['So BE'].sum()
    b8=a[(a.dt>='2026-09-07')&(a.dt<='2026-11-01')]['So BE'].sum()/8
    b8c=cat[(cat.dt>='2026-09-07')&(cat.dt<='2026-11-01')]['So BE'].sum()/8
    r=a[(a.dt>='2026-10-05')&(a.dt<='2026-11-08')]; rp=r['Net Sales Amt'].sum()/max(r['Net Sales BE'].sum(),1e-9)
    if r['Net Sales BE'].sum()<1:
        r=g[g.sku.str.contains(c['pat'],case=False)&(g.dt>='2026-10-05')]; rp=r['Net Sales Amt'].sum()/r['Net Sales BE'].sum()
    dd=a[(a.dt>='2026-11-02')].groupby(a.dt.dt.strftime('%d'))['So BE'].sum().to_dict()
    d=dict(id=c['id'],cust=a.cname.iloc[0] if len(a) else c['cust'],typ=c['typ'],target=c['target'],wk=wk,wk_cat=wkc,w44=w44,b8=round(b8,1),b8c=round(b8c,1),rp=round(rp),days=dd)
    if 'dist' in c:
        xx=x[(x.dist.astype(str)==c['dist'])&(x.dt>='2026-11-02')&x.ProName.str.contains(c['xp'],case=False)&(x['trx type']=='ORDER')]
        d['so_box']=round(xx.box.sum(),1); d['so_rp']=round(xx.NetSalesExcPPN.sum())
        d['L_hit']=sorted(set(xx.OutletCode)&set(c['L'])); d['L_n']=len(c['L'])
        d['L_box']=round(xx[xx.OutletCode.isin(c['L'])].box.sum(),1)
        d['xmax']=str(x[x.dist.astype(str)==c['dist']].dt.max().date())
    res.append(d)
for d in res: print(d)
pd.to_pickle(res,'rev45.pkl')
```

## w46data2.py
```python
import pandas as pd, numpy as np, pickle
from tfill import n0, n1, jt
s=pd.read_pickle('s.pkl'); x=pd.read_pickle('x.pkl'); x=x[x['trx type']=='ORDER'].copy(); x['d']=x.dist.astype(str)
D=pd.read_pickle('D.pkl'); J=pd.read_pickle('jaga.pkl'); T=pd.read_pickle('stock.pkl'); rev=pd.read_pickle('rev45.pkl')
g=s[(s.br=='Tangerang')&(s.ch=='General Trade')].copy()
P={}
P['meta']=dict(week=46, wk_range='9–15 November 2026', final='Senin 9 Nov', issue='Selasa 10 Nov', due='Senin 16 Nov 09:00',
  data='order W45 s/d Min 8 Nov · terkirim andal s/d 1 Nov · ScyllaX s/d 5 Nov · stok SAP Min 8 Nov · S&D 30 Okt')
def dget(code,k2,col):
    r=D[(D.code==code)&(D.k2==k2)]; return float(r[col].iloc[0]) if len(r) else 0.0
# ---------------- W45 REVIEW ----------------
R={r['id']:r for r in rev}
V=[ # id, short target, result text, verdict, W46 decision, note
 ('CC-2610-129','Fresh Fruit','Pine','115 / mgg', 'Pine', 155, "155 BE · sell-out 53 box (2–8 Nov) · L1 3/5 outlet", 'Met','Tutup (Met)','order 6 Nov 120 BE (terbooking setelah 8 Nov)'),
 ('CC-2610-134','Sastro Fruits & Vegetables','HE','150 / mgg','Melon',119, "HE 119 BE · melon total 320 BE (Golden 146, Korean 30)", 'Partial','Tutup — melon pulih (320 vs ±99)','kategori melon 3× biasa; HE 79% target'),
 ('CC-2610-138','Ranum Buah','Pine','410 / mgg','Pine',440, "440 BE · sell-out 500 box · L2 0/5 outlet", 'Met','Tutup (Met)','order pulih; outlet Lampiran belum kembali'),
 ('CC-2610-139','Koperasi KopkarFresh Cikpo','HE','110 / mgg','Melon',38, "HE 38 BE · melon 93 BE (Rock 23, Golden 30)", 'Partial','Tutup — melon pulih (93 vs ±100)','kategori melon kembali; HE 35% target'),
 ('CC-2610-136','Raksi Segar Nusantara','HE','40 / mgg','Melon',0, "HE 0 · Rock C 20 BE · sell-out 11 box (2–8 Nov) · L3 3/5", 'Partial','Lanjut 1× (melon SP2 / > 3 hr) → L1','sell-out naik (4 → 11 box); stok subdist ±8 box'),
 ('CC-2610-147','Harum Buah Segar','HE','90 / mgg','Melon',0, "HE 0 · Rock C 5 BE · sell-out 3 box · L4 1/8", 'Partial','Tutup — sell-out outlet macet','sell-out melon 2–8 Nov 3 box; stok subdist ±15'),
 ('CC-2610-145','Duta Buah','HE','60 / mgg','Melon',170, "HE 170 BE (3, 4 & 8 Nov)", 'Met','Tutup (Met) → baris baru 006','pembeli HE baru: 60 → 170 BE dalam 2 minggu'),
 ('CC-2610-135','Sastro Fruits & Vegetables','Kiwi','120 / mgg','Import',147, "kiwi 147 BE (SZ27 60 · Pnt 48 · SZ30 22 · Green 17)", 'Met','Tutup (Met)','order 7–9 Nov masuk setelah draft'),
 ('CC-2610-146','Yitno Jaya - Kramat Jati','HE','30 / mgg','Melon',0, "0 BE melon", 'Failed','Tutup','pembeli HE W44 (20 BE) tidak kembali'),
 ('CC-2610-140','Ranum Buah','Dazzle','40 BE (sekali)','Import',0, "0 BE · Ranum jual Dazzle 25 box dari stok sendiri · L5 2/8", 'Failed','Tutup — Dazzle 112 box, habis W47 lewat MT/online','stok TGR 264 → 112 box (MT & online)'),
]
P['rev']=[]
for vid,cust,prod,tgt,kat,res,txt,verd,dec,note in V:
    r=R[vid]; rp=r['rp']; wk=r['wk'] if vid not in('CC-2610-136','CC-2610-147') else r['wk_cat']
    val=wk*rp; inc=max(0,wk-r['b8'])*rp if vid not in('CC-2610-136','CC-2610-147') else max(0,wk-r['b8c'])*rp
    P['rev'].append(dict(id=vid,cust=cust,prod=prod,kat=kat,target=tgt,res=txt,verdict=verd,dec=dec,note=note,wk=wk,b8=(r['b8'] if vid not in('CC-2610-136','CC-2610-147') else r['b8c']),w44=r['w44'],rp=rp,val=val,inc=inc,so_box=r.get('so_box'),so_rp=r.get('so_rp')))
cnt=pd.Series([v['verdict'] for v in P['rev']]).value_counts().to_dict(); P['cnt45']=cnt
P['inc45']=sum(v['inc'] for v in P['rev'] if v['verdict'] in('Met','Partial'))
P['so45']=sum((v['so_rp'] or 0) for v in P['rev'] if v['verdict'] in('Met','Partial'))
P['revsup']=[
 ('CC-2610-141','PM Local (melon) + MT','Stok HE TGR ≤ 1 mgg cover s/d Jum 6 Nov', 'Stok HE 8 Nov 591 box vs order HE TGR W45 > 800 BE (GT 607) = < 1 mgg cover; harga HE SP GT ±319 rb (di atas pita 275–300 rb) tetapi order GT naik 468 → 607 BE','Met','Tutup → 007 (alokasi HE W46)'),
 ('CC-2610-142','Supply Planning Import + PM Import','Tanggal lot Century + harga tertulis s/d Rab 4 Nov','Tidak ada tanggal tertulis per 9 Nov; DS Century 7 Nov belum tersedia (S&D 30 Okt); stok Century 0','Failed','Terbit ulang → 008 (DS Century)'),
 ('CC-2610-143','PM Import + BM Tangerang','Kiwi SZ30+SZ27 keluar ≥ 600 box/mgg; Setia ≥ 1 order','Order TGR SZ30+SZ27 W45 971 BE (MT 729); Setia 15 BE (5 Nov); Sastro kiwi 147 BE (Met)','Met','Tutup → 010 (laju kiwi W46)'),
 ('CC-2610-144','PM Import','Stok Dazzle TGR < 150 box s/d 6 Nov','Stok 8 Nov 112 box (dari 264); GT hanya 2 BE — keluar lewat MT & online','Met','Tutup — sisa 112 box habis W47 (S&D)'),
 ('CC-2610-116','Supply Planning + BM Tangerang','Porsi GT tertulis s/d 4 Nov; fill GT ≥ 50%','Banana GT terkirim 2–5 Nov 38% (26 Okt–1 Nov 39%); porsi tertulis tidak ada di data','Failed','Ditahan W46 — Cavendish di luar command'),
]
# ---------------- STOCK helpers ----------------
def st(p): return float(T[T.desc.str.contains(p,case=False,regex=True)].qty.sum())
def ag(p): return float(T[T.desc.str.contains(p,case=False,regex=True)].aged4.sum())
STK=dict(he=st('Honey Emerald'),he_sp=st('Honey Emerald Melon GH SP 15'),gold=st('Golden Melon'),rock=st('Rock Melon'),
 kiwi30=st('Sungold Kiwi 5,5 SZ 30'),kiwi27=st('Sungold Kiwi 5,5 SZ 27'),green30=st('Green Kiwi 9,5 SZ 30'),singo=st('Singo Pear'),gpear=st('Golden Pear'),
 dazzle=st('Dazzle'),cl8=st('Honi Pine SPA? ?CL 8'),cl10=st('Honi Pine SPA? ?CL 10'),cl15=st('CL 15'),cr9=st('CR 9'),cr7=st('CR 7'),cr5=st('CR 5'),
 c313=st('C3 13KG 1S$'),c35=st('C3 5KG'),sml=st('Sml Hand 13KG'),fb=st('Bn FB 13KG|Bn FB1 13KG'),fk=st('Bn FK 13KG'),
 guava=st('Crystal Guava'),lemon=st('Lemon'))
STK['cl8']=st('Honi Pine SP CL 8')+st('Honi Pine SPA CL 8'); STK['cl10']=st('Honi Pine SP CL 10')+st('Honi Pine SPA CL 10')
STK['c313_ag']=ag('C3 13KG 1S$'); STK['cr9_ag']=ag('CR 9'); STK['cl8_ag']=ag('Honi Pine SPA? ?CL 8'); STK['he_ag']=ag('Honey Emerald'); STK['fk_ag']=ag('Bn FK 13KG')
P['stk']=STK; print(STK)
# ---------------- LAMPIRAN ----------------
def outl(dist,pat,top=8,mode='drop',tgt=None):
    a=x[(x.d==dist)&~x.OutletName.astype(str).str.upper().str.contains('SHOP SALES')&x.ProName.str.contains(pat,case=False,regex=True)&~x.ProName.str.contains('Juice|Milk|Yog|Chips|Cut|Fill',case=False)]
    k=['OutletCode','OutletName','OtTypeName']
    b=a[(a.dt>='2026-10-05')&(a.dt<='2026-11-01')].groupby(k).box.sum()/4
    w=a[a.dt>='2026-11-02'].groupby(k).box.sum()
    last=a.groupby(k).dt.max()
    sku=a[a.dt>='2026-10-05'].groupby(k+['ProName']).box.sum().reset_index().sort_values('box').groupby(k).tail(1).set_index(k).ProName
    Q=pd.concat([b.rename('biasa'),w.rename('w45'),last.rename('akhir'),sku.rename('sku')],axis=1).fillna({'biasa':0,'w45':0}).reset_index()
    Q=Q[Q.biasa>0]; Q['gap']=Q.biasa-Q.w45
    Q=Q.sort_values('gap' if mode=='drop' else 'biasa',ascending=False).head(top)
    Q['target']=Q.biasa.round().clip(lower=1) if tgt is None else Q.biasa.apply(tgt)
    return Q
def short_sku(p):
    p=p.replace('Sunpride ','').replace('Zespri ','')
    for a,b in [('Sungold Kiwi 5,5KG SZ ','SG SZ'),('Green Kiwi 9,5KG SZ ','Green SZ'),('Sungold Kiwi Pnt ','SG Pnt '),('Green Kiwi Pnt ','Green Pnt '),('Honey Emerald Melon GH','HE'),('Honi Pine','Pine'),(' 15KG',''),('PCS',''),('Red Globe Grape','Red Globe'),('Mandarin Shantang','Shantang')]: p=p.replace(a,b)
    return p[:22]
L=[]
q=outl('3101202','Melon',mode='buyers'); L.append(dict(code='L1',id='CC-2610-136',cust='RAKSI SEGAR NUSANTARA',desc='outlet pembeli melon — tawarkan melon SP2 / SP1 > 3 hr (Golden, HE)',q=q))
q=outl('3101242','Melon',mode='buyers'); L.append(dict(code='L2',id='CC-2611-001',cust='RANUM BUAH',desc='outlet pembeli melon Okt — tawarkan melon SP2 / SP1 > 3 hr (Golden, HE, Korean)',q=q))
# ---- kiwi Lampiran (13 mgg pembeli kiwi; SP1 kiwi boleh ke GT TGR karena stok kiwi lebih)
import math
def outk(dist,top=8):
    a=x[(x.d==dist)&x.ProName.str.contains('Kiwi',case=False)&~x.OutletName.astype(str).str.upper().str.contains('SHOP SALES|CUSTOMER GUDANG|CUSTOMER ONLINE')]
    k=['OutletCode','OutletName','OtTypeName']; a=a.assign(OtTypeName=a.OtTypeName.fillna(''))
    b=a[a.dt>='2026-08-10'].groupby(k).box.sum()
    w=a[a.dt>='2026-11-02'].groupby(k).box.sum(); last=a.groupby(k).dt.max()
    sku=a[a.dt>='2026-08-10'].groupby(k+['ProName']).box.sum().reset_index().sort_values('box').groupby(k).tail(1).set_index(k).ProName
    Q=pd.concat([b.rename('b13'),w.rename('w45'),last.rename('akhir'),sku.rename('sku')],axis=1).fillna({'b13':0,'w45':0}).reset_index()
    Q=Q[Q.b13>0].sort_values('b13',ascending=False).head(top)
    Q['biasa']=Q.b13/13; Q['target']=Q.b13.apply(lambda v: max(1,math.ceil(v/13*2)))
    return Q
for code,lid,cust,desc in [('3100518','CC-2611-019','HARUM BUAH SEGAR','outlet pembeli kiwi 13 mgg (bazar & swalayan) — tawarkan Sungold punnet SP1 / curah SP1'),
                           ('3101242','CC-2611-020','RANUM BUAH','outlet pembeli kiwi 13 mgg — tawarkan Sungold SZ27 / SZ30 SP1 · SP2 + punnet SP1'),
                           ('3100434','CC-2611-021','GO BANANA','outlet pembeli kiwi 13 mgg — tawarkan Sungold SZ27 SP1 · SP2, Green SZ30 SP1')]:
    L.append(dict(code=f"L{len(L)+1}",id=lid,cust=cust,desc=desc,q=outk(code,top=(6 if code=='3100518' else 8)),kiwi=True))
for l in L:
    l['q']['sku_s']=l['q'].sku.map(short_sku)
    print(l['code'], l['q'][['OutletCode','OutletName','biasa','w45','akhir','sku_s','target']].to_string())
P['lamp']=L
# ---------------- STOCK 9 Nov — GT TGR share (aged > 3 days any grade + SP2/SP3; import SP2/SP3, SP1 if aged confirmed) ----------------
G=pd.read_pickle('stockgt9.pkl')
def gtq(p): return float(G[G.desc.str.contains(p,case=False,regex=True)]['gt'].sum())
def cnd(p): return float(G[G.desc.str.contains(p,case=False,regex=True)]['gt_cond'].sum())
def tq(p): return float(G[G.desc.str.contains(p,case=False,regex=True)]['total'].sum())
E=dict(he=gtq('Honey Emerald'),he_fresh=float(G[G.desc.str.contains('Honey Emerald Melon GH SP 15')]['fresh'].sum()),gold=gtq('Golden Melon'),rock=gtq('Rock Melon'),korean=gtq('Korean Musk'),sweet=gtq('Sweet Net'),
  cl8=gtq('Honi Pine SPA? ?CL 8'),cl10=gtq('Honi Pine SPA? ?CL 10'),cl15=gtq('CL 15'),cr9=gtq('CR 9'),cr7=gtq('CR 7'),cr5=gtq('CR 5'),
  kiwi_curah=gtq(r'Kiwi \d,5 SZ'),kiwi_curah_c=cnd(r'Kiwi \d,5 SZ'),kiwi_pnt=gtq(r'Kiwi (?:Pnt|Pouch)'),kiwi_pnt_c=cnd(r'Kiwi (?:Pnt|Pouch)'),kiwi_ori=tq(r'Kiwi \d,5 SZ')-gtq(r'Kiwi \d,5 SZ')-cnd(r'Kiwi \d,5 SZ'),
  c313=gtq('C3 13KG 1S$'),fs=gtq('Bn FS1? 13KG'),sweety=gtq('Sweety'),single=gtq('Single'),mini=gtq('Bn Mini'),fb=gtq('Bn FB1? 13KG'),guava=gtq('Crystal Guava'),
  singo=gtq('Singo'),gpear=gtq('Golden Pear'),dazzle=gtq('Dazzle'))
E['melon']=E['he']+E['gold']+E['rock']+E['korean']; E['ban_aged']=E['c313']+E['fs']+E['sweety']+E['single']+E['mini']+E['fb']
P['gtstk']=E; P['stkdate']='9 Nov'
mel_txt=f"melon GT TGR {n0(E['melon'])} box (HE {n0(E['he'])})"
OH=pd.read_pickle('melon_onhand.pkl').set_index('sub')
# ---------------- W46 LINES ----------------
def code_of(n): return g[g.cname==n].code.iloc[0]
lines=[
 dict(id='CC-2610-136',status='Lanjut',cust='Raksi Segar Nusantara',code='0003101202',fmt='Subdist',lamp='Lampiran L1',rung='① Jaga',typ='P',k2='M',
      do='Melon SP2 / SP1 > 3 hr (Golden, HE)',prod='Golden / HE (SP2 · SP1 > 3 hr)',target=40,tgt='40 / mgg',
      why=f"W45 melon 20 vs ±36; sell-out 11 box 2–8 Nov, stok subdist ±{n0(OH.loc['Raksi Segar Nusantara','est'])} → L1",
      detail=f"Lanjutan ke-1 dari 136 (Partial). Melon W45 20 BE (Rock C) vs biasa ±36. Sell-out outlet naik 4 → 11 box (2–8 Nov), 3/5 outlet Lampiran beli; sell-in − sell-out sejak 20 Okt ±{n0(OH.loc['Raksi Segar Nusantara','est'])} box. {mel_txt}.",sales='Ryan Enrico',like=0.6),
 dict(id='CC-2611-001',status='Baru',cust='Ranum Buah',code='0003101242',fmt='Subdist',lamp='Lampiran L2',rung='① Jaga',typ='P',k2='M',
      do='Melon SP2 / SP1 > 3 hr (Golden, HE, Korean)',prod='Golden / HE / Korean (SP2 · SP1 > 3 hr)',target=40,tgt='40 / mgg',
      why=f"Melon W45 0 vs ±54; sell-out 20 box 2–8 Nov, stok subdist ≈ 0 → L2",
      detail=f"Melon W45 0 BE vs biasa ±54 (Golden utama; Golden GT W45 243 BE vs 861 W44). Sell-in − sell-out sejak 20 Okt ≈ −8 box: outlet menghabiskan stok. {mel_txt}.",sales='Achmad Farchan Syahrul',like=0.6),
 dict(id='CC-2611-002',status='Baru',cust='Diqin Buah - Tanah Tinggi',code='0003100264',fmt='Grosir',lamp='Grosir — tanpa outlet',rung='① Jaga',typ='P',k2='M',
      do='Ganti Rock → HE (SP2 / SP1 > 3 hr)',prod='Honey Emerald (SP2 · SP1 > 3 hr)',target=30,tgt='30 / mgg',
      why=f"Melon W45 10 vs ±36; Rock GT TGR {n0(E['rock'])} box · Grosir: tanpa outlet",detail=f"Melon W45 10 BE vs biasa ±36 (Rock utama, Rock GT TGR {n0(E['rock'])} box) → HE. Target 30 (jatah melon GT TGR terbatas).",sales='Reza Adiputranto',like=0.6),
 dict(id='CC-2611-003',status='Baru',cust='Toko Buah Soleram',code='0003101931',fmt='Grosir',lamp='Grosir — tanpa outlet',rung='① Jaga',typ='P',k2='P',
      do='Ganti CR 7 → CR 9 / CR 5 (SP1 > 3 hr) atau CL 10',prod='Honi Pine CR 9 / CR 5 / CL 10',target=100,tgt='100 / mgg',
      why=f"Pine W45 70 vs ±101; CR 7 GT TGR {n0(E['cr7'])} box · Grosir: tanpa outlet",detail=f"Pine W45 70 BE vs biasa ±101 (CR 7 utama). Jatah GT TGR 9 Nov: CR 7 {n0(E['cr7'])}, CR 9 {n0(E['cr9'])}, CR 5 {n0(E['cr5'])}, CL 10 {n0(E['cl10'])} box; harga CR sama ±144 rb.",sales='Moreno Ananto',like=0.6),
 dict(id='CC-2611-004',status='Baru',cust='Mamad Buah Grosir',code='0003100996',fmt='Grosir',lamp='Grosir — tanpa outlet',rung='① Jaga',typ='P',k2='M',
      do='Ganti Rock → HE (SP2 / SP1 > 3 hr)',prod='Honey Emerald (SP2 · SP1 > 3 hr)',target=25,tgt='25 / mgg',
      why='Melon W45 10 vs ±25; Rock ketat · Grosir: tanpa outlet',detail='Melon W45 10 BE vs biasa ±25 (Rock utama) → HE.',sales='Reza Adiputranto',like=0.6),
 dict(id='CC-2611-005',status='Baru',cust='Grosir Buah Import Apple Shop',code='0003100476',fmt='Grosir',lamp='Grosir — tanpa outlet',rung='① Jaga',typ='P',k2='P',
      do='Ambil order Pine CL 10 (SP1 > 3 hr / SP2)',prod='Honi Pine CL 10',target=25,tgt='25 / mgg',
      why=f"Pine W45 0 vs ±25 (beli 3 mgg Okt); CL 10 GT TGR {n0(E['cl10'])} · Grosir: tanpa outlet",detail=f"Pine 20–40 BE/mgg di 3 minggu Okt, 0 sejak 2 Nov. CL 10 jatah GT TGR {n0(E['cl10'])} box (9 Nov).",sales='Moreno Ananto',like=0.6),
 dict(id='CC-2611-006',status='Baru',cust='Duta Buah',code='0003100295',fmt='Grosir',lamp='Grosir — tanpa outlet',rung='③ Perdalam',typ='G',k2='M',
      do='Order ulang HE — jaga pembeli baru',prod='Honey Emerald (SP2 · SP1 > 3 hr)',target=100,tgt='100 / mgg',
      why='W45 Met 170 BE (3, 4 & 8 Nov); sebelum Okt 0 melon · Grosir: tanpa outlet',detail=f"Pembeli HE baru: W44 60 → W45 170 BE (Met). Target 100 (bukan 170): {mel_txt} untuk 6 baris + Sastro.",sales='Reza Adiputranto',like=0.5),
]
_k=s[(s.br=='Tangerang')&(s.ch=='General Trade')&s.sku.str.contains('Kiwi')]
def _kv(code): a=_k[_k.code==code]; return a[(a.dt>='2026-10-05')&(a.dt<='2026-11-01')]['So BE'].sum()/4, a[a.dt>='2026-11-02']['So BE'].sum()
KIWI_GT=E['kiwi_curah']+E['kiwi_curah_c']+E['kiwi_pnt']+E['kiwi_pnt_c']
kiwi_txt=f"Jatah kiwi GT TGR (SP1 boleh, stok kiwi lebih): curah {n0(E['kiwi_curah']+E['kiwi_curah_c'])} + punnet {n0(E['kiwi_pnt']+E['kiwi_pnt_c'])} = {n0(KIWI_GT)} box; Sastro ±150/mgg"
lines+=[
 dict(id='CC-2611-019',status='Baru',cust='Harum Buah Segar',code='0003100518',fmt='Subdist',lamp='Lampiran L3',rung='③ Perdalam',typ='G',k2='I',
      do='Tawarkan kiwi SP1 (stok kiwi lebih) — punnet untuk bazar/swalayan',prod='Sungold Pnt 4PCS SP1 · SZ27/SZ30 SP1',target=20,tgt='20 (uji)',
      why='Kiwi 0 sejak 14 Agt (dulu lot 378 BE ORI); 20 outlet bazar beli 70 box 13 mgg → L3',
      detail=f"Ganti Import yang habis (Red Globe 0) dengan kiwi. Harum beli kiwi ORI 5 mgg (Mei–Agt, 50–200 BE/lot, ±460–485 rb). {kiwi_txt}.",sales='Rizqi Akbar Nabawi',like=0.5),
 dict(id='CC-2611-020',status='Baru',cust='Ranum Buah',code='0003101242',fmt='Subdist',lamp='Lampiran L4',rung='③ Perdalam',typ='G',k2='I',
      do='Tawarkan kiwi SP1 / SP2 ke outlet pembeli kiwi',prod='Sungold SZ27 / SZ30 SP1 · SP2 + Pnt SP1',target=40,tgt='40 (uji)',
      why='Kiwi 0 sejak 14 Agt (dulu lot 693 BE); 30 outlet beli 272 box 13 mgg (terakhir 3 Nov) → L4',
      detail=f"Outlet Ranum masih membeli kiwi dari stok lama (Minta, MB NOO lia tekno 3 Nov). Lot Longan habis → kiwi pengisi Import. {kiwi_txt}.",sales='Achmad Farchan Syahrul',like=0.5),
 dict(id='CC-2611-021',status='Baru',cust='Go Banana',code='0003100434',fmt='Subdist',lamp='Lampiran L5',rung='③ Perdalam',typ='G',k2='I',
      do='Naikkan kiwi: SP1 di harga ±SP2 sekarang',prod='Sungold SZ27 SP1 · SP2, Green SZ30 SP1',target=40,tgt='40 / mgg',
      why='Kiwi aktif: biasa ±17, W45 28 BE (4 Nov); KEYRA 214 box 13 mgg → L5',
      detail=f"Go Banana satu-satunya subdist yang masih rutin beli kiwi (SZ30/SZ27 SP2 ±434–450 rb). 3 outlet: KEYRA FRUITS, Toba Bintang Jaya, Ay Fresh. {kiwi_txt}.",sales='Rizqi Akbar Nabawi',like=0.5),
]
for l in lines:
    rr=D[(D.code==l['code'])&(D.k2==l['k2'])]
    l['b4']=float(rr.b4.iloc[0]) if len(rr) else 0; l['wk']=float(rr.wk.iloc[0]) if len(rr) else 0
    l['rp']=float(rr.rp.iloc[0]) if len(rr) and rr.rp.iloc[0]==rr.rp.iloc[0] else 300000.0
    base=l['wk'] if (l['rung']=='① Jaga' or l['status']=='Lanjut') else l['b4']
    l['val']=max(0,l['target']-base)*l['rp']*4.3/1e6; l['score']=l['val']*l['like']
    assert l['sales']; assert l['fmt']=='Grosir' or l['lamp'].startswith('Lampiran'), l['id']
for l in lines:
    if l['id'] in ('CC-2611-019','CC-2611-020','CC-2611-021'):
        l['b4'],l['wk']=_kv(l['code']); l['rp']=450000.0
        l['val']=max(0,l['target']-l['b4'])*l['rp']*4.3/1e6; l['score']=l['val']*l['like']
P['lines']=lines; P['dropped']=[]; P['mel_lines']=sum(l['target'] for l in lines if l['k2']=='M')
for l in lines: print(l['id'],l['cust'],l['k2'],round(l['b4']),round(l['wk']),l['target'],round(l['rp']),round(l['val'],1),round(l['score'],1),l['sales'])
from collections import Counter
print(Counter(l['sales'] for l in lines), 'Jaga', sum('Jaga' in l['rung'] for l in lines), 'melon BE', P['mel_lines'])
pickle.dump(P,open('P.pkl','wb'))
# ---------------- PERFORMANCE / FILL / SUPPLY ----------------
KN={'C':'Banana','P':'Pineapple','I':'Import','G':'Guava','L':'Local'}
g['kk']=g.kat
tg=pd.read_pickle('target.pkl'); tg=tg[(tg.SalesOffice=='Tangerang')&(tg.Channel=='General Trade')&tg.PRODH2.isin(KN.values())]
perf=[]
for k,nm in KN.items():
    r={}
    for mon,a,b in [('Sep','2026-09-01','2026-09-30'),('Okt','2026-10-01','2026-10-31')]:
        r[mon]=g[(g.kk==k)&(g.dt>=a)&(g.dt<=b)]['Net Sales Amt'].sum()
        r['bbp'+mon]=tg[(tg.PRODH2==nm)&(tg.Month==('September' if mon=='Sep' else 'October'))]['TARGET IDR'].sum()
    f=g[(g.kk==k)&(g.dt>='2026-11-02')&(g.dt<='2026-11-05')]; r['fill']=f['Delivery BE'].sum()/f['So BE'].sum()
    r['kat']=nm; perf.append(r)
tot={'kat':'Total buah'}
for c in ['Sep','bbpSep','Okt','bbpOkt']: tot[c]=sum(r[c] for r in perf)
f=g[(g.dt>='2026-11-02')&(g.dt<='2026-11-05')]; tot['fill']=f['Delivery BE'].sum()/f['So BE'].sum()
perf.append(tot); P['perf']=perf
for r in perf: print(r['kat'], round(r['Sep']/1e9,1), round(r['bbpSep']/1e9,1), round(r['Okt']/1e9,1), round(r['bbpOkt']/1e9,1), round(r['fill'],2))
# Nov MTD (delivered/net sales 1-5 Nov) — indicative
P['nov_mtd']=g[(g.dt>='2026-11-01')&(g.dt<='2026-11-05')]['Net Sales Amt'].sum()
# fill per k2 two weeks
fl=g[(g.dt>='2026-10-19')&(g.dt<='2026-11-01')].copy(); fl['wk']=np.where(fl.dt<='2026-10-25','w43','w44')
F=fl.groupby(['k2','wk'])[['So BE','Delivery BE']].sum(); F['f']=F['Delivery BE']/F['So BE']; P['fill']=F.f.unstack().to_dict('index'); print(P['fill'])
# kurang kirim by sku group 26 Oct-1 Nov
f=g[(g.dt>='2026-11-02')&(g.dt<='2026-11-05')].copy()
def grp(sk):
    for a,b in [('Cavendish Bn FB','Cav FB / FB1'),('Cavendish Bn Sml','Cav Small Hand'),('Cavendish Bn FS','Cav FS / FS1'),('Sweety','Cav Sweety'),('C3 13KG','Cav C3 13KG'),('C3 5KG','Cav C3 5KG'),('Cavendish Bn FK','Cav FK'),('Honi Pine CL','Pine CL'),('Honi Pine CR','Pine CR'),('Guava','Crystal Guava'),('Golden Melon','Golden Melon'),('Honey Emerald','Honey Emerald'),('Kiwi','Kiwi')]:
        if a in sk: return b
    return 'lain'
f['grp']=f.sku.map(grp); f['amt_per']=f['Net Sales Amt']/f['Net Sales BE'].replace(0,np.nan)
K=f.groupby('grp').agg(order=('So BE','sum'),deliv=('Delivery BE','sum'),amt=('Net Sales Amt','sum'),ns=('Net Sales BE','sum'))
K['short']=K.order-K.deliv; K['rp']=K.amt/K.ns; K['short_rp']=K.short*K.rp; K['fill']=K.deliv/K.order
K=K[K.index!='lain'].sort_values('short',ascending=False); print(K.round(2).head(8)); P['short']=K.head(6).reset_index().to_dict('records')
P['short_tot']=dict(be=K.short.sum(), rp=K.short_rp.sum())
# ---------------- STOCK COVER (all channel TGR run-rate 5 Oct-1 Nov) ----------------
t=s[(s.br=='Tangerang')&(s.dt>='2026-10-05')&(s.dt<='2026-11-01')]
def rr(pat,ch=None):
    a=t[t.sku.str.contains(pat,case=False,regex=True)]
    if ch: a=a[a.ch==ch]
    return a['Net Sales BE'].sum()/4
t45=s[(s.br=='Tangerang')&(s.dt>='2026-11-02')&(s.dt<='2026-11-08')]
def o45(pat,ch=None):
    a=t45[t45.sku.str.contains(pat,case=False,regex=True)]
    if ch: a=a[a.ch==ch]
    return a['So BE'].sum()
cov=[]
for nm,sp,bp in [('Sungold Kiwi SZ30','Sungold Kiwi 5,5 SZ 30','Sungold Kiwi 5,5KG SZ 30'),('Sungold Kiwi SZ27','Sungold Kiwi 5,5 SZ 27','Sungold Kiwi 5,5KG SZ 27'),('Green Kiwi SZ30','Green Kiwi 9,5 SZ 30','Green Kiwi 9,5KG SZ 30'),
                 ('Singo Pear','Singo Pear','Singo Pear'),('Golden Pear','Golden Pear','Golden Pear'),('Dazzle Apple','Dazzle','Dazzle'),
                 ('Honey Emerald','Honey Emerald','Honey Emerald'),('Golden Melon','Golden Melon','Golden Melon'),('Rock Melon','Rock Melon','Rock Melon'),
                 ('Honi Pine CL 8','Honi Pine SPA? ?CL 8','Honi Pine CL 8'),('Honi Pine CL 10','Honi Pine SPA? ?CL 10','Honi Pine CL 10'),('Honi Pine CL 15','CL 15','Honi Pine CL 15'),('Honi Pine CR 9','CR 9','Honi Pine CR 9'),('Honi Pine CR 7','CR 7','Honi Pine CR 7'),
                 ('Crystal Guava','Crystal Guava','Guava'),('Cav C3 13KG','C3 13KG 1S$','C3 13KG'),('Cav C3 5KG','C3 5KG','C3 5KG'),('Cav Small Hand','Sml Hand 13KG','Sml Hand'),('Cav FK','Bn FK 13KG','Bn FK 13KG')]:
    q=float(T[T.desc.str.contains(sp,case=False,regex=True)].qty.sum()); a4=float(T[T.desc.str.contains(sp,case=False,regex=True)].aged4.sum())
    r4=rr(bp); cov.append(dict(sku=nm,stk=q,aged=a4,gt=rr(bp,'General Trade'),mt=rr(bp,'Modern Trade'),tgr=r4,cover=q/r4 if r4 else np.nan,o45=o45(bp),o45gt=o45(bp,'General Trade'),o45mt=o45(bp,'Modern Trade')))
C=pd.DataFrame(cov); print(C.round(1).to_string()); P['cover']=C
pickle.dump(P,open('P.pkl','wb'))
# ---------------- AO / RO + data gate ----------------
m=pd.read_pickle('masters.pkl'); mo=m['mo']; ms=m['msub']
tgrd=ms[ms['Cabang (nama standar)']=='Tangerang']
ao=[]
for _,r in tgrd.iterrows():
    d=str(r['DistributorCode (ScyllaX)']).split('.')[0]
    ro=((mo.DistributorCode.astype(str).str.split('.').str[0]==d)&(mo.Status=='Active')&mo.OutletStatus.isin(['Active','New Open Outlet'])).sum()
    a=x[x.d==d]
    ao.append(dict(dist=r.DistributorName,d=d,ro=int(ro),ao_okt=a[(a.dt>='2026-10-01')&(a.dt<='2026-10-31')].OutletCode.nunique(),
                   ao_sep=a[(a.dt>='2026-09-01')&(a.dt<='2026-09-30')].OutletCode.nunique(),ao_nov=a[a.dt>='2026-11-01'].OutletCode.nunique(),
                   last=a.dt.max()))
AO=pd.DataFrame(ao); AO['r']=AO.ao_okt/AO.ro.replace(0,np.nan); AO=AO.sort_values('ro',ascending=False); print(AO.to_string())
P['aoro']=AO
P['gate']=dict(green=int((AO['last']>=pd.Timestamp('2026-11-06')).sum()),red=AO[(AO['last']<pd.Timestamp('2026-11-06'))|AO['last'].isna()].dist.tolist(),n=len(AO))
print(P['gate'])
pickle.dump(P,open('P.pkl','wb'))
# ---------------- SUPPLY COMMANDS W46 ----------------
mel=P['mel_lines']
P['sup']=[
 dict(id='CC-2611-007',typ='P',owner='PM Local (melon) + Supply Planning',prob=f"Jatah GT TGR (SP1 > 3 hr + SP2) melon 9 Nov {n0(E['melon'])} box (HE {n0(E['he'])}); HE SP1 segar {n0(E['he_fresh'])} box masuk jatah setelah > 3 hr. Order HE GT W45 607 BE di ±292 rb (SP ±319 rb) — naik dari 468 BE walau harga naik, karena Golden habis (GT 861 → 243 BE). 5 baris melon = {mel} BE + Sastro ±150.",
      cmd='Alokasikan HE yang lewat 3 hari ke GT TGR setiap hari (≥ 300 box W46). Harga HE SP GT tahan ±310–320 rb — permintaan masih di atas jatah; tidak perlu diskon.',test='Alokasi tertulis s/d Sel 10 Nov; melon GT TGR terkirim ≥ 90% W46',lines='136, 001, 002, 004, 006'),
 dict(id='CC-2611-008',typ='R',owner='Supply Planning Import + PM Import',prob='Century 0 sejak ±20 Okt (142 Failed). Import GT TGR W45 718 BE vs 4.915 (W44); terkirim Import 2–5 Nov 31% — Shantang, Red Globe, Longan, Jeruk Medan, Century stok 0.',
      cmd='Tulis tanggal + volume DS Century untuk GT TGR dan kalender lot Import GT 4 minggu ke depan (Red Globe, Longan, Jeruk).',test='Tanggal DS + volume tertulis s/d Rab 11 Nov',lines='antrean Import (16 sinyal)'),
 dict(id='CC-2611-009',typ='G',owner='Sales GT + SCM',prob='DS Papaya California Cilacap → GT TGR dimulai lagi (S&D 30 Okt, aksi 9). Pembeli GT: Ranum ±800–1.200 BE/bln (Jun–Sep), Raksi 362 BE Okt (W45 104 BE), Ayotama 204 BE Okt.',
      cmd='Tetapkan tanggal kirim pertama + volume/mgg DS papaya ke GT TGR; CC membuka baris papaya (Ranum, Raksi, Ayotama) begitu tanggal tertulis.',test='Tanggal + volume tertulis s/d Rab 11 Nov',lines='W47 (calon)'),
 dict(id='CC-2611-010',typ='P',owner='PM Import + MT',prob=f"Kiwi ORI {n0(E['kiwi_ori'])} box = MT/cabang lain (SZ30 ±28 mgg cover). Order TGR SZ30+SZ27 W45 971 BE (MT 729). Stok kiwi lebih → SP1 kiwi kini boleh ke GT TGR: curah {n0(E['kiwi_curah']+E['kiwi_curah_c'])} + punnet {n0(E['kiwi_pnt']+E['kiwi_pnt_c'])} box; Sastro W45 147 BE (Met).",
      cmd='Program MT + cabang lain untuk kiwi ORI. Lepas SP1 + SP2 kiwi ke GT TGR: Sastro ±150 + baris 019–021 (110 BE). Bila 019–021 jalan, siapkan harga lot ORI untuk subdist W47.',test='Kiwi SZ30+SZ27 order TGR ≥ 900 BE W46; kiwi GT TGR ≥ 250 BE',lines='019, 020, 021 + Sastro')
]
# ---------------- SECTION B ----------------
_f=g[(g.dt>='2026-11-02')&(g.dt<='2026-11-05')].groupby('k2')[['So BE','Delivery BE']].sum(); BF=(_f['Delivery BE']/_f['So BE']).to_dict()
P['B']=[
 ('Banana', f"{BF['C']*100:.0f}%", 'Cavendish dikeluarkan dari command W46 (keputusan 9 Nov) — tidak ada baris, A2, atau command supply banana minggu ini.', 'Jangan janji volume Cavendish lewat command; ikuti order desk biasa.'),
 ('Pineapple', f"{BF['P']*100:.0f}%", f"Jatah GT TGR 9 Nov (SP1 > 3 hr + SP2): CL 8 {n0(E['cl8'])} · CL 10 {n0(E['cl10'])} · CL 15 {n0(E['cl15'])} · CR 9 {n0(E['cr9'])} · CR 5 {n0(E['cr5'])} · CR 7 {n0(E['cr7'])}. Pine masuk harian.", 'CR 7 hampir 0 → tawarkan CR 9 / CR 5 / CL 10 (Soleram, Apple Shop). CL 15 & CR segar = MT.'),
 ('Import', f"{BF['I']*100:.0f}%", f"GT TGR: kiwi curah SP2 {n0(E['kiwi_curah'])} + SP1 {n0(E['kiwi_curah_c'])} (SP1 boleh — stok kiwi lebih) · punnet {n0(E['kiwi_pnt']+E['kiwi_pnt_c'])}; ORI {n0(E['kiwi_ori'])} = MT. Cherry SP2, Royal Gala SP3 → A2. Century/Shantang/Red Globe/Longan 0.", 'Kiwi SP1 + SP2 boleh dijanjikan (Sastro, 019–021) — terkirim kiwi 71% vs Import lain 12%. Apel/cherry grade rendah → A2. Lot Import lain habis → jangan janji (008).'),
 ('Guava', f"{BF['G']*100:.0f}%", f"Stok 483 box, semua SP1 segar (1 hr) → MT/cabang; GT TGR {n0(E['guava'])}. S&D: ±54 ton/mgg, kurang.", 'Tidak ada baris Guava; janji hanya bila SP1 > 3 hr / SP2.'),
 ('Local / melon', f"{BF['M']*100:.0f}%", f"{mel_txt} + HE segar {n0(E['he_fresh'])} menua > 3 hr minggu ini. Golden/Rock jatah GT {n0(E['gold'])}/{n0(E['rock'])}. Sweet Net C, Pandan SP2, Kopyor, Mangopride → A2. Papaya DS (009).", f"5 baris melon ({mel} BE). Jangan isi subdist yang masih pegang melon (lihat C)."),
]
P['B_note']="Terkirim = Delivery ÷ Order (BI, GT Tangerang, 2–5 Nov — delivery ≥ 4 hari, lengkap di interim terbaru). Stok = SAP Tangerang Senin 9 Nov; jatah GT TGR = umur > 3 hr (semua grade) + SP2/SP3; ORI / segar 0–3 hr = MT & cabang lain. Import tanpa kolom umur: SP1 dihitung bila gudang konfirmasi > 3 hr. Pemilik: melon 007 · Import 008 · papaya 009 · kiwi 010. Cavendish di luar command W46."
c=P['cnt45']
P['C']=[
 f"Hasil W45 (order lengkap s/d 8 Nov): {c.get('Met',0)} Met · {c.get('Partial',0)} Partial · {c.get('Failed',0)} Failed. Tambahan ±Rp {n1(P['inc45']/1e6)} jt/mgg. Sastro kiwi, Fresh Fruit Pine, Harum Pine & Berkah melon pulih lewat order 6–9 Nov.",
 f"Melon: Golden habis (GT 861 → 243 BE) → pembeli pindah ke HE (468 → 607 BE di ±319 rb SP). Jatah GT TGR {mel_txt} → 5 baris melon {mel} BE; harga rek. tahan ±320 rb.",
 "Subdist masih pegang melon (sell-in − sell-out 20 Okt–8 Nov): Jaya Abadi ±82, Citra ±68, Bogor ±68, Sejahtera ±65, DCC ±63, Batavia ±54, Go Banana ±53, Jaya Makmur ±50 → tidak diisi.",
 "Import GT W45 718 BE (W44 4.915), terkirim 31%: lot habis → 008. Antre: Ayotama/Harum/Ranum Import lot (stok 0) — Harum & Ranum diganti kiwi (019, 020), semangka (stok 0), Harum melon (sell-out 3 box), papaya (009).",
]
P['antre']=[
 ('Harum Buah Segar','Import',0,770,454.9,'Red Globe 6,5KG','Terkirim Import cabang 31%; SKU stok 0 → 008','Rizqi Akbar Nabawi'),
 ('Ranum Buah','Import',0,603,398.2,'Longan 10KG','Longan stok 0 → 008','Achmad Farchan Syahrul'),
 ('Ayotama Nusantara Mandiri','Banana',0,1750,739.8,'Cav FB 13KG','Cavendish di luar command W46','Reza Adiputranto'),
 ('Jaya Makmur Fruits','Import',200,355,103.9,'Jeruk Medan C 10KG','Terkirim 0% (stok 0) → 008','Sahrin Sahrin'),
 ('Sastro Fruits & Vegetables','Import',163,238,79.3,'Kiwi Pnt 4PCS','Import cabang 31% → B; kiwi W45 Met (147)','Reza Adiputranto'),
 ('Ranum Buah','Local',147,343,46.2,'Semangka Merah TB','Semangka stok 0 di TGR','Achmad Farchan Syahrul'),
 ('Harum Buah Segar','Melon',5,50,27.7,'HE SP','Stok subdist ±15 box, sell-out 3 box/mgg → tidak diisi','Rizqi Akbar Nabawi'),
 ('Sinar Buah Banten','Melon',0,32,21.9,'Golden Melon','Stok melon subdist ±44 box','Sahrin Sahrin'),
 ('Jaya Makmur Fruits','Melon',0,32,21.6,'Golden Melon','Stok melon subdist ±50 box','Sahrin Sahrin'),
 ('Katalia Jaya','Melon',0,33,21.3,'Golden Melon','Stok subdist ±22 box (sell-out 16/mgg) → cek W47','Sahrin Sahrin'),
 ('Bandar Buah - Cikampek','Melon',0,32,20.1,'Golden Melon','ScyllaX berhenti 1 Nov; stok ±17 box','Ni DG Putri Mega Gita'),
 ('Grosir Buah Joni','Pineapple',30,93,22.4,'Pine CL 10','Terkirim 67%; belum di roster','—'),
]
# grosir melon & others that passed gates but queued by cap
_a=J[(J.route=='A')&(J.cname.isin(['Setia Abadi','Suparmin','Kartolo Buah','UD.YAKINI BUAH','KOMPAK SUKSES','Pesona Buah - Ps. Tanah Tinggi','JML - Tanah Tinggi']))&(J.k2=='M')]
for _,r in _a.sort_values('score',ascending=False).iterrows():
    P['antre'].append((r.cname,'Melon',round(r.wk),round(r.b4),round(r.score,1),'HE / Golden','Lolos gate; Jaga 6/6 & jatah melon GT TGR habis untuk 5 baris → W47',r.sales if isinstance(r.sales,str) else '—'))
_b=J[(J.route=='A')&(J.cname=='Grosir Buah Segar-Segar')]
for _,r in _b.iterrows(): P['antre'].append((r.cname,'Pineapple',round(r.wk),round(r.b4),round(r.score,1),'Pine','Lolos gate; belum di roster salesman','—'))
P['signals']=dict(n=len(J),routeB=int((J.route!='A').sum()))
pickle.dump(P,open('P.pkl','wb'))
print(P['signals'], round(P['inc45']/1e6,1), round(P['so45']/1e6,1), P['cnt45'])
```

## elastic.py
```python
import pandas as pd, numpy as np, re
s=pd.read_pickle('s.pkl')
g=s[(s.br=='Tangerang')&(s.ch=='General Trade')&(s.dt>='2026-05-04')&(s.dt<='2026-11-08')].copy()
g['wk']=g.dt.dt.to_period('W-SUN').dt.start_time
def suf(sk):
    t=sk.split()[-1]
    return {'1':'SP','2':'SP2','3':'SP3','ORI':'ORI','13,5':'B135','14':'B14','15':'B15','16':'B16','1S':'B1S'}.get(t,t)
u=g.sku.drop_duplicates(); g['gr']=g.sku.map(dict(zip(u,u.map(suf))))
def weekly(mask):
    a=g[mask].groupby('wk').agg(q=('So BE','sum'),ns=('Net Sales BE','sum'),amt=('Net Sales Amt','sum'),n=('code','nunique'))
    a=a[(a.ns>0)&(a.q>0)]; a['p']=a.amt/a.ns; return a
def elast(a,minpts=6):
    if len(a)<minpts or a.p.max()/a.p.min()<1.08: return None
    X=np.log(a.p); Y=np.log(a.q); b=np.polyfit(X,Y,1)[0]
    r=np.corrcoef(X,Y)[0,1]
    return b,r,len(a)
if __name__=='__main__':
    for nm,pat,gr in [('HE','Honey Emerald','SP'),('HE','Honey Emerald','SP2'),('Golden','Golden Melon','SP'),('Golden','Golden Melon','SP2'),('Rock','Rock Melon','SP'),
                      ('Pine CL8/10','Honi Pine CL (8|10)PCS','SP'),('Pine CL8/10','Honi Pine CL (8|10)PCS','SP2'),('Pine CR','Honi Pine CR (7|9)PCS','SP'),('Pine CR','Honi Pine CR (7|9)PCS','SP2'),
                      ('Kiwi SG curah','Sungold Kiwi 5,5KG SZ (27|30)','SP'),('Kiwi SG curah','Sungold Kiwi 5,5KG SZ (27|30)','SP2'),('Kiwi SG curah','Sungold Kiwi 5,5KG SZ (27|30)','ORI'),
                      ('Dazzle','Dazzle','SP2'),('Dazzle','Dazzle','SP3'),('Fuji','Fuji Apple 17KG','SP2'),('Fuji','Fuji Apple 17KG','ORI'),('Royal Gala','Royal Gala','SP2'),
                      ('C3 13KG','Cavendish Bn C3 13KG','B135'),('C3 13KG','Cavendish Bn C3 13KG','B15'),('FS','Cavendish Bn FS 13KG','B135'),('Pandan Small','Pandan Cocopride Small','SP'),('Pandan Small','Pandan Cocopride Small','SP2'),('Kopyor','Kopyor','SP'),('Papaya C','P.California C 15KG','SP')]:
        a=weekly(g.sku.str.contains(pat,regex=True)&(g.gr==gr))
        e=elast(a)
        print(f"{nm:14s} {gr:4s} wks={len(a):2d} p_med={a.p.median()/1e3 if len(a) else 0:6.0f}k q_med={a.q.median() if len(a) else 0:7.0f} pmin={a.p.min()/1e3 if len(a) else 0:5.0f} pmax={a.p.max()/1e3 if len(a) else 0:5.0f} e={e}")
```

## price.py
```python
"""Recommended price per command line + perish-risk list (W46, stock 9 Nov).
Method: P_rec = customer's own realized price (main grade) in the weeks it bought its usual volume (5 Okt-1 Nov);
if the line target exceeds that volume, P_rec = P_b4 * (target/Q_b4)^(1/e), clipped to the GT TGR 26-week min for the grade.
e = GT TGR elasticity where the weekly log-log fit is credible (negative, |r|>=0.3, 8+ wks, -4<e<-0.5), else event-based/prior."""
import pandas as pd, numpy as np, pickle, math
from elastic import g, weekly, elast
from tfill import n0
P=pickle.load(open('P.pkl','rb')); R=pd.read_pickle('roster.pkl')
def sales(code,name):
    r=R[R.code==code]
    if len(r): return r.sales.iloc[0]
    n=str(name).lower()
    if 'kebun buah' in n: return 'Rizqi Akbar Nabawi'
    if 'kopkar' in n or 'koperasi' in n: return 'Bahar Ashary / Ryan Enrico (konfirmasi)'
    return 'Belum di roster'
B4=(g.dt>='2026-10-05')&(g.dt<='2026-11-01'); W45=(g.dt>='2026-11-02')
def r5(x): return int(round(x/5000.0)*5000)
# elasticities
E={}
def est(key,pat,gr,prior,note_prior):
    a=weekly(g.sku.str.contains(pat,regex=True)&(g.gr==gr)); e=elast(a,8)
    if e and -4<e[0]<-0.5 and e[1]<=-0.3: E[key]=(round(e[0],1),f"regresi mingguan GT TGR 26 mgg (r {e[1]:.2f}, {e[2]} mgg)")
    else: E[key]=(prior,note_prior)
    E[key]=E[key]+(a.p.min() if len(a) else np.nan, a.p.median() if len(a) else np.nan, a.q.median() if len(a) else np.nan)
est('HE_SP','Honey Emerald','SP',-1.5,'prior ±−1,5 — W45 harga HE SP naik 286→319 rb tapi volume GT naik (314→345 BE): elastisitas jangka pendek tidak terlihat, jadi tahan harga')
est('MEL_SP2','Melon','SP2',-3.0,'prior (regresi lemah)')
est('PINE_SP','Honi Pine C[LR]','SP',None,'harga rata ±1% 27 mgg → tidak ada bukti elastisitas')
est('KIWI_SP2','Sungold Kiwi 5,5KG SZ (27|30)','SP2',-1.0,'prior (regresi lemah r −0,19)')
est('BAN','Cavendish Bn C3 13KG','B135',-2.2,'prior')
est('APPLE','Apple','SP2',-1.5,'prior (regresi tidak kredibel)')
est('LOCAL','Pandan Cocopride','SP2',-1.5,'prior')
for k,v in E.items(): print(k,v[:2],'pmin',v[2])
LINE={ # id: (family regex, elasticity key)
 'CC-2610-136':('Melon','HE_SP'),'CC-2611-001':('Melon','HE_SP'),'CC-2611-002':('Honey Emerald','HE_SP'),
 'CC-2611-003':('Honi Pine CR (7|9)','PINE_SP'),'CC-2611-004':('Honey Emerald','HE_SP'),'CC-2611-005':('Honi Pine CL (8|10)','PINE_SP'),
 'CC-2611-006':('Honey Emerald','HE_SP')}
def mkt(pat,gr,mask=None):
    a=g[g.sku.str.contains(pat,regex=True)&(g.gr==gr)&(B4 if mask is None else mask)]
    return a['Net Sales Amt'].sum()/a['Net Sales BE'].sum() if a['Net Sales BE'].sum()>0 else np.nan
PR={}
KIWI=('CC-2611-019','CC-2611-020','CC-2611-021')
allL=P['lines']+P.get('dropped',[])
for l in allL:
    if l['id']=='CC-2611-014':
        b=g[(g.code==l['code'])&(g.gr=='B15')&(g.dt>='2026-08-10')]; pb=b['Net Sales Amt'].sum()/b['Net Sales BE'].sum()
        c3=g[g.sku.str.contains('Cavendish Bn C3 13KG')&(g.gr=='B135')&B4]; pc=c3['Net Sales Amt'].sum()/c3['Net Sales BE'].sum()
        PR[l['id']]=dict(txt=f'Kode 15 {r5(pb)/1000:.0f} · C3 {r5(pc)/1000:.0f} rb',detail=[dict(gr='B15',pref=pb,src='harga kode 15 Ujang Nadi 13 mgg',p45=np.nan,e=-2.2,enote='banana C3 regresi mingguan',prec=r5(pb))]); continue
    if l['id'] in KIWI:
        R13=(g.dt>='2026-08-10')&g.sku.str.contains('Sungold Kiwi|Green Kiwi')&~g.sku.str.contains('Pnt|Pouch')
        def mp(gr,mask=None):
            a=g[R13&(g.gr==gr)&(True if mask is None else mask)]; return a['Net Sales Amt'].sum()/a['Net Sales BE'].sum() if a['Net Sales BE'].sum()>0 else np.nan
        sp2=mp('SP2'); ori=mp('ORI'); own2=mp('SP2',g.code==l['code'])
        sp1=r5((sp2+ori)/2); p2=r5(own2 if not pd.isna(own2) and own2>sp2 else sp2); pnt=500000
        en='prior −1,0 (kiwi stok ±28 mgg → harga dorong volume)'
        src1=f"tengah SP2 ({sp2/1000:.0f} rb) & ORI ({ori/1000:.0f} rb) GT TGR 13 mgg"
        if p2>sp1: sp1=p2; src1=f"= harga SP2 yang biasa dibayar ({p2/1000:.0f} rb) — SP1 untuk harga yang sama"
        det=[dict(gr='SP1 curah',pref=sp1,src=src1,p45=np.nan,e=-1.0,enote=en,prec=sp1),
             dict(gr='SP2 curah',pref=p2,src=('harga SP2 sendiri 13 mgg' if p2!=r5(sp2) else 'harga SP2 GT TGR 13 mgg'),p45=np.nan,e=-1.0,enote=en,prec=p2),
             dict(gr='Pnt SP1',pref=pnt,src='GT nasional punnet SP1 ±547 rb −10% (stok lebih)',p45=np.nan,e=-1.0,enote=en,prec=pnt)]
        use={'CC-2611-019':[2,0],'CC-2611-020':[0,1,2],'CC-2611-021':[0]}[l['id']]
        PR[l['id']]=dict(txt=' · '.join(f"{det[i]['gr'].replace(' curah','')} {det[i]['prec']/1000:.0f}" for i in use)+' rb',detail=[det[i] for i in use])
        print(l['id'],PR[l['id']]['txt']); continue
    pat,ek=LINE[l['id']]
    c=g[(g.code==l['code'])&g.sku.str.contains(pat,regex=True)&(g.dt>='2026-09-01')]
    main=c.groupby('gr')['So BE'].sum().sort_values(ascending=False)
    mm=main[[x for x in main.index if x in('SP','SP2','SP3')]]
    grs=[x for x in mm.index if mm[x]>=0.25*mm.sum()][:2] or (['SP2'] if 'Kiwi' in pat else ['SP'])
    parts=[];detail=[]
    for gr in grs:
        cb=c[(c.gr==gr)&B4]; ns=cb['Net Sales BE'].sum()
        if ns>0: pref=cb['Net Sales Amt'].sum()/ns; src='harga sendiri 5 Okt–1 Nov'
        else:
            cb=c[c.gr==gr]; ns=cb['Net Sales BE'].sum()
            pref=cb['Net Sales Amt'].sum()/ns if ns>0 else mkt(pat,gr); src='harga sendiri sejak Sep' if ns>0 else 'harga pasar GT TGR'
        p45=mkt(pat,gr,W45)
        e,enote,pmin=E[ek][0],E[ek][1],E[ek][2]
        qb4=l['b4'] if l['b4']>0 else max(l['target'],1)
        prec=pref
        if e and l['target']>qb4*1.05:
            prec=pref*(l['target']/qb4)**(1/e); prec=max(prec, min(pmin,pref) if not np.isnan(pmin) else 0.85*pref)
        if ek=='HE_SP' and not pd.isna(p45) and p45>prec:
            prec=p45; src=src+'; dinaikkan ke harga pasar GT W45 (volume tetap naik di harga itu)'
        prec=r5(prec)
        parts.append(f"{gr} {prec/1000:.0f}")
        detail.append(dict(gr=gr,pref=pref,src=src,p45=p45,e=e,enote=enote,prec=prec))
    PR[l['id']]=dict(txt=' · '.join(parts)+' rb',detail=detail)
    print(l['id'],l['cust'][:18],PR[l['id']]['txt'],[(d['gr'],round(d['pref']/1e3),None if pd.isna(d['p45']) else round(d['p45']/1e3),d['e']) for d in detail])
P['price']=PR; P['elast']=E
pickle.dump(P,open('P.pkl','wb'))
```

## risk.py
```python
"""Perish-risk stock for GT TGR (stock 9 Nov): GT-eligible qty that won't clear at its historical GT TGR run-rate
inside the remaining shelf window. Clearance price from elasticity, floored at the lowest historical grade price."""
import pandas as pd, numpy as np, pickle
from elastic import g
R=pd.read_pickle('roster.pkl')
def r5(x): return int(round(x/5000.0)*5000)
def sales(code,name):
    r=R[R.code==code]
    if len(r): return r.sales.iloc[0]
    n=str(name).lower()
    if 'kebun buah' in n: return 'Rizqi Akbar Nabawi'
    if 'kopkar' in n or 'koperasi' in n: return 'Bahar Ashary / Ryan Enrico (konfirmasi)'
    return 'Belum di roster'
P=pickle.load(open('P.pkl','rb')); E=P['elast']; G=pd.read_pickle('stockgt9.pkl')
R8=(g.dt>='2026-09-14')&(g.dt<='2026-11-08'); R13=(g.dt>='2026-08-10')
# stock regex, label, BI family regex, BI grades (price/run-rate basis), shelf window days, elasticity key, reason
CAND=[
 (r'Raja Bulu','Pisang Raja Bulu (tua)',r'Raja Bulu',['B15','B16'],3,'BAN','pisang ≥ 4 hr di gudang Kadu'),
 (r'^Cherry','Cherry 5KG SP2',r'Cherry',None,7,'APPLE','umur simpan pendek; GT TGR ±1 BE/bln'),
 (r'Royal Gala','Royal Gala SP3 / SP2',r'Royal Gala',['SP3','SP2'],21,'APPLE','SP3; GT TGR 0 sejak 29 Okt'),
 (r'Fuji Apple','Fuji SP2 / SP3',r'Fuji Apple 17KG',['SP3','SP2'],21,'APPLE','sisa lot; grade rendah'),
 (r'Crimson','Crimson Grape 500G',r'Crimson Grape',None,10,'APPLE','umur simpan pendek; GT jarang'),
 (r'Sweet Net Melon C','Sweet Net Melon C (tua)',r'Sweet Net',None,5,'MEL_SP2','melon C ≥ 4 hr; GT ±2 BE/mgg'),
 (r'Pandan Cocopride SP (Small|Big)','Pandan Cocopride Small/Big SP2 & tua',r'Pandan Cocopride (Small|Big)',['SP2'],10,'LOCAL','SP2/tua 186 box vs SP2 GT ±15/mgg'),
 (r'Mangopride','Mangopride HM SP2',r'Mangopride',None,5,'LOCAL','mangga SP2; GT 6 BE sejak Mei'),
 (r'Kopyor','Kopyor Coconut (tua)',r'Kopyor',['SP'],10,'LOCAL','umur ≥ 4 hr'),
 (r'Persian Lime','Persian Lime C / SP',r'Persian Lime',None,10,'LOCAL','GT 0 sejak Agu'),
 (r'P.California C','Papaya California C (tua)',r'P.California C 15KG',['SP'],4,'LOCAL','pepaya ≥ 4 hr'),
]
rows=[]
for sp,lab,bp,grs,win,ek,why in CAND:
    st=G[G.desc.str.contains(sp,regex=True)]; q=float(st['gt'].sum())
    if q<1: continue
    fam=g[g.sku.str.contains(bp,regex=True)]
    fg=fam[fam.gr.isin(grs)] if grs else fam
    pdown=None
    if ek=='BAN':
        dn=fam[R13&fam.gr.isin(['B15','B16'])]; pdown=dn['Net Sales Amt'].sum()/dn['Net Sales BE'].sum() if dn['Net Sales BE'].sum()>0 else None
        fg=fam[fam.gr.isin(['B135','B1S','B14'])] if fam.gr.isin(['B135','B1S','B14']).any() else fam
    rr=(fam if ek=='BAN' else fg)[R8]['Net Sales BE'].sum()/8
    pr=fg[R13]; pref=pr['Net Sales Amt'].sum()/pr['Net Sales BE'].sum() if pr['Net Sales BE'].sum()>0 else (fam['Net Sales Amt'].sum()/max(fam['Net Sales BE'].sum(),1))
    pmin_hist=(fg.groupby('wk').apply(lambda d: d['Net Sales Amt'].sum()/d['Net Sales BE'].sum() if d['Net Sales BE'].sum()>0 else np.nan)).min() if len(fg) else np.nan
    days=q/(rr/7) if rr>0 else 999
    e=E[ek][0] or -1.5
    qneed=q/(win/7)
    prec=pref
    if rr>0 and qneed>rr: prec=pref*(qneed/rr)**(1/e)
    floor=0.7*pref if np.isnan(pmin_hist) else min(pmin_hist,0.85*pref)
    if rr==0: prec=0.85*pref
    prec=max(prec,floor,0.6*pref)
    by=fam[R13].groupby(['code','cname'])['Net Sales BE'].sum().sort_values(ascending=False)
    if grs: 
        low=fam[R13&fam.gr.isin(grs)].groupby(['code','cname'])['Net Sales BE'].sum().sort_values(ascending=False)
        by=pd.concat([low,by[~by.index.isin(low.index)]])
    buyers=[(c,n,round(v)) for (c,n),v in by.head(3).items() if v>0]
    risk=(days>win) or rr==0
    rows.append(dict(ek=ek,pdown=pdown,lab=lab,q=q,win=win,rr=rr,days=days,pref=pref,prec=r5(prec),disc=1-prec/pref,why=why,e=e,risk=risk,
        buyers=[n.strip()[:22] for c,n,v in buyers],buyer_sales=[sales(c,n) for c,n,v in buyers],buyer_raw=buyers))
Rk=pd.DataFrame(rows)
pd.set_option('display.width',250)
print(Rk[['lab','q','win','rr','days','pref','prec','disc','risk']].round(2).to_string())
for r in rows: print(r['lab'],r['buyer_raw'],r['buyer_sales'])
P['risk']=[r for r in rows if r['risk']]
pickle.dump(P,open('P.pkl','wb'))
# A2 rows (IDs continue after lines/supply)
SAL={'Achmad Farchan Syahrul':'Farchan','Reza Adiputranto':'Reza','Rizqi Akbar Nabawi':'Rizqi','Ni DG Putri Mega Gita':'Putri','Moreno Ananto':'Moreno','Ryan Enrico':'Ryan','Bahar Ashary / Ryan Enrico (konfirmasi)':'Bahar/Ryan','Sahrin Sahrin':'Sahrin','Belum di roster':'—'}
E_=P['gtstk']; A2=[]
for r in P['risk']:
    if r['ek']=='BAN' and r['rr']==0: to='tidak ada pembeli GT TGR → MT / online (PM)'; sl='PM / MT'
    else:
        to=' · '.join(r['buyers']) if r['buyers'] else 'tidak ada pembeli GT TGR → MT'
        sl=' · '.join(dict.fromkeys(SAL.get(x,x) for x in r['buyer_sales'])) or 'PM / MT'
    pr='—' if r['pref']==0 else f"{r['prec']/1000:.0f} rb ({-r['disc']*100:.0f}%)" if r['disc']>0.005 else f"{r['prec']/1000:.0f} rb"
    win={3:'Rab 11 Nov',4:'Kam 12 Nov',5:'Jum 13 Nov',7:'Sen 16 Nov',10:'Kam 19 Nov',21:'Sen 30 Nov'}[r['win']]
    A2.append(dict(lab=r['lab'],q=r['q'],why=r['why'] if r['rr']>0 else r['why']+'; tidak ada penjualan GT 8 mgg',price=pr,to=to,sales=sl,due=win,val=r['q']*(r['pref'] or 0),
                   rr=r['rr'],pref=r['pref'],prec=r['prec'],e=r['e'],buyers=r['buyers'],buyer_sales=r['buyer_sales']))
for i,a in enumerate(A2): a["id"]=f"CC-2611-{11+i:03d}"
P['A2']=A2; P['A2_val']=sum(a['val'] for a in A2)
pickle.dump(P,open('P.pkl','wb'))
for a in A2: print(a['id'],a['lab'],round(a['q']),a['price'],a['to'],a['sales'],a['due'])
print('val at risk',round(P['A2_val']/1e6,1))
# ---- LA2: target outlets (ScyllaX, TGR subdists, 13 weeks) for each A2 SKU
import math
x=pd.read_pickle('x.pkl'); x=x[x['trx type']=='ORDER'].copy(); x['d']=x.dist.astype(str)
ms=pd.read_pickle('masters.pkl')['msub']; t=ms[ms['Cabang (nama standar)']=='Tangerang']
dcode=t['DistributorCode (ScyllaX)'].astype(str).str.split('.').str[0]
SUBN=dict(zip(dcode,t.DistributorName)); SUBBI=dict(zip(dcode,'000'+dcode))
mo=pd.read_pickle('masters.pkl')['mo']; TYP=dict(zip(mo.OutletCode.astype(str),mo.OutletType.fillna('').astype(str)))
xt=x[x.d.isin(SUBN.keys())&(x.dt>='2026-08-10')&~x.ProName.str.contains('Juice',case=False)&~x.OutletName.astype(str).str.upper().str.contains('SHOP SALES|CUSTOMER GUDANG|CUSTOMER ONLINE')]
XP={'Pisang Raja Bulu':'Raja Bulu','Cherry':'Cherry','Royal Gala':'Royal Gala','Sweet Net':'Sweet Net','Pandan':'Pandan Cocopride','Mangopride':'Mangopride','Kopyor':'Kopyor','Persian Lime':'Persian Lime'}
for a in P['A2']:
    pat=next(v for k,v in XP.items() if a['lab'].startswith(k))
    q=xt[xt.ProName.str.contains(pat,case=False)].groupby(['d','OutletCode','OutletName']).agg(box=('box','sum'),last=('dt','max')).reset_index().sort_values('box',ascending=False).head(4)
    tot=q.box.sum(); outs=[]
    for _,r in q.iterrows():
        tg=max(1,min(math.ceil(a['q']*r.box/tot) if tot else 1, math.ceil(r.box/13*3)))
        outs.append(dict(sub=SUBN[r.d],sales=sales(SUBBI[r.d],SUBN[r.d]),code=r.OutletCode,name=str(r.OutletName).strip()[:22],typ=(TYP.get(str(r.OutletCode),'') or '')[:14],box=r.box,last=r['last'],target=tg))
    a['outlets']=outs
    subs=list(dict.fromkeys(o['sub'] for o in outs))
    gro=[b for b in a['buyers'] if not any(b.startswith(s_[:12]) for s_ in subs)]
    a['to']=(' · '.join(gro[:2])+(' · ' if gro[:2] and subs else '')+(' · '.join(s_[:16] for s_ in subs)+f" ({len(outs)} outlet LA2)" if subs else '')) or 'tidak ada pembeli GT TGR → MT'
    sl=[SAL.get(s_,s_) for s_ in a['buyer_sales'][:2]]+[SAL.get(o['sales'],o['sales']) for o in outs]
    a['sales']=' · '.join(dict.fromkeys(z for z in sl if z and z!='—')) or 'PM / MT'
    print(a['id'],a['lab'],'|',a['to'],'|',a['sales']); [print('   ',o['sub'],o['code'],o['name'],round(o['box']),o['target']) for o in outs]
pickle.dump(P,open('P.pkl','wb'))
```

## pricing.py
```python
import pandas as pd, numpy as np, pickle
s=pd.read_pickle('s.pkl'); P=pickle.load(open('P.pkl','rb'))
s['wk']=s.dt.dt.to_period('W-SUN').dt.start_time
def grade(sk):
    if sk.startswith('Sunpride A '): return 'SPA'
    t=sk.split()[-1]
    return {'1':'SP','2':'SP2','3':'SP3','ORI':'ORI'}.get(t,t)
u=s.sku.drop_duplicates(); gm=dict(zip(u,u.map(grade))); s['gr']=s.sku.map(gm)
r=s[s.dt>='2026-08-31']
def wkp(mask):
    a=r[mask].groupby('wk').agg(ns=('Net Sales BE','sum'),amt=('Net Sales Amt','sum'),so=('So BE','sum'),n=('code','nunique'))
    a['rp']=a.amt/a.ns.replace(0,np.nan); return a
he_gt=wkp((r.br=='Tangerang')&(r.ch=='General Trade')&r.sku.str.contains('Honey Emerald')&(r.gr=='SP'))
he_gt_all=wkp((r.br=='Tangerang')&(r.ch=='General Trade')&r.sku.str.contains('Honey Emerald'))
he_mt=wkp((r.ch=='Modern Trade')&r.sku.str.contains('Honey Emerald')&(r.gr=='SP'))
print('HE GT SP\n',he_gt.round(0).tail(6)); print('HE GT all\n',he_gt_all.round(0).tail(6)); print('HE MT\n',he_mt.round(0).tail(4))
# HE GT W45 by grade price
w=r[(r.br=='Tangerang')&(r.ch=='General Trade')&r.sku.str.contains('Honey Emerald')&(r.dt>='2026-10-26')]
print(w.groupby(['wk','gr']).agg(so=('So BE','sum'),ns=('Net Sales BE','sum'),amt=('Net Sales Amt','sum')).assign(rp=lambda d:d.amt/d.ns).round(0))
gold=wkp((r.br=='Tangerang')&(r.ch=='General Trade')&r.sku.str.contains('Golden Melon'))
print('Golden GT\n',gold.round(0).tail(6))
kiwi_gt=wkp((r.ch=='General Trade')&r.sku.str.contains('Sungold Kiwi 5,5KG SZ 30')); kiwi_mt=wkp((r.ch=='Modern Trade')&r.sku.str.contains('Sungold Kiwi 5,5KG SZ 30'))
print('kiwi so GT/MT', kiwi_gt.so.tail(6).round(0).tolist(), kiwi_mt.so.tail(6).round(0).tolist())
for nm,pat in [('CL 8','Honi Pine CL 8'),('CL 10','Honi Pine CL 10'),('CL 15','Honi Pine CL 15'),('CR 7','Honi Pine CR 7'),('CR 9','Honi Pine CR 9')]:
    a=r[(r.br=='Tangerang')&(r.ch=='General Trade')&r.sku.str.contains(pat)&(r.dt>='2026-10-05')]
    print(nm, round(a['Net Sales Amt'].sum()/a['Net Sales BE'].sum()), round(a['Net Sales BE'].sum()), 'W45 so', a[a.dt>='2026-11-02']['So BE'].sum())
dz=r[(r.br=='Tangerang')&r.sku.str.contains('Dazzle')&(r.dt>='2026-10-26')].groupby(['ch','wk'])[['So BE','Net Sales BE','Net Sales Amt']].sum(); print(dz)
pickle.dump(dict(he_gt=he_gt,he_mt=he_mt,he_gt_all=he_gt_all,gold=gold,kiwi_gt=kiwi_gt,kiwi_mt=kiwi_mt),open('pr.pkl','wb'))
```

## pricing2.py
```python
import pandas as pd, numpy as np, pickle
s=pd.read_pickle('s.pkl'); s['wk']=s.dt.dt.to_period('W-SUN').dt.start_time
def grade(sk):
    if sk.startswith('Sunpride A '): return 'SPA'
    t=sk.split()[-1]; return {'1':'SP','2':'SP2','3':'SP3','ORI':'ORI'}.get(t,t)
u=s.sku.drop_duplicates(); s['gr']=s.sku.map(dict(zip(u,u.map(grade))))
R={}
r13=s[(s.dt>='2026-08-10')&(s.dt<='2026-11-08')]
def px(mask):
    a=r13[mask]; return a['Net Sales Amt'].sum()/max(a['Net Sales BE'].sum(),1e-9), a['Net Sales BE'].sum()
GT=r13.ch=='General Trade'; MT=r13.ch=='Modern Trade'; TG=r13.br=='Tangerang'
R['pear']=[('Century Pear 15 kg','GT · ORI',*px(GT&TG&r13.sku.str.contains('Century Pear')),15),
 ('Singo Pear 9 kg','GT · SP2',*px(GT&r13.sku.str.contains('Singo')&(r13.gr=='SP2')),9),
 ('Singo Pear 9 kg','GT · SP',*px(GT&r13.sku.str.contains('Singo')&(r13.gr=='SP')),9),
 ('Singo Pear 9 kg','MT · ORI',*px(MT&r13.sku.str.contains('Singo')&(r13.gr=='ORI')),9),
 ('Golden Pear 9 kg','GT · ORI',*px(GT&r13.sku.str.contains('Golden Pear')&(r13.gr=='ORI')),9),
 ('Golden Pear 9 kg','MT · ORI',*px(MT&r13.sku.str.contains('Golden Pear')&(r13.gr=='ORI')),9)]
for x in R['pear']: print(x[0],x[1],round(x[2]),round(x[3]),round(x[2]/x[4]))
w=s[(s.dt>='2026-10-12')&(s.ch=='Modern Trade')]
mtp=[]
for nm in ['Golden Pear','Singo Pear']:
    a=w[w.sku.str.contains(nm)].groupby('wk').agg(so=('So BE','sum'),ns=('Net Sales BE','sum'),amt=('Net Sales Amt','sum')); a['rp']=a.amt/a.ns.replace(0,np.nan); mtp.append((nm,a)); print(nm,a.round(0))
R['mtpear']=mtp
# import GT TGR by family W42-W45
g=s[(s.br=='Tangerang')&(s.ch=='General Trade')&(s.kat=='I')&(s.dt>='2026-10-12')&(s.dt<='2026-11-08')]
def fam(x):
    for k in ['Century Pear','Red Globe','Longan','Shantang','Jeruk Medan','Fuji','Kiwi','Dazzle','Singo','Golden Pear','Shine Muscat']:
        if k in x: return k
    return 'Lainnya'
g=g.assign(f=g.sku.map(fam)); T=g.groupby(['f','wk'])['So BE'].sum().unstack(fill_value=0); T['tot']=T.sum(1); T=T.sort_values('tot',ascending=False); print(T.round(0))
R['impfam']=T
# kiwi national SZ30 weekly GT/MT 10 wks
k=s[s.sku.str.contains('Sungold Kiwi 5,5KG SZ 30')&(s.dt>='2026-08-31')&(s.dt<='2026-11-08')]
K=k.groupby(['wk','ch'])['So BE'].sum().unstack(fill_value=0); print(K[['General Trade','Modern Trade']].round(0)); R['kiwi']=K
kp=[]
for lab,m in [('GT · ORI',GT&r13.sku.str.contains('Sungold Kiwi 5,5KG SZ 30')&(r13.gr=='ORI')),('GT · SP2',GT&r13.sku.str.contains('Sungold Kiwi 5,5KG SZ 30')&(r13.gr=='SP2')),('MT · ORI',MT&r13.sku.str.contains('Sungold Kiwi 5,5KG SZ 30')&(r13.gr=='ORI'))]:
    kp.append((lab,*px(m)))
print(kp); R['kiwip']=kp
# dazzle weekly price GT/MT
d=s[s.sku.str.contains('Dazzle')&(s.dt>='2026-09-07')&(s.dt<='2026-11-08')]
Dz=d.groupby(['wk','ch']).agg(ns=('Net Sales BE','sum'),amt=('Net Sales Amt','sum')); Dz['rp']=Dz.amt/Dz.ns.replace(0,np.nan); Dz=Dz.rp.unstack(); print(Dz[['General Trade','Modern Trade']].round(0)); R['dz']=Dz
# pine GT TGR price 5 Oct-8 Nov SP
pn=[]
for nm,pat in [('Pine CL 8 SP','Honi Pine CL 8PCS 1'),('Pine CL 10 SP','Honi Pine CL 10PCS 1'),('Pine CL 15 SP','Honi Pine CL 15PCS 1'),('Pine CR 7 SP','Honi Pine CR 7PCS 1'),('Pine CR 9 SP','Honi Pine CR 9PCS 1')]:
    a=s[(s.br=='Tangerang')&(s.ch=='General Trade')&(s.sku=='Sunpride '+pat)&(s.dt>='2026-10-05')&(s.dt<='2026-11-08')]
    pn.append((nm,a['Net Sales Amt'].sum()/max(a['Net Sales BE'].sum(),1),a['Net Sales BE'].sum()))
print(pn); R['pine']=pn
pickle.dump(R,open('pr2.pkl','wb'))
```

## build_sheet.py
```python
from pptx import Presentation
import tfill as tf
from common import *
import re
mo=pd.read_pickle('masters.pkl')['mo']; TYP=dict(zip(mo.OutletCode.astype(str),mo.OutletType.fillna('').astype(str)))
def ss(x):
    x=re.sub(r'\s+(1|2|3|ORI)$','',x)
    for a,b in [('Shantang Daun Orange 6','Shantang'),('Century Pear SZ 60 ORI','Century Pear'),('Century Pear SZ 60','Century Pear'),('Red Globe 6,5KG SZ 20','Red Globe'),('Red Globe 6,5KG SZ 22','Red Globe')]: x=x.replace(a,b)
    return x

p=Presentation('tpl_sheet5.pptx')
s=p.slides[0]; sh=s.shapes; M=P['meta']; lines=P['lines']
tf.set_text(sh[3],[["Sales GT Tangerang · Minggu 46 (9–15 November 2026)"],["Final Senin 9 Nov (order W45 s/d 8 Nov, stok SAP 9 Nov) · dibagikan Selasa 10 Nov · tiap baris = 1 ID di Command Register"]])
nb=sum(l['status']=='Baru' for l in lines); nl=len(lines)-nb; nj=sum(jaga(l)=='Jaga' for l in lines)
tf.set_text(sh[10],[['A  TAKTIS — WAJIB MINGGU INI  ',f'{len(lines)}/10 slot · {nb} baru · {nl} lanjutan · Jaga {nj}/6 · jatah GT TGR = umur > 3 hr + SP2/SP3 · A2 stok berisiko hal. 4']])
t=sh[11].table; old_h=tf.table_height(t)
tf.add_col(t,5,0.7*914400); tf.set_widths(t,[0.35,0.42,1.05,1.1,0.95,0.62,0.72,1.37,0.57,0.35]); tf.cell(t.cell(0,6),'Harga rek. / box')
rows=[[sid(l['id']),l['status'],l['cust'],l['do'],l['prod'],l['tgt'],P['price'][l['id']]['txt'],l['why'],SHORT[l['sales']],'☐'] for l in lines]
def sty(i,j,row):
    z='F3F6F2' if i%2==0 else 'FFFFFF'
    if j==1 and row[1]=='Lanjut': z='FFF2CC'
    return dict(fill=z,bold=(j in (2,6)),color=('04582D' if j==6 else None))
tf.fill_table(t,rows,styles=sty)
dy=tf.table_height(t)-old_h
tf.shift([x for x in sh if x.top>sh[11].top+10],dy)
tf.set_text(sh[12],["Kuning = lanjutan W45 (maks 1×). Harga rek. = harga realisasi customer sendiri saat volume biasa (5 Okt–1 Nov), disesuaikan elastisitas bila target > biasa; batas bawah = harga terendah 26 mgg grade itu (metode: register sheet HARGA_REK). Baris Subdist punya Lampiran (L1–L5); Grosir jual langsung → dinilai dari order BI. HE: harga ditahan ±320 rb (tanpa diskon). Kiwi SP1 boleh ke GT (stok kiwi lebih) → 3 baris uji kiwi (019–021). 8 sinyal antre (Jaga 6/6). Cavendish di luar command W46."])
tf.set_text(sh[13],[['B  INFO SUPPLY CABANG TANGERANG (GT)  ','apa yang bisa dijanjikan ke outlet — masalah supply bukan tugas sales']])
t2=sh[14].table; tf.cell(t2.cell(0,1),'Terkirim 2–5 Nov')
def sty2(i,j,row):
    v=float(row[1].rstrip('%'))/100
    return dict(fill='FDE9E7' if v<0.6 else ('FFF2CC' if v<0.9 else 'E6F0E2'))
h0=tf.table_height(t2); tf.fill_table(t2,[list(r) for r in P['B']],styles=sty2)
tf.set_text(sh[15],[P['B_note']])
tf.set_text(sh[16],[['C  INFO MINGGU INI — STRATEGIS (tanpa aksi minggu ini)']]+['• '+c for c in P['C']])
para=[pp for pp in sh[18].text_frame.paragraphs]
tf.set_text(sh[18],[['Bagaimana CC menilai (Senin):'],[para[1].text],['Pertanyaan: CC Lead · deck WCR W46 · 3 kode S untuk 1 kategori dalam seminggu → kategori ditahan 2 minggu · Lampiran, A2 & LA2: hal. 2–5']])
# ---------- Lampiran pages ----------
L=P['lamp']
def lamp_rows(l):
    q=l['q']; out=[]
    for k,(_,r) in enumerate(q.iterrows(),1):
        out.append([str(k),r.OutletCode,r.OutletName.strip()[:24],(TYP.get(str(r.OutletCode),'') or r.OtTypeName or '')[:15],bnum(r.biasa),n0(r.w45),delta(r.biasa,r.w45),ss(r.sku_s),dmy(r.akhir),n0(r.target),'☐',''])
    return out
NOTE="Satuan = box ScyllaX. Biasa = rata-rata per minggu 5 Okt–1 Nov; W45 = 2–8 Nov (ScyllaX s/d 8 Nov). SKU utama dari ProCode ScyllaX. Target box = kembali ke biasa (coba: 1 box) — tawarkan melon SP2 / SP1 > 3 hr di harga rek. baris. Tulis box yang benar-benar diorder."
NOTE3="Satuan = box ScyllaX. Biasa = rata-rata pembelian Import per minggu 5 Okt–1 Nov; W45 = 2–5 Nov. SKU utama = buah Import yang biasa dibeli (stok kini 0). Target = box kiwi Sungold SZ30 / SZ27 untuk dicoba (±15% dari biasa, 1–5 box)."
def put(slide, hdr, tbl, note, l, notetext):
    tf.set_text(hdr,[[f"{l['code']}  {l['id']} · {l['cust']}  ", l['desc']]])
    t=tbl.table; tf.cell(t.cell(0,5),'W45 box')
    if l.get('kiwi'): tf.cell(t.cell(0,4),'Biasa box/mgg (13 mgg)'); tf.cell(t.cell(0,9),'Target kiwi box')
    tf.fill_table(t,lamp_rows(l),styles=lambda i,j,row: dict(fill='F3F6F2' if i%2==0 else 'FFFFFF'))
    tf.set_text(note,[notetext]); note.top=tbl.top+tf.table_height(t)+int(0.05*914400)
s3=p.slides[2]; x=s3.shapes
hdr_runs=[r.text for r in x[5].text_frame.paragraphs[0].runs]; print(hdr_runs)
put(s3,x[5],x[6],x[7],L[0],NOTE); put(s3,x[8],x[9],x[10],L[1],NOTE)

NOTEK="Satuan = box ScyllaX. Biasa = rata-rata per minggu 13 mgg (10 Agt–8 Nov) karena kiwi dibeli per lot; W45 = 2–8 Nov. Target box = ±2 minggu volume biasa outlet (min 1). Tawarkan SP1 (stok kiwi lebih) / SP2 di harga rek. baris; tulis box yang benar-benar diorder."
import copy as _cp
s5=p.slides[3]; v=s5.shapes
hdrs=[(v[5],v[6],v[7]),(v[8],v[9],v[10])]
# clone block 2 for L5
new=[]
for sh_ in (v[8],v[9],v[10]):
    el=_cp.deepcopy(sh_._element); v._spTree.append(el); new.append(el)
v=s5.shapes; blocks=[(v[5],v[6],v[7]),(v[8],v[9],v[10]),(v[-3],v[-2],v[-1])]
KL=[l for l in L if l.get('kiwi')]
top=v[5].top
for (h_,t_,n_),l in zip(blocks,KL):
    h_.top=top; t_.top=top+int(0.38*914400)
    put(s5,h_,t_,n_,l,NOTEK)
    top=n_.top+n_.height+int(0.1*914400)
print('kiwi page bottom', round(top/914400,2))
s2=p.slides[1]; y=s2.shapes
drop=[y[5]._element,y[6]._element,y[7]._element,y[11]._element,y[12]._element,y[13]._element]
tf.set_text(y[1],['STOK BERISIKO (A2)']); tf.set_text(y[2],['Untuk salesman pemilik customer di kolom Tawarkan ke · tidak dihitung dalam 10 slot · outlet target di hal. 5'])
# L3 has 8 rows: table grew; move nothing else (remove blocks 2,3)
A2=P['A2']
hdr,tb,nt=y[8],y[9],y[10]
hdr.top=y[5].top
tf.set_text(hdr,[['A2  HABISKAN STOK BERISIKO — jatah GT TGR  ',f"{len(A2)} SKU · ±Rp {n1(P['A2_val']/1e9)} M nilai stok · tawarkan di harga rek. sebelum batas"]])
t=tb.table; tf.del_cols(t,[7,8]); tf.set_widths(t,[0.22,0.35,1.2,0.42,1.42,0.78,1.6,0.68,0.45,0.38])
for j,h in enumerate(['#','ID','SKU (jatah GT TGR)','Stok box','Kenapa berisiko','Harga rek. / box','Tawarkan ke','Salesman','Batas','Hasil']): tf.cell(t.cell(0,j),h)
rows=[[str(k),sid(a['id']),a['lab'],n0(a['q']),a['why'],a['price'],a['to'],a['sales'],a['due'].replace(' Nov',''),'☐'] for k,a in enumerate(A2,1)]
from pptx.enum.text import PP_ALIGN
AL={2:PP_ALIGN.LEFT,4:PP_ALIGN.LEFT,6:PP_ALIGN.LEFT,7:PP_ALIGN.LEFT,8:PP_ALIGN.CENTER,9:PP_ALIGN.CENTER,3:PP_ALIGN.RIGHT,5:PP_ALIGN.RIGHT}
for j,h in AL.items(): tf.cell(t.cell(0,j),t.cell(0,j).text,align=h)
tf.fill_table(t,rows,styles=lambda i,j,r: dict(fill=('F3F6F2' if i%2==0 else 'FFFFFF'),bold=(j in (2,5)),align=AL.get(j)))
tb.top=hdr.top+hdr.height+int(0.06*914400); tf.est_heights(t,k=0.72)
tf.set_text(nt,["Stok = SAP 9 Nov, hanya jatah GT TGR (umur > 3 hr, SP2/SP3). Berisiko = stok tidak habis dalam sisa umur simpan pada run-rate GT TGR 8 mgg (pisang 3 hr, melon/mangga 5, cherry 7, kelapa/lime 10, apel 21 hr). Harga rek. = harga grade itu di GT TGR 13 mgg, diturunkan sesuai elastisitas sampai stok habis sebelum batas; batas bawah = harga terendah historis grade itu. Tawarkan ke = pembeli GT TGR 13 mgg (pembeli grade rendah didahulukan). Tidak dihitung dalam 10 slot."])
nt.top=tb.top+tf.table_height(t)+int(0.05*914400)
for e in drop: e.getparent().remove(e)
# ---------- LA2 page: target outlets for A2 ----------
s4=p.slides[4]; z=s4.shapes
rows=[]; 
for a in P['A2']:
    for o in a.get('outlets',[]):
        rows.append([a['id'],a['lab'],o['sub'],SAL.get(o['sales'],o['sales']) if 'SAL' in dir() else o['sales'],o['code'],o['name'],o['typ'],n0(o['box']),dmy(o['last']),a['price'].split(' (')[0],n0(o['target'])])
SALS={'Achmad Farchan Syahrul':'Farchan','Reza Adiputranto':'Reza','Rizqi Akbar Nabawi':'Rizqi','Ni DG Putri Mega Gita':'Putri','Moreno Ananto':'Moreno','Ryan Enrico':'Ryan','Bahar Ashary / Ryan Enrico (konfirmasi)':'Bahar/Ryan','Sahrin Sahrin':'Sahrin','Bahar Ashary':'Bahar'}
for r in rows: r[3]=SALS.get(r[3],r[3])
half=len(rows)//2
# split on SKU boundary nearest the middle
cut=min(range(1,len(rows)), key=lambda i: (rows[i][0]==rows[i-1][0], abs(i-half)))
parts=[rows[:cut],rows[cut:]]
W=[0.22,1.15,0.95,0.6,0.72,1.08,0.55,0.45,0.42,0.6,0.42,0.34]
H_=['#','SKU (A2)','Subdist','Salesman','Kode outlet','Nama outlet','Tipe','Beli 13 mgg','Akhir','Harga rek.','Target box','Hasil']
from pptx.enum.text import PP_ALIGN
for k,(hdr,tb,nt) in enumerate([(z[5],z[6],z[7]),(z[8],z[9],z[10])]):
    tf.set_text(hdr,[[f"LA2  OUTLET TARGET STOK BERISIKO ({k+1}/2)  ","outlet subdist yang membeli SKU ini 13 mgg terakhir (ScyllaX) — bantu habiskan A2"]])
    t=tb.table; tf.set_widths(t,W)
    for j,h in enumerate(H_): tf.cell(t.cell(0,j),h)
    pr=[[str(i+1+(0 if k==0 else cut)),f"{sid(r[0])} {r[1]}"]+r[2:]+['☐'] for i,r in enumerate(parts[k])]
    tf.fill_table(t,pr,styles=lambda i,j,row: dict(fill='F3F6F2' if i%2==0 else 'FFFFFF',bold=(j==1),align=(PP_ALIGN.LEFT if j in (1,2,3,4,5,6) else None)))
    tf.est_heights(t,k=0.6)
    if k==1: hdr.top=z[7].top+z[7].height+int(0.12*914400); tb.top=hdr.top+hdr.height+int(0.06*914400)
    tf.set_text(nt,["Outlet = pembeli SKU yang sama di subdist GT TGR, 13 mgg (ScyllaX, tanpa SHOP SALES). Target box = bagian stok A2 sesuai porsi beli outlet, maks ±3 minggu volume biasanya. Order lewat subdist-nya; salesman = pemilik subdist di roster. Tulis box yang benar-benar diorder." if k==1 else "Lanjut di bawah."])
    nt.top=tb.top+tf.table_height(t)+int(0.05*914400)
# reorder: slide 3 (L1,L2) before slide 2 (A2)
lst=p.slides._sldIdLst; ids=list(lst); lst.remove(ids[2]); lst.insert(1,ids[2])
ids=list(lst); lst.remove(ids[3]); lst.insert(2,ids[3])
p.save(OUT+'Command_Mingguan_W46_GT_Tangerang.pptx'); print('saved')
```

## build_review.py
```python
from pptx import Presentation
import tfill as tf, pickle, pandas as pd
from common import *
pr=pickle.load(open('pr.pkl','rb')); D=pd.read_pickle('D.pkl')
p=Presentation(A+'6e97c06b-Review_W44_GT_Tangerang_FINAL.pptx'); S=p.slides
c=P['cnt45']; rev=P['rev']
# 1 cover
sh=S[0].shapes
tf.set_text(sh[4],['Review Command W45']); tf.set_text(sh[19],['GT Tangerang · FINAL','Verify & Learn · Senin 9 Nov 2026'])
tf.set_text(sh[20],['ScyllaX 8 Nov · stok 9 Nov','S&D 30 Okt']); tf.set_text(sh[21],['CC Lead · Commercial Analyst','BI order s/d 8 Nov'])
# 2 results
sh=S[1].shapes
tf.set_text(sh[3],[f"W45 final: {c['Met']} Met, {c['Partial']} Partial, {c['Failed']} Failed"])
tf.set_text(sh[5],[['HASIL W45   ','Order 6–8 Nov masuk terlambat → FF Pine, Duta HE, kiwi Sastro jadi Met. HE subdist macet di outlet.']])
tf.set_text(sh[7],['Met','4','129 FF Pine 155 · 138 Ranum Pine 440 · 145 Duta HE 170 · 135 Sastro kiwi 147'])
tf.set_text(sh[8],['Partial','4','Sastro HE, Kopkar HE, Raksi, Harum'])
tf.set_text(sh[9],['Failed','2','146 Yitno HE · 140 Ranum Dazzle'])
tf.set_text(sh[10],['Harga HE SP GT','±319 rb','W44 ±286 rb · order SP 314 → 365 BE (naik)'])
tf.set_text(sh[11],['Baris sales W42–W45 (34)','Met 18%','Partial 53% · Failed 24% · Invalid 6%'])
rows=[[sid(v['id']),v['cust'],f"{v['prod']} {v['target']}",v['res'],v['verdict'],v['dec']] for v in rev]
VC={'Met':'E6F0E2','Partial':'FFF2CC','Failed':'FDE9E7'}
t=sh[12].table; tf.cell(t.cell(0,2),'Target W45'); tf.cell(t.cell(0,3),'Hasil W45 (order & sell-out 2–8 Nov)'); tf.cell(t.cell(0,5),'Keputusan W46'); tf.fill_table(t,rows,styles=lambda i,j,r: dict(fill=VC[r[4]]) if j==4 else dict(fill='F3F6F2' if i%2==0 else 'FFFFFF'))
tf.set_text(sh[13],["Baris sales M / P / F / I: W42 0/40/20/40% · W43 11/67/22/0% · W44 10/60/30/0% · W45 40/40/20/0%. Semua command W42–W45 (43): Met 23% · Partial 44% · Failed 26% · Invalid 7%."])
tf.set_text(sh[14],[['ARTINYA  ','Pine & pembeli langsung (Duta, Sastro) jalan; HE ke subdist berhenti saat outlet tidak menyerap. W46: melon subdist hanya ke Raksi & Ranum (stok ≈ 0), sisanya Grosir.']])
# 3 impact
sh=S[2].shapes
mp=[v for v in rev if v['verdict'] in('Met','Partial')]
inc=sum(v['inc'] for v in mp); val=sum(v['val'] for v in mp)
tf.set_text(sh[3],[f"Dampak W45: +Rp {n1(inc/1e6)} jt/mgg — Duta HE, Sastro & Pine"])
tf.set_text(sh[5],[['DAMPAK Rp   ',f"Nilai order baris Met/Partial Rp {n1(val/1e6)} jt. Pine didukung sell-out (Ranum 500, FF 53 box); Duta & Sastro = pembeli langsung."]])
rows=[]
for v in mp:
    so=f"{n0(v['so_box'])} box · {n1(v['so_rp']/1e6)} jt" if v.get('so_box') is not None else '— (Grosir)'
    rows.append([sid(v['id']),v['cust'],v['kat'],v['verdict'],n0(v['wk']),n0(v['b8']),n0(v['w44']),n0(v['rp']),f"Rp {n1(v['val']/1e6)} jt",(f"Rp {n1(v['inc']/1e6)} jt" if v['inc']>0 else '0'),(f"Rp {n1(v['inc']*4.3/1e6)} jt" if v['inc']>0 else '—'),so])
rows.append(['','Total','','',n0(sum(v['wk'] for v in mp)),n0(sum(v['b8'] for v in mp)),n0(sum(v['w44'] for v in mp)),'',f"Rp {n1(val/1e6)} jt",f"Rp {n1(inc/1e6)} jt",f"Rp {n1(inc*4.3/1e6)} jt",''])
t=sh[7].table; tf.cell(t.cell(0,4),'Order W45'); tf.cell(t.cell(0,6),'W44')
tf.fill_table(t,rows,styles=lambda i,j,r: dict(fill='F3F6F2' if i%2==0 else 'FFFFFF',bold=(r[1]=='Total')))
tf.est_heights(t); tf.below(sh[8],sh[7])
inc_so=[v for v in mp if v['inc']>0]
tf.set_text(sh[8],['•  Cara hitung: Rp/BE = harga realisasi customer sendiri 5 Okt–8 Nov. Tambahan = order W45 di atas rata-rata 8 mgg (7 Sep–1 Nov); di bawah = 0. Raksi & Harum dihitung pada kategori melon.',
 f"•  Duta Buah: 170 BE vs ±8 biasa → +Rp {n1(V['inc']/1e6) if (V:=next(v for v in mp if v['id']=='CC-2610-145')) else ''} jt/mgg; pembeli HE baru 2 minggu berturut-turut (3, 4 & 8 Nov).",
 f"•  Ranum Pine: 440 BE (+105 vs biasa) dan outlet beli 500 box 2–8 Nov → tambahan nyata. Sastro HE 119 BE (biasa ±49); melon total Sastro 320 BE (Golden 146).",
 ])
tf.set_text(sh[9],[['ARTINYA  ',f"Dari Rp {n1(inc/1e6)} jt/mgg, ±Rp {n1(next(v for v in mp if v['id']=='CC-2610-138')['inc']/1e6)} jt (Ranum Pine) didukung sell-out; Duta & Sastro dari pembeli langsung. Kopkar, Raksi, Harum di bawah biasa."]])
# 4 melon
sh=S[3].shapes
he=pr['he_gt'].loc[:'2026-11-02'].iloc[-10:]; cats=[f"{d.day}/{d.month}" for d in he.index]
tf.chart_data(sh[7].chart,cats,[('Harga HE SP GT TGR (rb/box)',[round(v/1000) for v in he.rp])])
tf.chart_data(sh[8].chart,cats,[('Order HE SP GT TGR (BE)',[round(v) for v in he.so])])
try: tf.set_text(sh[8].chart.chart_title.text_frame,['Order HE SP — GT Tangerang (So BE)'])
except Exception as e: print('title',e)
tf.color_points(sh[8].chart)
tf.set_text(sh[3],['Melon: harga HE ±319 rb, order HE GT tetap naik — Golden habis'])
tf.set_text(sh[5],[['MELON   ','HE SP GT TGR: ±286 rb & 314 BE (W44) → ±319 rb & 365 BE (W45); semua HE GT 468 → 607 BE. Golden GT 861 → 243 BE (stok 59).']])
MO=pd.read_pickle('melon_onhand.pkl').set_index('code')
SUBS=[(nm,c,MO.si[c],MO.so[c],d) for nm,c,d in [('Raksi','0003101202','Lanjut 136 (L1)'),('Ranum','0003101242','Baris 001 (L2)'),('Harum','0003100518','Tutup 147'),('Go Banana','0003100434','Antre (W47)'),('Depok City Cau','0003100254','Tidak diisi'),('Jaya Abadi','0003100735','Tidak diisi')]]
rows=[]
for nm,code,si,so,dec in SUBS:
    w=D[(D.code==code)&(D.k2=='M')]; w=float(w.wk.iloc[0]) if len(w) else 0
    rows.append([nm,n0(w),n0(si),n0(so),f"{si-so:+.0f}",dec])
t=sh[9].table; tf.cell(t.cell(0,1),'Order W45'); tf.cell(t.cell(0,2),'Sell-in 20 Okt–8 Nov'); tf.cell(t.cell(0,3),'Sell-out 20 Okt–8 Nov'); tf.cell(t.cell(0,5),'W46')
tf.fill_table(t,rows,styles=lambda i,j,r: dict(fill='F3F6F2' if i%2==0 else 'FFFFFF'))
tf.est_heights(t); tf.below(sh[10],sh[9],0.08)
tf.set_text(sh[10],['BACAAN','•  Harga HE SP GT naik ±286 → ±319 rb (+12%) tetapi order SP naik 314 → 365 BE: pembeli pindah dari Golden (habis) ke HE. Bacaan "order −15%" sebelumnya salah — order 6–8 Nov masuk terlambat.',
 '•  Duta Buah (170 BE) & Sastro (HE 119 + Golden 146) membeli langsung; subdist dengan stok melon (Jaya Abadi, Bogor, DCC ±60–80 box) tidak diisi.',
 '•  Jatah GT TGR (stok 9 Nov, umur > 3 hr + SP2/SP3): melon ±267 box (HE 230) → 5 baris melon W46 (235 BE) di harga rek. ±290–320 rb.'])
tf.set_text(sh[11],[['ARTINYA  ','Permintaan HE di atas jatah GT TGR: tahan harga ±320 rb, jangan diskon. Batasnya alokasi HE > 3 hr (007), bukan harga.']])
# 5 supply
sh=S[4].shapes
VS=P['revsup']; cs=pd.Series([r[4] for r in VS]).value_counts().to_dict()
tf.set_text(sh[3],[f"Command supply W45: {cs.get('Met',0)} Met, {cs.get('Failed',0)} Failed"])
tf.set_text(sh[5],[['SUPPLY & HARGA   ','HE, Dazzle & kiwi lolos test; Century tanpa tanggal; banana GT 38% (Cavendish di luar command W46).']])
t=sh[7].table; tf.cell(t.cell(0,5),'W46')
tf.fill_table(t,[[sid(r[0]),r[1],r[2],r[3],r[4],r[5]] for r in VS],styles=lambda i,j,r: dict(fill=VC.get(r[4],'FFFFFF')) if j==4 else dict(fill='F3F6F2' if i%2==0 else 'FFFFFF'))
tf.est_heights(t); tf.below(sh[8],sh[7],0.08)
tf.set_text(sh[8],['PELAJARAN','•  141 Met pada test stok (cover < 1 mgg); order HE GT naik 468 → 607 BE walau harga ±319 rb — W46 kunci alokasi HE > 3 hr, harga ditahan.',
 '•  143 Met: order BI kiwi TGR SZ30+SZ27 971 BE (MT 729) & Setia 15 BE. Cover SZ30 ±28 mgg → ORI ke MT; SP1 dibuka untuk GT TGR (010).',
])
tf.set_text(sh[9],[['ARTINYA  ','Masalah supply GT TGR terbesar: Import GT W45 718 BE, terkirim 31% (Century, Shantang, Red Globe, Longan = 0) — keputusan lot, bukan kunjungan.']])
# 6 W46 lines
sh=S[5].shapes
tf.set_text(sh[3],[f"W46 final: 1 lanjutan, {len(P['lines'])-1} baru, {len(P['sup'])} command supply / harga"])
tf.set_text(sh[5],[['COMMAND W46   ','Jaga 6/6 (melon ×4, Pine ×2); Perdalam 4: HE Duta + kiwi SP1 Harum, Ranum, Go Banana. A2: 011–018.']])
t=sh[7].table
tf.fill_table(t,[[sid(l['id']),l['status'],l['cust'],f"{l['do']} · {l['prod']}",l['tgt']] for l in P['lines']],styles=lambda i,j,r: dict(fill=('FFF2CC' if (j==1 and r[1]=='Lanjut') else ('F3F6F2' if i%2==0 else 'FFFFFF'))))
tf.set_text(sh[8],['COMMAND SUPPLY / HARGA','007 · PM Local: HE > 3 hr ke GT TGR ≥ 300 box; harga HE SP tahan ±310–320 rb','008 · Import: tanggal DS Century + kalender lot Import GT','009 · Sales + SCM: DS papaya Cilacap → GT TGR','010 · PM Import + MT: kiwi ORI → MT; SP1/SP2 → GT TGR (019–021)'])
tf.set_text(sh[9],["Final dari BI order s/d 8 Nov, ScyllaX s/d 8 Nov, stok SAP 9 Nov. Dibagikan Selasa 10 Nov; verdict W46 Senin 16 Nov. Detail harga: deck \"Analisis Harga & Run-rate W46\"."])
tf.set_text(sh[10],[['KEPUTUSAN  ','Setujui alokasi HE & tahan harga ±320 rb (PM Local), tanggal DS Century (Import), DS papaya (Sales/SCM), pelepasan kiwi SP1 ke GT TGR (PM Import). Cavendish di luar command W46.']])
p.save(OUT+'Review_W45_GT_Tangerang_FINAL.pptx'); print('ok')
```

## build_wcr.py
```python
from pptx import Presentation
import tfill as tf, pandas as pd
from common import *
p=Presentation(A+'6d408ad2-WCR_W45_Backing_Analysis_Tangerang.pptx'); S=p.slides
c=P['cnt45']; L=P['lines']; SUP=P['sup']; St=P['stk']; perf=P['perf']; tot=perf[-1]
Z=lambda i,j,r: dict(fill='F3F6F2' if i%2==0 else 'FFFFFF')
def H(sl,title,kick,body):
    tf.set_text(sl.shapes[3],[title]); tf.set_text(sl.shapes[5],[[kick+'   ',body]])
# 1 cover
sh=S[0].shapes
tf.set_text(sh[4],['Weekly Commercial Review W46']); tf.set_text(sh[19],['GT Tangerang — backing analysis Command W46','Final · Senin 9 Nov 2026'])
tf.set_text(sh[20],['BI s/d 8 Nov','ScyllaX 8 Nov · stok 9 Nov']); tf.set_text(sh[21],['Chair: GTM Head','Senin 9 Nov 2026 · 13:30'])
# 2 summary
s=S[1]; sh=s.shapes
H(s,f"W46 Tangerang: {len(L)} baris (1 lanjut, {len(L)-1} baru) + {len(SUP)} supply / harga",'RINGKASAN','Jatah GT TGR (SAP 9 Nov): melon ±267 → 5 baris melon + 2 Pine + 3 kiwi SP1 + 8 SKU A2; Cavendish di luar.')
tf.set_text(sh[7],['Hasil W45',f"{c['Met']} Met",f"{c['Partial']} Partial · {c['Failed']} Failed (final)"])
tf.set_text(sh[8],['Dampak W45',f"Rp {n1(P['inc45']/1e6)} jt",'tambahan / mgg — Duta HE & Ranum Pine'])
tf.set_text(sh[9],['Sinyal W46',str(P['signals']['n']),f"{P['signals']['routeB']} supply · 25 melon/stok 0"])
tf.set_text(sh[10],['Command W46',f"{len(L)} + {len(SUP)}",'baris sales + supply/harga'])
tf.set_text(sh[11],['Import GT W45','718 BE','dari 4.915 (W44) — lot habis'])
tf.set_text(sh[12],['YANG PERLU DIKETAHUI',
 f"•  Performa: GT TGR Okt {tot['Okt']/tot['bbpOkt']*100:.0f}% BBP (Sep {tot['Sep']/tot['bbpSep']*100:.0f}%); banana terkirim 38% (2–5 Nov) — Cavendish dikeluarkan dari command W46 (keputusan 9 Nov).",
 "•  W45 (order lengkap s/d 8 Nov): harga HE SP GT ±319 rb tetapi order SP 314 → 365 BE; semua HE GT 468 → 607 BE karena Golden habis (GT 861 → 243). Tahan harga, kunci alokasi (hal. 4, 8).",
 "•  Import GT W45 718 BE, terkirim 31%: Century, Shantang, Red Globe, Longan, Jeruk Medan stok 0 → 25 sinyal Import ke supply (008). Kiwi dinilai per SKU (terkirim 71%, stok ±28 mgg) → 3 baris kiwi SP1 (hal. 7, 9).",
 "•  Melon masih di 15 subdist (sell-in − sell-out sejak 20 Okt ±15–80 box) → tidak dibuat baris; hanya Raksi & Ranum (stok ≈ 0) (hal. 9).",
 "•  Order 6–8 Nov masuk terlambat: FF Pine, Sastro kiwi, Duta HE jadi Met; Harum Pine & Berkah melon pulih. Baru: melon Ranum, Diqin, Mamad; Pine Soleram & Apple Shop; HE Duta; kiwi Harum, Ranum, Go Banana. Aturan stok: > 3 hr (termasuk SP1) + SP2/SP3 = GT TGR."])
tf.set_text(sh[13],['KEPUTUSAN SENIN','•  PM Local: HE > 3 hr ke GT TGR ≥ 300 box; harga HE SP tahan ±310–320 rb (007).','•  Import: tanggal DS Century + kalender lot GT (008).','•  Sales + SCM: DS papaya Cilacap → GT TGR (009).','•  PM Import + MT: kiwi ORI → MT; SP1/SP2 → GT TGR (010).','•  Cavendish: di luar command W46 (keputusan 9 Nov).',f"•  Setujui {len(L)} baris W46."])
tf.set_text(sh[14],[['ARTINYA  ','W46 lebih sedikit baris karena barang yang biasa dibeli GT tidak ada atau masih di subdist: lot Import, alokasi melon, stok subdist — bukan kunjungan.']])
# 3 map
s=S[2]; sh=s.shapes
H(s,'Cara membaca: 9 tahap command v3.0 → angka W46 Tangerang','PETA OS v3.0','Urutan pembuatan command TOM v3.0 §4.3 — tiap tahap punya angka W46 dan halaman bukti.')
vals={9:['BI · ScyllaX · stok · S&D','4 sumber','hal. 5'],12:['stempel subdist + delivery ≥ 4 hr',f"{P['gate']['green']} hijau · {len(P['gate']['red'])} merah",'hal. 5'],15:['tangga peluang',f"{P['signals']['n']} sinyal",'hal. 6'],
      18:['supply · stok · harga','45 → supply · 10 stok 0','hal. 7–8'],21:['nilai × peluang · jatah GT TGR','14 lolos · 8 antre (cap)','hal. 9–10'],24:['aksi · SKU · outlet',f"{len(L)} baris · 5 Lampiran",'hal. 11–12'],
      27:['WCR Sen → issue Sel','Senin 9 Nov','hal. 13'],30:['kode hasil per baris','s/d Sen 16 Nov','hal. 13'],33:['W45 dinilai',f"{c['Met']} Met · Rp {n1(P['inc45']/1e6)} jt",'hal. 4']}
for k,v in vals.items(): tf.set_text(sh[k],v)
nj=sum(jaga(l)=='Jaga' for l in L)
rows=[['Tipe command (R/P/G/S/D)',f"Baris: P {sum(l['typ']=='P' for l in L)} · G {sum(l['typ']=='G' for l in L)} · supply R {sum(x['typ']=='R' for x in SUP)} · P {sum(x['typ']=='P' for x in SUP)} · G {sum(x['typ']=='G' for x in SUP)} → forward share (P+G) {round((sum(l['typ'] in 'PG' for l in L)+sum(x['typ'] in 'PG' for x in SUP))/(len(L)+len(SUP))*100)}%"],
      ['Lanjut maks 1×','9 baris W45 ditutup (4 Met); hanya 136 Raksi (Partial, baru di W45) lanjut 1× dengan Lampiran L1'],
      ['Maks 6 Jaga · 2 baris/customer · 5/salesman',f"Jaga {nj}/6 (penuh) + 4 Perdalam (Duta HE, kiwi ×3); Ranum 2 baris; Reza 3/5; 8 sinyal antre; 10/10 slot"],
      ['Command Card 9 field','contoh lengkap: 001 Ranum — melon SP2/tua ke subdist yang outletnya menghabiskan stok (hal. 12)']]
t=sh[34].table; tf.cell(t.cell(0,1),'Di W46 Tangerang'); tf.fill_table(t,rows,styles=Z)
tf.set_text(sh[35],[['ATURAN  ','Final Senin dengan order minggu penuh (terbooking s/d Min 8 Nov). Fill dari delivery 2–5 Nov (≥ 4 hari, lengkap).']])
# 4 verify W45
s=S[3]; sh=s.shapes
H(s,f"Tahap 9 (W45): {c['Met']} Met, {c['Partial']} Partial, {c['Failed']} Failed — Rp {n1(P['inc45']/1e6)} jt tambahan",'TAHAP 9 · VERIFY W45','Order 6–8 Nov terlambat: Pine FF & Ranum, HE Duta, kiwi Sastro Met. Supply: 141, 143, 144 Met · 142, 116 Failed.')
rows=[[sid(v['id']),v['cust'],f"{v['prod']} {v['target']}",v['res'],v['verdict'],(f"Rp {n1(v['inc']/1e6)} jt" if v['inc']>0 else ('0' if v['verdict']!='Failed' else '—')),v['dec']] for v in P['rev']]
VC={'Met':'E6F0E2','Partial':'FFF2CC','Failed':'FDE9E7'}
t=sh[7].table; tf.cell(t.cell(0,2),'Target W45'); tf.cell(t.cell(0,3),'Hasil 2–8 Nov'); tf.cell(t.cell(0,6),'W46')
tf.fill_table(t,rows,styles=lambda i,j,r: dict(fill=VC[r[4]]) if j==4 else Z(i,j,r)); tf.est_heights(t); tf.below(sh[8],sh[7])
tf.set_text(sh[8],['Tambahan = order W45 di atas rata-rata 8 mgg × Rp/BE customer (deck Review W45 hal. 3). Sell-out = ScyllaX subdist → outlet 2–8 Nov. Final.'])
tf.set_text(sh[9],[['PELAJARAN  ','HE ke subdist jalan hanya bila outlet menghabiskan stok (Ranum ya; Harum tidak). W46: melon subdist hanya Raksi (lanjut 1×) & Ranum; sisanya Grosir.']])
# 5 data
s=S[4]; sh=s.shapes
red=', '.join(P['gate']['red'][:3])
H(s,'Tahap 1–2: data W46 dan gate delivery ≥ 4 hari','TAHAP 1–2 · DATA',f"{P['gate']['green']}/25 subdist hijau (ScyllaX ≥ 6 Nov); merah: Ayotama, Batavia, Mandiri, Bandar, BSN.")
tf.fill_table(sh[7].table,[['BI interim v1.4.2 (6)','order s/d 8 Nov (booking); delivery andal s/d 5 Nov','Sinyal W46, verdict W45, performa, fill, harga realisasi'],
 ['ScyllaX (interim)','s/d 8 Nov (20 subdist) · produk SSN saja','Lampiran outlet, AO/RO, sell-out & stok subdist melon'],
 ['Stok SAP Tangerang','Senin 9 Nov · All (termasuk banana)','Stock gate SKU, cover vs run-rate, umur ≥ 4 hari'],
 ['Notulen S&D','Weekly 30 Okt (Pak Sandy, DSP)','Supply: banana C3 ±37rb, Honi ±43rb, CR 7, DS Century & papaya, Dazzle s/d W47']],styles=Z)
tf.fill_table(sh[8].table,[['Sinyal Jaga','order 2–8 Nov vs rata-rata 5 Okt–1 Nov','minggu W45 penuh vs 4 minggu normal'],
 ['Terkirim (fill)','2–5 Nov','delivery 6–8 Nov belum lengkap'],
 ['Performa','Sep · Okt vs BBP (RAW_TARGET)','Okt kini lengkap'],
 ['Verdict W45','order 2–8 Nov · ScyllaX 2–8 Nov','sheet W45 dibagikan Selasa 3 Nov']],styles=Z)
tf.set_text(sh[9],[['CATATAN  ','Order GT 6–8 Nov bertambah sejak file sebelumnya (6 Nov 14.885 → 15.802 BE; 8 Nov 11.169 → 12.515 BE). Verdict W45 dihitung ulang: 3 baris naik ke Met, 9 sinyal pulih.']])
# 6 detect perf
s=S[5]; sh=s.shapes
H(s,f"Tahap 3: GT Tangerang {tot['Sep']/tot['bbpSep']*100:.0f}% BBP Sep, {tot['Okt']/tot['bbpOkt']*100:.0f}% di Okt",'TAHAP 3 · DETECT','Okt membaik lewat Pine & Import (lot Shantang/Red Globe); Nov dimulai tanpa lot Import dan dengan banana 38%.')
route={'Banana':'Di luar command W46','Pineapple':'CR 7 habis → 2 baris (Soleram, Apple Shop)','Import':'Lot habis → 008; kiwi per SKU → 019–021','Guava':'Supply → B','Local':'Melon jatah GT → 5 baris + 007','Total buah':'—'}
nl={'Banana':0,'Pineapple':2,'Import':3,'Guava':0,'Local':5,'Total buah':len(L)}
rows=[[r['kat'],f"{n1(r['Sep']/1e9)} M",f"{n1(r['bbpSep']/1e9)} M",f"{r['Sep']/r['bbpSep']*100:.0f}%",f"{n1(r['Okt']/1e9)} M",f"{r['Okt']/r['bbpOkt']*100:.0f}%",f"{r['fill']*100:.0f}%",route[r['kat']],str(nl[r['kat']])] for r in perf]
t=sh[7].table; tf.cell(t.cell(0,6),'Terkirim 2–5 Nov'); tf.cell(t.cell(0,8),'Baris W46')
tf.fill_table(t,rows,styles=lambda i,j,r: dict(fill='F3F6F2' if i%2==0 else 'FFFFFF',bold=(r[0]=='Total buah')))
tf.set_text(sh[8],['Net sales BI GT Tangerang (buah) vs BBP 2026. Rute: terkirim < 90% atau stok SKU kosong = supply; harga di luar wewenang salesman = command PM.'])
tf.set_text(sh[9],[f"•  Banana = {perf[0]['Okt']/tot['Okt']*100:.0f}% revenue Okt; terkirim 38% (2–5 Nov) → porsi alokasi; Cavendish ditahan dari command W46.",
 "•  Import 103% BBP Okt berkat lot Shantang & Red Globe (Harum, Ranum ±1.000–1.900 BE/mgg) — lot habis, W45 718 BE: Nov akan jatuh tanpa lot baru (008).",
 "•  Local 72% BBP Okt: melon naik lewat harga HE W44 dan bertahan W45 (HE menggantikan Golden yang habis). Pine 139% BBP Okt."])
tf.set_text(sh[10],[['ARTINYA  ','Sales W46: melon, Pine & kiwi SP1 (barang ada di jatah GT TGR). Import lain menunggu lot; Cavendish di luar command.']])
# 7 stock
s=S[6]; sh=s.shapes
H(s,'Tahap 4: stok 9 Nov + S&D menentukan apa yang dijanjikan','TAHAP 4 · STOK','Kiwi surplus; HE, Pine, Guava ketat (normal fresh); Golden/Rock & lot Import kosong; banana ada tapi terkirim 38%.')
C=P['cover'].set_index('sku')
def cv(x): return f"{x:.0f} mgg" if x>=1.5 else ('< 1 mgg' if x<1 else f"{n1(x)} mgg")
rows=[['Sungold Kiwi SZ30 + SZ27',n0(St['kiwi30']+St['kiwi27']),n0(C.loc['Sungold Kiwi SZ30','tgr']+C.loc['Sungold Kiwi SZ27','tgr']),cv((St['kiwi30']+St['kiwi27'])/(C.loc['Sungold Kiwi SZ30','tgr']+C.loc['Sungold Kiwi SZ27','tgr'])),'—','ORI = MT/cabang; GT TGR SP/SP2 '+n0(P['gtstk']['kiwi_curah'])+' + punnet '+n0(P['gtstk']['kiwi_pnt']+P['gtstk']['kiwi_pnt_c'])+' (SP1 boleh) → Sastro, 019–021'],
 ['Green Kiwi SZ30',n0(St['green30']),n0(C.loc['Green Kiwi SZ30','tgr']),cv(C.loc['Green Kiwi SZ30','cover']),'—','Ikut program kiwi'],
 ['Singo / Golden Pear',f"{n0(St['singo'])} / {n0(St['gpear'])}",n0(C.loc['Singo Pear','tgr']+C.loc['Golden Pear','tgr']),'±3 mgg','MT ambil ±700–900/mgg','MT; bukan pengganti Century di GT'],
 ['Honey Emerald',n0(St['he']),n0(C.loc['Honey Emerald','tgr']),cv(C.loc['Honey Emerald','cover']),'jatah GT TGR: '+n0(P['gtstk']['he'])+' (> 3 hr + SP2)','5 baris melon (235 BE) → 007'],
 ['Golden / Rock Melon',f"{n0(St['gold'])} / {n0(St['rock'])}",n0(C.loc['Golden Melon','tgr']+C.loc['Rock Melon','tgr']),'< 1 mgg','jatah GT TGR: '+n0(P['gtstk']['gold'])+' / '+n0(P['gtstk']['rock']),'Ranum (Golden), lainnya HE'],
 ['Honi Pine CL 8 / 10 / 15',f"{n0(St['cl8'])} / {n0(St['cl10'])} / {n0(St['cl15'])}",n0(C.loc['Honi Pine CL 8','tgr']+C.loc['Honi Pine CL 10','tgr']+C.loc['Honi Pine CL 15','tgr']),'< 1 mgg','jatah GT TGR: '+n0(P['gtstk']['cl8'])+' / '+n0(P['gtstk']['cl10'])+' / '+n0(P['gtstk']['cl15']),'Apple Shop (005): CL 10; CL 15 segar = MT'],
 ['Honi Pine CR 7 / CR 9',f"{n0(St['cr7'])} / {n0(St['cr9'])}",n0(C.loc['Honi Pine CR 7','tgr']+C.loc['Honi Pine CR 9','tgr']),'< 1 mgg','jatah GT TGR: '+n0(P['gtstk']['cr7'])+' / '+n0(P['gtstk']['cr9'])+' (tua) · CR 5 '+n0(P['gtstk']['cr5']),'Soleram (003): CR 9 / CR 5 ganti CR 7'],
 ['Century / Shantang / Red Globe / Longan','0','—','—','DS Century belum jalan','25 sinyal Import → supply (008)'],
 ['Banana C3 13KG',n0(St['c313']),n0(C.loc['Cav C3 13KG','tgr']),cv(C.loc['Cav C3 13KG','cover']),'jatah GT TGR: '+n0(P['gtstk']['c313'])+' (umur > 3 hr)','Di luar command W46']]
t=sh[7].table; tf.cell(t.cell(0,1),'Stok 8 Nov'); tf.cell(t.cell(0,4),'Jatah GT TGR (9 Nov)'); tf.cell(t.cell(0,5),'Arti untuk command W46')
tf.fill_table(t,rows,styles=Z); tf.est_heights(t); tf.below(sh[8],sh[7])
tf.set_text(sh[8],['Stok = SAP Tangerang 9 Nov. Kolom 5 = jatah GT TGR SAP 9 Nov (umur > 3 hr semua grade + SP2/SP3; Import tanpa umur: SP2/SP3, SP1 bila gudang konfirmasi > 3 hr; kiwi SP1 dibuka karena stok kiwi lebih). ORI & segar 0–3 hr = MT/cabang lain. Run-rate = Net Sales BE TGR 5 Okt–1 Nov.'])
tf.set_text(sh[9],[['ARTINYA  ','Stock gate W46 dihitung pada jatah GT TGR, bukan total gudang: kiwi ±19rb box di gudang, GT TGR dapat SP2 + SP1 ±795 box (ORI tetap MT).']])
# 8 price
s=S[7]; sh=s.shapes
H(s,'Tahap 4 (harga): HE responsif dua arah; Import = soal lot','TAHAP 4 · HARGA','Ringkasan deck Pricing & Run-rate W46 — harga realisasi BI per grade & channel, dibandingkan run-rate.')
rows=[['Honey Emerald','GT SP ±286 (W44) → ±319 rb (W45); MT ±433 rb','GT SP 314 → 365 BE; semua HE GT 468 → 607','Tahan ±310–320 rb + alokasi ≥ 300 (007)'],
 ['Golden / Rock Melon','GT Golden ±194 → ±124 rb (grade C/SP2)','stok 59 / 48; MT ambil 500 / 284 BE','Tanpa baris; HE pengganti'],
 ['Pear vs Century','Century ±18 rb/kg; Singo ±47 rb/kg; Golden ±55 rb/kg','MT Singo ±920 & Golden ±720 BE W45; GT ±2','Lot / DS Century untuk GT (008)'],
 ['Kiwi Sungold','GT ORI ±514 rb vs MT ±565 rb','SZ30 nasional W45 GT 94 · MT 381 BE (turun)','ORI → MT; SP1 ke GT ±440–500 rb (010, 019–021)'],
 ['Dazzle','GT ±300–400 rb vs MT ±880 rb','stok 112 box, keluar lewat MT & online','Tanpa tindakan; habis W47'],
 ['Pine','CL ±133 rb, CR ±143 rb (rata)','CL 8/10/15 ±14,8rb BE GT W45','Tidak ada tindakan harga']]
tf.fill_table(sh[7].table,rows,styles=Z)
tf.set_text(sh[8],[['ARTINYA  ','5 dari 10 baris bergantung pada jatah melon (007), 3 pada pelepasan kiwi SP1 (010). Import lain menunggu lot (008).']])
# 9 funnel
s=S[8]; sh=s.shapes
H(s,f"Tahap 3–6: dari {P['signals']['n']} sinyal ke {len(L)} baris W46",'TAHAP 5 · CORONG','Import & Guava ke supply; stok 0 & melon subdist keluar; 14 lolos → 6 baris + 4 Perdalam, 8 antre.')
tf.chart_data(sh[7].chart,['Sinyal terdeteksi','Supply: kategori < 60% / Guava','Supply: terkirim customer < 90%','Stok SKU 0 (semangka, papaya)','Stok melon di subdist / gate merah','Lolos gate','Baris W46'],[('Series',[P['signals']['n'],40,5,10,15,14,len(L)])])
try: tf.set_text(sh[7].chart.chart_title.text_frame,['Sinyal GT Tangerang minggu 2–8 Nov (customer × kategori)'])
except Exception as e: print(e)
tf.cell(sh[8].table.cell(0,0),'Slot W46'); tf.fill_table(sh[8].table,[['Lanjutan W45 (Raksi, maks 1×)','1'],['Jaga baru (melon, Pine)',str(nj-1)],['Perdalam (HE Duta, kiwi ×3)','4'],['Antre (Jaga 6/6 · jatah melon)','8'],['Total (maks 10)',str(len(L))]],styles=Z)
tf.below(sh[9],sh[8],0.08)
tf.fill_table(sh[9].table,[['Setia, Suparmin, Kartolo +4 · melon Grosir','Jaga 6/6 · jatah melon → W47'],['Grosir Buah Segar-Segar · Pine','Belum di roster'],['Ranum · semangka','Stok semangka 0 di TGR']],styles=Z)
tf.set_text(sh[10],['Jaga = order 2–8 Nov < 70% rata-rata 5 Okt–1 Nov, rata-rata ≥ 10 BE, skor ≥ Rp 0,3 jt/bln. Melon subdist lolos hanya bila stok subdist (sell-in − sell-out sejak 20 Okt) ≈ 0.'])
tf.set_text(sh[11],[['ARTINYA  ','70 sinyal (supply 45, stok SKU 0 10, melon tertahan 15) bukan pekerjaan salesman; 10 baris (Jaga 6/6 + 4 Perdalam) + 8 SKU A2.']])
# 10 score
s=S[9]; sh=s.shapes
H(s,'Tahap 5: skor dan kenapa targetnya angka itu','TAHAP 5 · SKOR & TARGET','Target = ukuran customer (4 mgg); Ranum = sell-out outlet; Duta = 2 mgg dipotong jatah HE; kiwi = uji outlet Lampiran.')
why={'CC-2610-136':'lanjut 1×: biasa ±36 + 4','CC-2611-001':'dipotong dari ±54 ke jatah melon (sell-out ±27/mgg)','CC-2611-002':'dipotong ke jatah HE (biasa ±36)','CC-2611-003':'rata-rata Pine 4 mgg (±101)','CC-2611-004':'rata-rata melon 4 mgg (±25)','CC-2611-005':'rata-rata Pine 4 mgg (±25)','CC-2611-006':'laju 2 mgg (±100 · 170) dipotong jatah HE','CC-2611-019':'uji: outlet L3 ±10 box / 2 mgg','CC-2611-020':'uji: outlet L4 ±39 box / 2 mgg','CC-2611-021':'biasa ±17 → 40 (KEYRA ±33 / 2 mgg)'}
rows=[[sid(l['id']),l['cust'][:22],l['status'],n1(l['val']),str(l['like']).replace('.',','),n1(l['score']),l['tgt'],why[l['id']]+' · '+P['price'][l['id']]['txt']] for l in L]
t=sh[7].table; tf.fill_table(t,rows,styles=lambda i,j,r: dict(fill=('FFF2CC' if (j==2 and r[2]=='Lanjut') else ('F3F6F2' if i%2==0 else 'FFFFFF'))))
tf.set_text(sh[8],[['CATATAN  ','Kolom terakhir = dasar target · harga rek. (harga customer sendiri saat volume biasa; elastisitas bila target > biasa). Skor = nilai/bln × peluang (Jaga 0,6 · Perdalam 0,5).']])
# 11 how to win
s=S[10]; sh=s.shapes
H(s,'Tahap 6: cara menang per baris — aksi, SKU, outlet','TAHAP 6 · CARA MENANG','Aksi dari sinyal · SKU dari stok 9 Nov & S&D · harga dari analisis harga · outlet dari Lampiran ScyllaX.')
LQ={q['id']:len(q['q']) for q in P['lamp']}
E=P['gtstk']; whysku={'P':f"GT TGR: CL 10 {n0(E['cl10'])} · CR 9 {n0(E['cr9'])} · CR 5 {n0(E['cr5'])}",'M':f"melon SP2/tua {n0(E['melon'])} (HE {n0(E['he'])})",'I':f"kiwi curah {n0(E['kiwi_curah']+E['kiwi_curah_c'])} + punnet {n0(E['kiwi_pnt']+E['kiwi_pnt_c'])} (SP1 boleh)",'C':f"banana > 3 hr {n0(E['ban_aged'])}"}
rows=[[sid(l['id']),l['cust'][:22],l['why'].split(' · Grosir')[0].split(' → ')[0],f"{l['do']} · {l['prod']}",whysku[l['k2']],l['tgt'],(f"{l['lamp'].replace('Lampiran ','')} · {LQ[l['id']]} outlet" if l['lamp'].startswith('Lampiran') else 'Grosir — tanpa outlet')] for l in L]
t=sh[7].table; tf.fill_table(t,rows,styles=Z)
tf.set_text(sh[8],[['BACA  ','Baris Subdist punya Lampiran outlet (L1–L5). Grosir (Diqin, Soleram, Mamad, Apple Shop, Duta) jual langsung → tanpa outlet, dinilai dari order BI.']])
# 12 example card: 001 Ranum
s=S[11]; sh=s.shapes; l=next(x for x in L if x['id']=='CC-2611-001')
H(s,'Contoh: CC-2611-001 Ranum — melon SP2 / tua','TAHAP 3–9 · 1 BARIS','Satu baris lewat 9 field Command Card: dari stok subdist & sell-out outlet ke SKU, harga, dan outlet Lampiran.')
tf.set_text(sh[7],['Cabang',f"Jatah GT TGR melon {n0(P['gtstk']['melon'])} box (> 3 hr + SP2, 9 Nov)"])
tf.set_text(sh[9],['Customer','Ranum melon W45 0 BE vs biasa ±54'])
tf.set_text(sh[11],['Bukti outlet','Sell-out 2–8 Nov 20 box; stok subdist ≈ 0 (−8)'])
tf.set_text(sh[13],['Harga','HE SP GT ±319 rb W45 → tahan; harga rek. melon ±290 rb (007)'])
tf.set_text(sh[15],['Cara menang','Melon > 3 hr / SP2 ke outlet pembeli (L2) @ ±290 rb'])
rows=[['1','Sinyal','Order melon W45 0 BE vs rata-rata 4 mgg ±54 BE'],['2','Bukti','Sell-in − sell-out sejak 20 Okt ≈ −8 box; outlet beli 20 box 2–8 Nov (Bu Wanto 7)'],
 ['3','Sebab','Golden (SKU utama) habis; jatah GT TGR hanya > 3 hr / SP2 — bukan permintaan turun'],['4','Dampak',f"±Rp {n1(l['val'])} jt/bln × peluang 0,6 = skor Rp {n1(l['score'])} jt"],
 ['5','Command','Melon SP2 / tua (Golden, HE, Korean) untuk outlet pembeli melon (L2)'],['6','Pemilik & batas','Achmad Farchan · ASM GT TGR · W46 (9–15 Nov)'],
 ['7','Pemutus & band','ASM (Band A); alokasi & harga HE = PM Local (007)'],['8','Baseline','Melon Ranum ±54 BE/mgg; sell-out outlet ±27 box/mgg (20 Okt–8 Nov)'],
 ['9','Test','Met: order melon ≥ 40 BE dan sell-out ≥ 50% order dalam 7 hari · Invalid bila jatah melon GT TGR 0']]
tf.cell(sh[16].table.cell(0,2),'Isi untuk CC-2611-001 · Ranum'); tf.fill_table(sh[16].table,rows,styles=Z)
tf.set_text(sh[17],[['ARTINYA  ','Baris melon subdist W46 hanya Raksi & Ranum (stok subdist habis di outlet); SKU dari jatah GT TGR, bukan total gudang.']])
# 13 review & calendar
s=S[12]; sh=s.shapes
H(s,'Tahap 7–9: keputusan Senin dan jadwal sampai verdict W46','TAHAP 7–9 · REVIEW',f"{len(SUP)} command supply/harga, 9 penutupan W45 — lalu kode hasil s/d Senin 16 Nov 09:00.")
rows=[[sid(x['id']),x['typ'],x['owner'],x['cmd'][:120]+('…' if len(x['cmd'])>120 else ''),x['test']] for x in SUP]
t=sh[7].table; tf.fill_table(t,rows,styles=Z); tf.est_heights(t)
tf.fill_table(sh[8].table,[['136 Raksi (Partial → lanjut 1×)','L1: melon SP2/tua ke outlet pembeli; terakhir kali'],['147 Harum HE (Partial)','Tutup; sell-out 3 box/mgg, stok ±15'],['Aturan jatah GT TGR (> 3 hr + SP2/SP3)','Tulis di Engine Rules v3.1; Import perlu kolom umur'],['Kopkar roster (Bahar / Ryan)','Konfirmasi pemilik akun']],styles=Z)
tl=[(10,11,'Sen 9 Nov','Verdict W45 final · sheet final · WCR 13:30'),(13,14,'Sel 10 Nov','Sheet + A2 dibagikan · alokasi HE GT TGR (007)'),(16,17,'Rab 11 Nov','DS Century (008) · DS papaya (009)'),(19,20,'Jum 13 Nov','Draft W47 · cek kiwi 019–021 (010)'),(22,23,'Sen 16 Nov','Cut-off kode W46 · verdict')]
for a,b,d,txt in tl: tf.set_text(sh[a],[d]); tf.set_text(sh[b],[txt])
# move table2 + timeline below table1 if needed
y=sh[7].top+tf.table_height(sh[7].table)+int(0.1*914400)
if y>sh[8].top:
    dy=y-sh[8].top; tf.shift([sh[k] for k in [8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23]],dy)
tf.set_text(sh[24],[['SETELAH REVIEW  ','Sheet + register (keputusan & alasan) setelah WCR → pre-read ASM/BM → dibagikan Selasa 10 Nov.']])
p.save(OUT+'WCR_W46_Backing_Analysis_Tangerang.pptx'); print('ok')
```

## build_pricing.py
```python
from pptx import Presentation
import tfill as tf, pandas as pd, pickle, numpy as np
from common import *
pr=pickle.load(open('pr.pkl','rb')); R=pickle.load(open('pr2.pkl','rb'))
p=Presentation(A+'b990c0f1-Pricing_RunRate_W45_Tangerang.pptx'); S=p.slides; St=P['stk']; C=P['cover'].set_index('sku')
Z=lambda i,j,r: dict(fill='F3F6F2' if i%2==0 else 'FFFFFF')
def H(sl,title,kick,body): tf.set_text(sl.shapes[3],[title]); tf.set_text(sl.shapes[5],[[kick+'   ',body]])
def rb(v): return f"{v/1000:.0f} rb"
he=pr['he_gt'].loc[:'2026-11-02']; hm=pr['he_mt']
# 1 cover
sh=S[0].shapes
tf.set_text(sh[19],['GT Tangerang · bahan Command W46','Senin 9 Nov 2026 · bersama notulen S&D 30 Okt']); tf.set_text(sh[20],['Stok SAP TGR 9 Nov','Harga = Net Sales ÷ BE']); tf.set_text(sh[21],['CC Lead · Commercial Analyst','BI realisasi Agt–Nov 2026'])
# 2 summary
s=S[1]; sh=s.shapes
H(s,'HE: harga naik, order tetap naik → tahan harga; Import kini soal lot','RINGKASAN','HE ketat (permintaan > jatah); kiwi ±28 mgg cover; lot Import GT kosong; MT menyerap pear di harga lebih tinggi.')
tf.set_text(sh[7],['HE: harga ↑, order ↑','+16%','order SP GT 365 BE di ±319 rb (dari 286 rb)'])
tf.set_text(sh[8],['Kiwi SZ30 cover','28 mgg',f"{n0(St['kiwi30'])} box vs {n0(C.loc['Sungold Kiwi SZ30','tgr'])}/mgg"])
tf.set_text(sh[9],['Import GT W45','718 BE','dari 4.915 (W44) — lot habis'])
tf.set_text(sh[10],['Pear MT W45','±490–535 rb','Golden & Singo (dari ±435 rb) — volume tetap'])
tf.set_text(sh[11],['Golden / Rock','59 / 48','box TGR 9 Nov — MT ambil 500 / 284 BE'])
tf.set_text(sh[12],['TEMUAN',
 "•  HE: harga SP GT naik ±286 → ±319 rb (+12%) tetapi order SP 314 → 365 BE (+16%); semua HE GT 468 → 607 BE. Golden GT habis (861 → 243 BE) → pembeli pindah ke HE.",
 "•  Pear: MT menyerap Golden & Singo 700–900 BE/mgg walau harga naik ±435 → ±490–535 rb (W45). GT tetap tidak membeli pear di 37–54 rb/kg; Century (±17 rb/kg) 0 sejak 20 Okt.",
 f"•  Kiwi: stok Sungold SZ30 {n0(St['kiwi30'])} box (±28 mgg). Order MT SZ30 turun W45 (381 vs ±700–950 BE). GT ORI ±513 rb vs MT ±565 rb — selisih kecil, GT tidak bergerak tanpa lot.",
 "•  Import GT W45 718 BE: Longan, Red Globe, Shantang, Century = 80% Import GT 4 mgg terakhir — semuanya lot musiman, kini hampir 0."])
tf.set_text(sh[13],['KEPUTUSAN HARGA','•  PM Local: HE SP GT tahan ±310–320 rb + HE > 3 hr ke GT TGR ≥ 300 box (007)','•  Import: tanggal DS Century; pear tetap MT (008)','•  PM Import + MT: kiwi ORI → MT; SP1 ke GT ±440–500 rb (010)','•  Pine, Guava, Dazzle: tanpa tindakan harga'])
tf.set_text(sh[14],[['ARTINYA  ','HE ±319 rb tidak menahan order karena Golden habis: tahan harga, batasnya alokasi. Untuk Import, yang hilang adalah barangnya, bukan harganya.']])
# 3 cover table
s=S[2]; sh=s.shapes
H(s,'Stok 9 Nov vs run-rate: kiwi surplus; melon, Pine, Guava ketat','RUN-RATE','Cover = stok ÷ run-rate TGR semua channel 4 mgg (5 Okt–1 Nov). Fresh normal < 1 mgg; Import > 6 mgg = risiko downgrade.')
notes={'Sungold Kiwi SZ30':'order nas. SZ30 turun W45','Sungold Kiwi SZ27':'MT 663 BE W45','Green Kiwi SZ30':'ikut program kiwi','Singo Pear':'MT ±920/mgg di ±489 rb','Golden Pear':'MT ±720/mgg; stok tipis','Dazzle Apple':'habis W47 (S&D)','Honey Emerald':'GT 607 BE W45 — alokasi','Golden Melon':'MT ambil 500 BE W45','Rock Melon':'MT ambil 284 BE W45','Honi Pine CL 8':'jual dulu (56 tua)','Honi Pine CL 10':'GT utama','Honi Pine CL 15':'MT utama','Honi Pine CR 9':'187 tua; GT tidak beli','Crystal Guava':'±54 ton/mgg, kurang'}
def stat(r):
    if r.cover>=6: return 'Surplus'
    if r.stk<=60: return 'Ketat'
    if r.aged>=0.15*r.stk and r.aged>50: return 'Umur'
    return 'Seimbang' if r.cover>=1 else 'Ketat'
rows=[]
for k in ['Sungold Kiwi SZ30','Sungold Kiwi SZ27','Green Kiwi SZ30','Singo Pear','Golden Pear','Dazzle Apple','Honey Emerald','Golden Melon','Rock Melon','Honi Pine CL 8','Honi Pine CL 10','Honi Pine CL 15','Honi Pine CR 9']:
    r=C.loc[k]; cv=f"{r.cover:.0f} mgg" if r.cover>=1.5 else ('< 1 mgg' if r.cover<1 else f"{n1(r.cover)} mgg")
    rows.append([k,n0(r.stk),n0(r.aged),n0(r['gt']),n0(r['mt']),n0(r.tgr),n0(r.o45),cv,notes.get(k,'') or '—',stat(r)])
t=sh[7].table; tf.cell(t.cell(0,6),'Order W45'); tf.cell(t.cell(0,8),'Catatan (S&D 30 Okt / W45)')
SC={'Surplus':'FDE9E7','Ketat':'FFF2CC','Umur':'FFF2CC','Seimbang':'E6F0E2'}
tf.fill_table(t,rows,styles=lambda i,j,r: dict(fill=SC[r[9]]) if j==9 else Z(i,j,r))
tf.set_text(sh[8],['Run-rate = Net Sales BE BI, gudang Tangerang semua channel. Order W45 = So BE 2–8 Nov TGR semua channel (kiwi SZ30+SZ27 971 BE, MT 729).'])
tf.set_text(sh[9],[['ARTINYA  ',f"Kiwi (3 ukuran) {n0(St['kiwi30']+St['kiwi27']+St['green30'])} box tetap masalah stok terbesar. Melon & Pine tidak perlu didorong — HE perlu alokasi, bukan diskon."]])
# 4 HE
s=S[3]; sh=s.shapes
H(s,'HE: harga ±319 rb, order GT tetap naik → tahan harga','HARGA MELON','GT Tangerang HE SP: ±286 rb & 314 BE (W44) → ±319 rb & 365 BE (W45). MT nasional ±433 rb.')
hh=he.iloc[-10:]; mm=hm.reindex(hh.index)
cats=[f"{d.day}/{d.month}" for d in hh.index]
tf.chart_data(sh[7].chart,cats,[('GT TGR (SP)',[round(v/1000) for v in hh.rp]),('MT nasional (SP)',[None if pd.isna(v) else round(v/1000) for v in mm.rp])])
t=sh[9].table; tf.cell(t.cell(0,1),'W44'); tf.cell(t.cell(0,2),'W45')
w43,w44,w45=he.loc['2026-10-19'],he.loc['2026-10-26'],he.loc['2026-11-02']
tf.fill_table(t,[['Harga HE SP',rb(w44.rp),rb(w45.rp)],['Order / mgg (BE)',n0(w44.so),n0(w45.so)],['Customer membeli',str(int(w44.n)),str(int(w45.n))],['Order vs minggu lalu',f"{(w44.so/w43.so-1)*100:+.0f}% (harga {(w44.rp/w43.rp-1)*100:+.0f}%)".replace('-','−'),f"{(w45.so/w44.so-1)*100:+.0f}% (harga {(w45.rp/w44.rp-1)*100:+.0f}%)"]],styles=Z)
tf.fill_table(sh[8].table,[['Pesaing (Sweetnet / HE, catatan Ranum)','285–315 rb','19–21 rb'],['SSN HE SP GT — W44','286 rb','19 rb'],['SSN HE SP GT — W45',rb(w45.rp),f"{w45.rp/15000:.0f} rb"],['SSN HE SP MT nasional — W45',rb(hm.loc['2026-11-02'].rp),f"{hm.loc['2026-11-02'].rp/15000:.0f} rb"]],styles=Z)
tf.set_text(sh[10],['BACAAN',"•  W45 harga SP GT ±319 rb (±21 rb/kg) = batas atas pesaing; order SP tetap naik 16% karena Golden GT habis (861 → 243 BE).",
 "•  Pembeli HE SP 30 → 22, tetapi semua HE GT 468 → 607 BE — volume terkumpul di pembeli besar (Duta 170, Sastro 119). Bacaan \"order −15%\" sebelumnya salah: order 6–8 Nov masuk terlambat.",
 "•  Jatah GT TGR (9 Nov, > 3 hr + SP2): HE 230 box + 304 SP1 segar yang menua minggu ini → 5 baris melon (235 BE)."])
tf.set_text(sh[11],[['REKOMENDASI  ','Tahan harga HE SP GT ±310–320 rb (jangan diskon); harga rek. per customer ±290–320 rb; jatah HE > 3 hr ≥ 300 box untuk GT TGR (007).']])
# 5 pear / import
s=S[4]; sh=s.shapes
H(s,'Import GT: lot habis; pear tetap barang MT','IMPORT & PEAR','Import GT TGR W45 718 BE (W44 4.915). Longan, Red Globe, Shantang, Century = 80% Import GT 4 mgg — lot musiman.')
rows=[[a,b,rb(c),str(kg),f"{c/kg/1000:.0f} rb",f"{(c/kg)/(R['pear'][0][2]/15):.1f}×".replace('.',',')] for a,b,c,_,kg in R['pear']]
t=sh[7].table; tf.cell(t.cell(0,0),'Produk (realisasi 13 mgg)'); tf.fill_table(t,rows,styles=Z)
mt=dict(R['mtpear']); wk=[pd.Timestamp('2026-10-12'),pd.Timestamp('2026-10-19'),pd.Timestamp('2026-10-26'),pd.Timestamp('2026-11-02')]
t=sh[8].table
for j,w in enumerate(wk,1): tf.cell(t.cell(0,j),f"{w.day}/{w.month}")
def g(nm,w,c):
    a=mt[nm]; return (n0(a.loc[w,c]) if c=='so' else (rb(a.loc[w,'rp']) if not pd.isna(a.loc[w,'rp']) else '—')) if w in a.index else '—'
tf.fill_table(t,[['Golden Pear BE']+[g('Golden Pear',w,'so') for w in wk],['Golden Pear rb/box']+[g('Golden Pear',w,'rp') for w in wk],['Singo Pear BE']+[g('Singo Pear',w,'so') for w in wk],['Singo Pear rb/box']+[g('Singo Pear',w,'rp') for w in wk]],styles=Z)
T=R['impfam']; cols=[c for c in T.columns if c!='tot']
t=sh[9].table; tf.cell(t.cell(0,0),'Import GT TGR (BE)'); tf.cell(t.cell(0,1),'3 mgg 12 Okt–1 Nov'); tf.cell(t.cell(0,2),'W45')
rows=[[f,n0(T.loc[f,cols[:-1]].sum()),n0(T.loc[f,cols[-1]])] for f in T.index[:8]]
tf.fill_table(t,rows,styles=Z)
tf.set_text(sh[10],['BACAAN','•  MT menyerap Golden & Singo 700–900 BE/mgg walau harga naik ±435 → ±490–535 rb → "harga terbaik" pear sudah diuji di MT.','•  GT: 4 lot musiman membawa ±11.300 BE dalam 4 mgg, lalu hampir 0 (W45 230 BE). Tanpa kalender lot, Import GT naik-turun ikut kedatangan barang.','•  Kiwi satu-satunya Import yang tetap jalan: GT 758 BE dalam 4 mgg.'])
tf.set_text(sh[11],[['REKOMENDASI  ','Jangan diskon pear untuk GT. Import: tanggal DS Century + kalender lot Red Globe / Longan / Jeruk untuk GT TGR (008).']])
# 6 kiwi
s=S[5]; sh=s.shapes
H(s,'Kiwi: ±28 mgg stok; order MT melambat di W45','KIWI',f"TGR 9 Nov: Sungold SZ30 {n0(St['kiwi30'])}, SZ27 {n0(St['kiwi27'])}, Green SZ30 {n0(St['green30'])} box. SZ30 nasional W45: GT 94 · MT 381 BE.")
K=R['kiwi'].iloc[-10:]
tf.chart_data(sh[7].chart,[f"{d.day}/{d.month}" for d in K.index],[('GT',[round(v) for v in K['General Trade']]),('MT',[round(v) for v in K['Modern Trade']])])
tf.fill_table(sh[8].table,[[a,rb(b),n0(c)] for a,b,c in R['kiwip']],styles=Z)
tf.fill_table(sh[9].table,[['minggu 27/7','1.731','460 rb'],['minggu 10/8','2.087','485 rb'],['17 Agt – 1 Nov','0','—'],['minggu 2/11 (W45)','15','±456 rb'],['Sastro W45 (kiwi semua)','147','±452 rb']],styles=Z)
tf.set_text(sh[10],['BACAAN','•  Setia kembali order (15 BE, 5 Nov) setelah 11 minggu — volume lot (1.700–2.100 BE) belum kembali.','•  MT SZ30 turun 701 → 381 BE (W44 → W45) walau program berjalan; order TGR SZ30+SZ27 971 BE ditopang SZ27 (MT 663).','•  GT ORI ±513 rb vs MT ±565 rb: selisih ±9% → diskon GT tambahan tidak menyelesaikan 19 rb box.'])
tf.set_text(sh[11],[['REKOMENDASI  ','Order TGR SZ30+SZ27 ≥ 900 BE/mgg (010). Kiwi SP1 dibuka ke GT TGR: SP1 curah ±440 rb, punnet SP1 ±500 rb → Harum, Ranum, Go Banana (019–021).']])
# 7 recommended price per line + A2 perish-risk
s=S[6]; sh=s.shapes
H(s,'Harga rekomendasi per baris W46 & stok berisiko','HARGA REK.','Baris: harga customer sendiri (5 Okt–1 Nov) ± elastisitas. A2: harga yang habiskan stok sebelum busuk.')
for k in (10,7):
    e_=sh[k]._element; e_.getparent().remove(e_)
sh=s.shapes
tl=[x for x in sh if getattr(x,'has_table',False) and x.has_table]
t1,t2=tl[0],tl[1]
t1.left=int(0.25*914400); t1.top=int(1.12*914400); t2.left=int(5.0*914400); t2.top=int(1.12*914400)
tf.set_widths(t1.table,[1.6,1.05,1.95]); tf.set_widths(t2.table,[1.75,0.95,2.05])
for j,h in enumerate(['Baris W46','Harga rek. / box','Dasar (harga ref · e)']): tf.cell(t1.table.cell(0,j),h)
for j,h in enumerate(['A2 stok berisiko (jatah GT)','Harga rek.','Tawarkan ke']): tf.cell(t2.table.cell(0,j),h)
def basis(l):
    d=P['price'][l['id']]['detail'][0]
    if l['k2']=='I': return 'SP1 = tengah SP2/ORI GT TGR · e −1,0 (stok kiwi lebih)'
    e=d['e']; es='e n/a' if e is None else f"e {e:.1f}".replace('.',',')
    return f"{d['src'].replace('harga sendiri','own')} {d['pref']/1000:.0f} rb · {es}"
rows=[[f"{sid(l['id'])} {l['cust'][:16]}",P['price'][l['id']]['txt'],basis(l)] for l in P['lines']]
tf.fill_table(t1.table,rows,styles=Z); tf.est_heights(t1.table,k=0.58)
rows=[[f"{a['lab'][:30]} · {n0(a['q'])}",a['price'],a['to'][:46]] for a in P['A2']]
tf.fill_table(t2.table,rows,styles=lambda i,j,r: dict(fill=('FDE9E7' if i==0 else ('F3F6F2' if i%2==0 else 'FFFFFF')))); tf.est_heights(t2.table,k=0.58)
tf.set_text([x for x in sh if x.name=='Rectangle 269'][0],[['REKOMENDASI  ',f"Pakai harga rek. di sheet (kolom baru). HE ±290–320 rb (tanpa diskon); kiwi SP1 ±440–500 rb. A2: ±Rp {n1(P['A2_val']/1e9)} M stok berisiko — Pandan SP2 & Sweet Net C terbesar; Cavendish di luar."]])
# 8 decisions
s=S[7]; sh=s.shapes
H(s,'Keputusan harga & supply untuk WCR Senin 9 Nov','KEPUTUSAN','Tiap keputusan punya pemilik, angka uji, dan baris salesman yang bergantung padanya.')
prod={'CC-2611-007':'Honey Emerald','CC-2611-008':'Century / lot Import','CC-2611-009':'Papaya California','CC-2611-010':'Kiwi','CC-2610-116':'Banana'}
rows=[[sid(x['id']),prod[x['id']],x['cmd'],x['prob'][:150]+('…' if len(x['prob'])>150 else ''),x['owner'],x['test'],x['lines']] for x in P['sup']]
t=sh[7].table; tf.cell(t.cell(0,6),'Baris W46'); tf.fill_table(t,rows,styles=Z)
tf.set_text(sh[8],[['ATURAN  ','Harga diputuskan PM, bukan salesman. Salesman menulis kode H + harga pesaing → bahan analisis harga minggu berikutnya.']])
p.save(OUT+'Pricing_RunRate_W46_Tangerang.pptx'); print('ok')
```

## build_snap.py
```python
from pptx import Presentation
import tfill as tf, pandas as pd
from common import *
p=Presentation(A+'2d395683-Snapshot_Cabang_W45_GT_Tangerang.pptx'); sh=p.slides[0].shapes
perf=P['perf']; tot=perf[-1]; S=P['stk']; c=P['cnt45']
tf.set_text(sh[3],[["GT Tangerang · Minggu 46 (9–15 November 2026)"],["Revenue Sep–Okt · terkirim 2–5 Nov · stok SAP 9 Nov · S&D 30 Okt · dibaca bersama Command Mingguan W46"]])
tf.set_text(sh[4],[["Intinya.  ",f"Okt {tot['Okt']/tot['bbpOkt']*100:.0f}% BBP (Sep {tot['Sep']/tot['bbpSep']*100:.0f}%). Banana terkirim 38% (Cavendish di luar command W46). Import GT W45 jatuh ke 718 BE (W44 4.915): lot habis (Century, Shantang, Red Globe, Longan) — kiwi satu-satunya yang berlimpah (±28 mgg). Jatah GT TGR (umur > 3 hr + SP2/SP3, SAP 9 Nov): melon ±267 box (HE: harga ditahan ±320 rb — order GT naik walau harga naik); 8 SKU stok berisiko di A2."]])
tf.set_text(sh[5],['Rev September',M_:=f"{n1(tot['Sep']/1e9)} M",f"{tot['Sep']/tot['bbpSep']*100:.0f}% BBP ({n1(tot['bbpSep']/1e9)} M)"])
tf.set_text(sh[6],['Oktober',f"{n1(tot['Okt']/1e9)} M",f"{tot['Okt']/tot['bbpOkt']*100:.0f}% BBP ({n1(tot['bbpOkt']/1e9)} M)"])
tf.set_text(sh[7],['Terkirim 2–5 Nov',f"{tot['fill']*100:.0f}%",'semua buah GT'])
tf.set_text(sh[8],['Banana terkirim',f"{perf[0]['fill']*100:.0f}%",f"stok C3 13KG {n0(S['c313'])} (9 Nov)"])
tf.set_text(sh[9],['Hasil W45',f"{c['Met']} Met",f"{c['Partial']} Partial · {c['Failed']} Failed · 0 Open"])
tf.set_text(sh[10],[['1  PERFORMA PER KATEGORI  ','Sep & Okt vs BBP, dan siapa yang menangani gap-nya']])
route={'Banana':'Terkirim 38% — Cavendish di luar command W46','Pineapple':'Di atas BBP; CR 7 habis → Soleram 003, Apple Shop 005','Import':'Okt OK, Nov jatuh: lot habis → 008; kiwi SP1 → 019–021','Guava':'Supply ±54 ton masih kurang → B, tanpa baris','Local':'Melon jatah GT ±267 box → 5 baris + 007; papaya 009','Total buah':f"Gap Okt −{n1((tot['bbpOkt']-tot['Okt'])/1e9)} M"}
rows=[[r['kat'],f"{n1(r['Sep']/1e9)} M",f"{n1(r['bbpSep']/1e9)} M",f"{r['Sep']/r['bbpSep']*100:.0f}%",f"{r['Okt']/r['bbpOkt']*100:.0f}%",f"{r['fill']*100:.0f}%",route[r['kat']]] for r in perf]
t=sh[11].table; tf.cell(t.cell(0,5),'Terkirim')
def pc(v): return float(v.rstrip('%'))
def sty(i,j,r):
    d=dict(fill='F3F6F2' if i%2==0 else 'FFFFFF',bold=(r[0]=='Total buah'),color='333333')
    if j in(3,4): d['color']='C00000' if pc(r[j])<80 else ('04582D' if pc(r[j])>=100 else '333333')
    if j==5: d['color']='C00000' if pc(r[j])<60 else '333333'
    return d
tf.fill_table(t,rows,styles=sty)
tf.set_text(sh[12],[['2  STOK & SUPPLY  ','SKU kelebihan stok · estimasi kurang kirim · AO / RO subdist']])
tf.set_text(sh[13],['Kelebihan / risiko stok — gudang TGR (SAP 9 Nov)'])
C=P['cover'].set_index('sku')
def cv(x): return f"±{x:.0f} mgg" if x>=1.5 else ('< 1 mgg' if x<1 else f"±{n1(x)} mgg")
ov=[('Sungold Kiwi SZ30','Sungold Kiwi SZ30'),('Green Kiwi SZ30','Green Kiwi SZ30'),('Sungold Kiwi SZ27','Sungold Kiwi SZ27'),('Singo Pear','Singo Pear'),('Cav C3 5KG','Cav C3 5KG'),('Cav FK (220 tua)','Cav FK')]
rows=[[nm,n0(C.loc[k,'stk']),n0(C.loc[k,'tgr']),cv(C.loc[k,'cover']),n0(C.loc[k,'aged'])] for nm,k in ov]
tf.fill_table(sh[14].table,rows,styles=lambda i,j,r: dict(fill='F3F6F2' if i%2==0 else 'FFFFFF'))
tf.set_text(sh[15],['Estimasi kurang kirim — GT 2–5 Nov'])
rows=[[r['grp'],n0(r['order']),f"{r['fill']*100:.0f}%",n0(r['short']),(f"{n1(r['short_rp']/1e9)} M" if r['short_rp']>=1e8 else f"{n0(r['short_rp']/1e6)} jt")] for r in P['short']]
tf.fill_table(sh[16].table,rows,styles=lambda i,j,r: dict(fill='F3F6F2' if i%2==0 else 'FFFFFF'))
st=P['short_tot']
tf.set_text(sh[17],[f"Total GT tak terkirim 2–5 Nov {n0(st['be'])} BE ≈ Rp {n1(st['rp']/1e9)} M (batas atas — termasuk order ulang); hampir semua banana (terkirim 38%). Cover = stok ÷ run-rate TGR semua channel 5 Okt–1 Nov. Melon, Pine, Guava < 1 mgg (normal fresh)."])
AO=P['aoro'].head(8)
t=sh[18].table
hdr=['Subdist (8 RO terbesar)']+[d.replace(' Nusantara Mandiri','').replace(' Buah Segar','').replace('Buah Segar Nusantara','BS Nusantara').replace('Depok City Cau','DCC').replace('Sinar Buah Banten','Sinar Banten').replace('Citra Buah Segar - Tgr','Citra BS').replace(' Buah','').replace('Batavia Fresh Fruit','Batavia').replace('Bogor Raya Segar','Bogor Raya').replace('Krista Jaya Internasional','Krista Jaya')[:12] for d in AO.dist]+['Total TGR (25)']
A_=P['aoro']
for j,h in enumerate(hdr): tf.cell(t.cell(0,j),h)
rows=[['RO (outlet terdaftar aktif)']+[n0(v) for v in AO.ro]+[n0(A_.ro.sum())],
      ['AO Sep']+[n0(v) for v in AO.ao_sep]+[n0(A_.ao_sep.sum())],
      ['AO Okt']+[n0(v) for v in AO.ao_okt]+[n0(A_.ao_okt.sum())],
      ['AO / RO Okt']+[f"{v*100:.0f}%" for v in AO.r]+[f"{A_.ao_okt.sum()/A_.ro.sum()*100:.0f}%"]]
for i,r in enumerate(rows):
    for j,v in enumerate(r): tf.cell(t.cell(1+i,j),v)
low=A_[A_.ro>=50].sort_values('r').head(3)
tf.set_text(sh[19],[f"AO = outlet dengan ≥ 1 transaksi ScyllaX (data interim, produk SSN); RO = master outlet aktif. AO/RO TGR {A_.ao_okt.sum()/A_.ro.sum()*100:.0f}%; terendah: "+', '.join(f"{d} {v*100:.0f}%" for d,v in zip(low.dist,low.r))+". ScyllaX berhenti: Ayotama (17 Okt), Batavia (25 Okt), Mandiri (31 Okt), Bandar (1 Nov)."])
tf.set_text(sh[20],[['3  CUSTOMER DI COMMAND MINGGUAN W46  ',f"{len(P['lines'])} baris · 1 lanjutan (hasil W45 dalam kurung) · {len(P['lines'])-1} baru"]])
KN={'M':'Melon','P':'Pineapple','I':'Import','C':'Banana'}
why={'CC-2610-136':'lanjut 1× ke outlet L1','CC-2611-001':'dipotong ke jatah melon (outlet L2)','CC-2611-002':'HE ganti Rock; dipotong jatah HE','CC-2611-003':'= biasa (CR 9 / CR 5 ganti CR 7)','CC-2611-004':'= biasa (HE ganti Rock)','CC-2611-005':'= biasa (CL 10)','CC-2611-006':'laju 2 mgg, dipotong jatah HE','CC-2611-019':'uji kiwi SP1 ke outlet L3','CC-2611-020':'uji kiwi SP1 ke outlet L4','CC-2611-021':'naikkan kiwi (KEYRA, L5)'}
rows=[[str(k),l['cust'],KN[l['k2']],('Lanjut (Partial)' if l['status']=='Lanjut' else 'Baru'),f"{n0(l['b4'])} → {n0(l['wk'])}",l['tgt'],P['price'][l['id']]['txt']+' · '+why[l['id']]] for k,l in enumerate(P['lines'],1)]
t=sh[21].table; tf.cell(t.cell(0,4),'Order BE/mgg: biasa → W45'); tf.cell(t.cell(0,6),'Harga rek. · kenapa target ini')
tf.fill_table(t,rows,styles=lambda i,j,r: dict(fill=('FFF2CC' if r[3].startswith('Lanjut') else ('F3F6F2' if i%2==0 else 'FFFFFF'))))
h_=tf.est_heights(t,k=0.5,minh=0.17); tf.below(sh[22],sh[21],0.05); print('tbl end',round((sh[21].top+h_)/914400,2),'pesan end',round((sh[22].top+sh[22].height)/914400,2))
tf.set_text(sh[22],[['PESAN UNTUK SALES'],
 ['•  ','Jatah GT TGR = umur > 3 hr + SP2/SP3 (kiwi: SP1 juga). Pakai harga rek. per baris; A2 di hal. 4. Cavendish: tanpa command.'],
 ['•  ','HE: jangan diskon (order naik di ±319 rb). Pine CR 7 habis → CR 9 / CR 5 / CL 10. Kiwi SP1 → Sastro, Harum, Ranum, Go Banana.']])
sh[22].height=int(0.55*914400); sh[23].top=sh[22].top+sh[22].height
tf.set_text(sh[23],["Sumber: BI interim v1.4.2 (6) & BBP 2026 (RAW_TARGET) · SAP stok TGR 9 Nov · S&D 30 Okt · ScyllaX s/d 8 Nov · master outlet. Run-rate = Net Sales BE 5 Okt–1 Nov semua channel TGR."])
p.save(OUT+'Snapshot_Cabang_W46_GT_Tangerang.pptx'); print('ok')
```

## build_reg.py
```python
import openpyxl, copy
from openpyxl.styles import Font
from common import *
wb=openpyxl.load_workbook(A+'cecc1469-CC_Command_Register_W45_Tangerang.xlsx')
def refill(ws, title, sub, header, rows, title_row=1):
    ws.title=title
    sty=[copy.copy(ws.cell(5,c)._style) for c in range(1,max(len(header),ws.max_column)+1)]
    hsty=[copy.copy(ws.cell(4,c)._style) for c in range(1,max(len(header),ws.max_column)+1)]
    for r in range(4, ws.max_row+1):
        for c in range(1, ws.max_column+1): ws.cell(r,c).value=None
    ws.cell(1,1).value=sub[0]; ws.cell(2,1).value=sub[1]
    for j,h in enumerate(header,1): ws.cell(4,j).value=h; ws.cell(4,j)._style=hsty[min(j-1,len(hsty)-1)]
    for i,row in enumerate(rows):
        for j,v in enumerate(row,1):
            c=ws.cell(5+i,j); c.value=v; c._style=sty[min(j-1,len(sty)-1)]
for ws in wb.worksheets: print(ws.title, ws.max_row)
w=wb.worksheets
# W45 review
refill(w[0],'W45_REVIEW',('Review Command W45 FINAL — GT Tangerang (Senin 9 Nov)','BI order 2–8 Nov (terbooking s/d Min 8 Nov); ScyllaX sell-out 2–8 Nov; stok SAP 9 Nov. Dinilai pada produk command; kategori dicatat bila berbeda.'),
  ['ID','Customer','Target','Hasil W45','Status','W46','Catatan'],
  [[v['id'],v['cust'],f"{v['prod']} {v['target']}",v['res'],v['verdict'],v['dec'],v['note']] for v in P['rev']])
refill(w[1],'W45_SUPPLY_REVIEW',('Review command supply / data W45','Command bukan-salesman: dinilai dari test-nya.'),
  ['ID','Pemilik','Test','Hasil','Status','W46'],[list(r) for r in P['revsup']])
rows=[]
for v in P['rev']:
    if v['verdict'] not in ('Met','Partial'): continue
    rows.append([v['id'],v['cust'],v['kat'],v['verdict'],round(v['wk'],1),round(v['b8'],1),round(v['w44'],1),round(v['rp']),round(v['val']),round(max(0,v['wk']-v['b8']),1),round(v['inc']),round(v['inc']*4.3),v['so_box'],v['so_rp']])
rows.append(['Total','','','','','','','',round(sum(r[8] for r in rows)),'',round(sum(r[10] for r in rows)),round(sum(r[11] for r in rows)),'',round(sum((r[13] or 0) for r in rows))])
refill(w[2],'W45_DAMPAK_Rp',('Dampak Rp W45 — baris Met / Partial (GT Tangerang)','Nilai = BE × Rp/BE customer sendiri (5 Okt–8 Nov). Tambahan = order W45 di atas rata-rata 8 mgg (7 Sep–1 Nov). Sell-out = penjualan subdist ke outlet 2–8 Nov (bukan revenue SSN). 134/136/139/147 dihitung pada kategori melon.'),
  ['ID','Customer','Kategori','Status','Order W45 BE','Biasa BE/mgg (8 mgg)','W44 BE','Rp/BE','Nilai order W45 (Rp)','Tambahan BE','Tambahan Rp/mgg','Bila bertahan Rp/bln','Sell-out outlet box','Sell-out outlet Rp'],rows)
refill(w[3],'W46_REGISTER',('Command Register W46 — GT Tangerang','Status: Lanjut = lanjutan ke-1 · Baru = sinyal order 2–8 Nov vs 5 Okt–1 Nov. Skor = nilai/bln × peluang. ID baru memakai seri bulan November (CC-2611-nnn).'),
  ['ID','Status','Customer','Format','Outlet target','Jenis','Tipe','Lakukan','Produk','Target','Harga rek. / box','Kenapa','Detail','Salesman','Nilai Rp jt/bln','Peluang','Skor Rp jt'],
  [[l['id'],l['status'],l['cust'],l['fmt'],l['lamp'],l['rung'],l['typ'],l['do'],l['prod'],l['tgt'],P['price'][l['id']]['txt'],l['why'],l['detail'],l['sales'],round(l['val'],1),l['like'],round(l['score'],1)] for l in P['lines']])
refill(w[4],'SUPPLY_FIELD_CMD',('Command supply / harga (bukan tugas salesman)','Dari stok SAP 9 Nov, BI s/d 8 Nov, analisis harga & run-rate W46, notulen S&D 30 Okt.'),
  ['ID','Tipe','Pemilik','Masalah','Command','Test','Baris terkait'],[[c['id'],c['typ'],c['owner'],c['prob'],c['cmd'],c['test'],c['lines']] for c in P['sup']])
refill(w[5],'ANTREAN',('Antrean sinyal (slot / gate)','Sinyal Jaga minggu 2–8 Nov yang tidak masuk baris (dengan alasan). Total sinyal 84; 45 keluar ke supply (B), 10 stok SKU 0, 15 melon masih di subdist / gate merah; 14 lolos → 6 baris + 8 antre (Jaga 6/6, jatah melon). Cavendish di luar command W46.'),
  ['Customer','Kategori','Order W45 BE','Biasa BE/mgg','Skor Rp jt','SKU utama','Kenapa tidak','Salesman'],[list(r) for r in P['antre']])
a2=wb.copy_worksheet(w[5]); hr=wb.copy_worksheet(w[5]); la=wb.copy_worksheet(w[5])
refill(a2,'A2_STOK_BERISIKO',('A2 — Habiskan stok berisiko (jatah GT TGR, SAP 9 Nov)','Berisiko = jatah GT TGR (umur > 3 hr, SP2/SP3) tidak habis dalam sisa umur simpan pada run-rate GT TGR 8 mgg. Harga rek. = harga grade di GT TGR 13 mgg × (kebutuhan/run-rate)^(1/e), batas bawah = harga terendah historis grade. Di luar 10 slot.'),
  ['ID','SKU','Stok box (GT TGR)','Run-rate GT TGR /mgg','Harga ref Rp/box','Harga rek.','Elastisitas','Kenapa','Tawarkan ke','Salesman','Batas','Nilai stok Rp'],
  [[a['id'],a['lab'],round(a['q']),round(a.get('rr',0) or 0,1),round(a.get('pref',0) or 0),a['price'],a.get('e',''),a['why'],a['to'],a['sales'],a['due'],round(a['val'])] for a in P['A2']])
rows=[]
for l in P['lines']:
    for d in P['price'][l['id']]['detail']:
        rows.append([l['id'],l['cust'],d['gr'],round(d['pref']),d['src'],None if pd.isna(d['p45']) else round(d['p45']),d['e'] if d['e'] is not None else 'n/a',d['enote'],l['tgt'],d['prec']])
refill(hr,'HARGA_REK',('Harga rekomendasi per baris W46 — metode','P_rec = harga realisasi customer sendiri (grade utama) saat volume biasa 5 Okt–1 Nov; bila target > volume biasa: P_rec = P_ref × (target / biasa)^(1/e), dibatasi harga terendah 26 mgg grade itu (tidak di bawah harga sendiri bila sudah lebih rendah). e dari regresi mingguan GT TGR bila kredibel, selain itu prior. HE: dinaikkan ke harga pasar GT W45 (±319 rb) karena order naik di harga itu — harga ditahan, tanpa diskon.'),
  ['ID','Customer','Grade','Harga ref Rp/box','Sumber harga ref','Harga pasar GT TGR W45','Elastisitas e','Dasar elastisitas','Target','Harga rek. Rp/box'],rows)
refill(la,'LA2_OUTLET',('LA2 — outlet target untuk stok berisiko (A2)','Outlet subdist GT TGR yang membeli SKU yang sama 13 mgg terakhir (ScyllaX, tanpa SHOP SALES). Target box = porsi stok A2 sesuai porsi beli outlet, maks ±3 minggu volume biasanya.'),
  ['ID A2','SKU','Subdist','Salesman','Kode outlet','Nama outlet','Tipe','Beli 13 mgg (box)','Terakhir beli','Harga rek.','Target box'],
  [[a['id'],a['lab'],o['sub'],o['sales'],o['code'],o['name'],o['typ'],round(o['box'],1),o['last'].strftime('%d %b %Y'),a['price'],o['target']] for a in P['A2'] for o in a.get('outlets',[])])
w[6].cell(1,1).value='Roster salesman GT Tangerang (sumber: daftar 3 Nov 2026)'
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None or c.has_style:
                f=copy.copy(c.font); c.font=Font(name='Poppins',size=f.size,bold=f.bold,italic=f.italic,color=f.color)
wb.save(OUT+'CC_Command_Register_W46_Tangerang.xlsx'); print('ok')
```

## build_v31.py
```python
import openpyxl, copy, datetime as dt
from common import *
wb=openpyxl.load_workbook(A+'aed166c8-CC_Operating_System_Workbook_v3.1_Register_W42-W45.xlsx')
ws=wb['Command Register']; vh=wb['Verdict History']
D9=dt.datetime(2026,11,9); D10=dt.datetime(2026,11,10); D15=dt.datetime(2026,11,15)
row={ws.cell(r,1).value:r for r in range(5,ws.max_row+1) if ws.cell(r,1).value}
V={v['id']:v for v in P['rev']}; VS={r[0]:r for r in P['revsup']}
def col(L):  # letter->index
    from openpyxl.utils import column_index_from_string as ci; return ci(L)
# ---- W45 verdicts into Command Register
for vid,v in V.items():
    r=row[vid]; ws.cell(r,col('AB')).value=v['verdict']; ws.cell(r,col('AC')).value=D9
    ws.cell(r,col('AD')).value=round(v['inc']) if v['verdict'] in('Met','Partial') else 0
    trail=ws.cell(r,col('AN')).value; trail='' if trail in(None,'—') else trail+' · '
    ws.cell(r,col('AN')).value=trail+f"W45 {v['verdict']}"
    if vid=='CC-2610-136':
        ws.cell(r,col('AE')).value=1; ws.cell(r,col('AL')).value='W46'; ws.cell(r,col('AM')).value='Aktif W46 — lanjutan ke-1 (verdict Sen 16 Nov)'
        ws.cell(r,col('AJ')).value=(ws.cell(r,col('AJ')).value or '')+'; W46: Lanjut 1× (melon SP2 / > 3 hr, 40 / mgg, Lampiran L1)'
    else:
        ws.cell(r,col('AM')).value='Ditutup — verdict W45'
for vid,rr in VS.items():
    r=row[vid]; ws.cell(r,col('AB')).value=rr[4]; ws.cell(r,col('AC')).value=D9; ws.cell(r,col('AD')).value=0
    trail=ws.cell(r,col('AN')).value; trail='' if trail in(None,'—') else trail+' · '
    ws.cell(r,col('AN')).value=trail+f"W45 {rr[4]}"
    if vid=='CC-2610-116':
        ws.cell(r,col('AL')).value='W45'
        ws.cell(r,col('AM')).value='Ditahan W46 — Cavendish dikeluarkan dari command (keputusan 9 Nov)'
        ws.cell(r,col('AJ')).value=(ws.cell(r,col('AJ')).value or '')+'; W46: ditahan — Cavendish di luar command W46'
    else: ws.cell(r,col('AM')).value='Ditutup — verdict W45'
# 129 closed: Partial (already filled via V) ; 
# ---- New W46 rows
tmpl=row['CC-2610-147']; tmpl_s=row['CC-2610-141']
last=max(row.values())
def add(vals, tr):
    global last
    last+=1; r=last
    for c in range(1,41):
        ws.cell(r,c)._style=copy.copy(ws.cell(tr,c)._style)
    for k,v in vals.items(): ws.cell(r,col(k)).value=v
    ws.cell(r,col('AF')).value=f'=IF($A{r}="","",IF($AC{r}<>"",$AC{r}-$B{r},TODAY()-$B{r}))'
    ws.cell(r,col('AG')).value=f'=IF($N{r}="","",IF(COUNTIFS($N$5:$N$1000,$N{r},$B$5:$B$1000,">="&($B{r}-90))>=3,"PROMOTE",""))'
    return r
KAT={'M':'melon','P':'Pine','I':'Import','L':'Local','C':'Banana'}
for l in P['lines']:
    if l['status']=='Lanjut': continue
    add({'A':l['id'],'B':D10,'C':l['typ'],'D':'A','E':'Tangerang','F':'Command Mingguan GT Tangerang','G':l['rung'],'H':f"{l['cust']} · {KAT[l['k2']]}",
         'I':(l['lamp'].replace('Lampiran ','') if l['lamp'].startswith('Lampiran') else l['lamp']),'J':l['prod'],'K':l['why'],'L':l['detail'],'O':round(l['val']*1e6),
         'Q':f"{l['do']} · {l['prod']}",'R':l['sales'],'S':'Salesman GT Tangerang','T':D15,'U':'ASM GT Tangerang · Band A','V':'WCR W46','W':l['tgt'],
         'X':f"Order ≥ {l['tgt']} di minggu 46",'Y':D15,'AE':0,'AH':'WCR W46 Tangerang · Pricing & Run-rate W46','AI':'Approve','AJ':f"W46: baru · harga rek. {P['price'][l['id']]['txt']}",'AK':'W46','AL':'W46','AM':'Aktif W46 (verdict Sen 16 Nov)','AN':'—'}, tmpl)
for l in P.get('dropped',[]):
    add({'A':l['id'],'B':D10,'C':l['typ'],'D':'A','E':'Tangerang','F':'Command Mingguan GT Tangerang','G':l['rung'],'H':f"{l['cust']} · melon",'I':'Grosir — tanpa outlet','J':l['prod'],'K':l['why'],'L':l['detail'],
         'Q':f"{l['do']} · {l['prod']}",'R':l['sales'],'S':'Salesman GT Tangerang','W':l['tgt'],'AE':0,'AH':'WCR W46 draft (8 Nov)','AI':'Drop','AJ':l['drop'],'AK':'W46','AL':'—','AM':'Tidak terbit (drop)','AN':'—'}, tmpl)
for c in P['sup']:
    if c['id']=='CC-2610-116': continue
    add({'A':c['id'],'B':D10,'C':c['typ'],'D':'B','E':'Tangerang','F':'B (bukan sales)','G':'①a Supply / harga','H':f"GT TGR · {c['owner']}",'K':c['prob'],'Q':c['cmd'],'S':c['owner'],
         'T':D15,'U':f"{c['owner']} · Band B",'V':'WCR W46','W':c['test'],'X':c['test'],'Y':D15,'AE':0,'AH':'Pricing & Run-rate W46 · WCR W46','AI':'Approve','AJ':'W46: baru','AK':'W46','AL':'W46','AM':'Aktif W46 (verdict Sen 16 Nov)','AN':'—'}, tmpl_s)
ws.cell(2,1).value=ws.cell(2,1).value.replace('CC-2610-001 … CC-2610-147','CC-2610-001 … CC-2611-018')+' Updated 9 Nov (interim file 6): W45 verdicts graded (final, order 2–8 Nov) and W46 lines added (IDs restart monthly: CC-2611-nnn).'
# ---- Verdict History
vrow={}
for r in range(5,vh.max_row+1):
    if vh.cell(r,2).value=='W45': vrow[vh.cell(r,1).value]=r
for vid,v in V.items():
    r=vrow[vid]; vh.cell(r,8).value=v['verdict']; vh.cell(r,9).value=v['res']+' · '+v['note']; vh.cell(r,10).value=round(v['inc']) if v['verdict'] in('Met','Partial') else 0
for vid,rr in VS.items():
    r=vrow[vid]; vh.cell(r,8).value=rr[4]; vh.cell(r,9).value=rr[3]; vh.cell(r,10).value=0
if 'CC-2610-137' in vrow: vh.cell(vrow['CC-2610-137'],8).value='—'
tv=vrow['CC-2610-147']; lastv=vh.max_row
def addv(vals):
    global lastv
    lastv+=1
    for c in range(1,11): vh.cell(lastv,c)._style=copy.copy(vh.cell(tv,c)._style)
    for j,v in enumerate(vals,1): vh.cell(lastv,j).value=v
for l in P['lines']:
    addv([l['id'],'W46','Tangerang',l['cust'],'Lanjut' if l['status']=='Lanjut' else 'Baru',f"{l['do']} · {l['prod']}",l['tgt'],'Open',l['why'],None])
for a in P['A2']:
    add({'A':a['id'],'B':D10,'C':'G','D':'A','E':'Tangerang','F':'Command Mingguan GT Tangerang — A2','G':'④ Habiskan stok','H':f"GT TGR · {a['lab']}",'I':'—','J':a['lab'],'K':a['why'],
         'L':f"Stok jatah GT TGR {round(a['q'])} box (SAP 9 Nov)",'O':round(a['val']),'Q':f"Tawarkan {a['lab']} di {a['price']} → {a['to']}",'R':a['sales'],'S':a['sales'],'T':D15,'U':'ASM GT Tangerang · Band A','V':'WCR W46',
         'W':f"Habis s/d {a['due']}",'X':f"Stok jatah GT TGR SKU ini 0 s/d {a['due']}",'Y':D15,'AE':0,'AH':'Register W46 · A2_STOK_BERISIKO','AI':'Approve','AJ':f"W46: baru · harga rek. {a['price']}",'AK':'W46','AL':'W46','AM':'Aktif W46 (verdict Sen 16 Nov)','AN':'—'}, tmpl)
    addv([a['id'],'W46','Tangerang',a['lab'],'A2 habiskan stok',f"Tawarkan di {a['price']} → {a['to']}",f"Habis s/d {a['due']}",'Open',a['why'],None])
for l in P.get('dropped',[]):
    addv([l['id'],'W46','Tangerang',l['cust'],'Drop sebelum terbit',f"{l['do']} · {l['prod']}",l['tgt'],'—',l['drop'],None])
for c in P['sup']:
    addv([c['id'],'W46','Tangerang',c['owner'],'Supply / harga',c['cmd'],c['test'],'Open',c['prob'],None])
vh.cell(1,1).value='Verdict History — W42 to W46'
vh.cell(2,1).value='One row per command × week. W42 = order 12–18 Okt; W43 = 19–25 Okt; W44 = 26 Okt–1 Nov (final 2 Nov); W45 = 2–8 Nov (final 9 Nov); W46 = issued 10 Nov, open.'
wb.save(OUT+'CC_Operating_System_Workbook_v3.1_Register_W42-W46.xlsx'); print('rows',last,lastv)
```

## build_wa.py
```python
from common import *
lines=P['lines']; c=P['cnt45']; E=P['gtstk']
o=[]
o.append("*COMMAND MINGGUAN GT TANGERANG — Minggu 46 (9–15 Nov)*")
o.append("_Final Senin 9 Nov · dibagi Selasa 10 Nov · order W45 s/d 8 Nov, stok SAP 9 Nov_")
o.append("")
o.append("*Hasil W45*")
o.append(f"• {c.get('Met',0)} Met · {c.get('Partial',0)} Partial · {c.get('Failed',0)} Failed. Met: Fresh Fruit Pine (155 BE), Ranum Pine (440 BE), Duta HE (170 BE), Sastro kiwi (147 BE). Tambahan ±Rp {n1(P['inc45']/1e6)} jt/mgg.")
o.append("• HE GT ±319 rb, order tetap naik 468 → 607 BE (Golden habis). *Jangan diskon HE* — pakai *harga rek.* di tiap baris.")
o.append(f"• Jatah GT TGR = stok umur > 3 hari (semua grade) + SP2/SP3. ORI & stok segar 0–3 hari = MT / cabang lain. Melon GT TGR ±{n0(E['melon'])} box. Cavendish di luar command minggu ini.")
o.append("")
o.append(f"*Baris minggu ini* ({len(lines)}/10)")
for k,l in enumerate(lines,1):
    st='lanjut' if l['status']=='Lanjut' else 'baru'
    lp=f" · {l['lamp'].replace('Lampiran ','')}" if l['lamp'].startswith('Lampiran') else ''
    o.append(f"{k}) *{l['cust']}* · {jaga(l)} · ID {sid(l['id'])} ({st}){lp}")
    o.append(f"   ➤ {l['do']} — {l['prod']} · 🎯 {l['tgt']} · 💰 {P['price'][l['id']]['txt']}  👤 {SHORT[l['sales']]}")
o.append("")
o.append(f"*A2 Habiskan stok berisiko* (di luar 10 slot)")
for a in P['A2']:
    o.append(f"• {sid(a['id'])} {a['lab']} — {n0(a['q'])} box · 💰 {a['price']} · ➤ {a['to']} · 👤 {a['sales']} · s/d {a['due']}")
o.append("")
o.append("*Jangan janji* (supply)")
for k,f_,_,todo in P['B']: o.append(f"• {k} {f_}: {todo}")
o.append("")
o.append("Balas kode per baris: *1 O 500* (O order·T tutup·S stok kosong·X tidak cocok·H harga — tulis harga pesaing)")
o.append("Lampiran outlet: *L1-1 O 2* = Lampiran 1, outlet 1, order 2 box · A2: *A2-012 O 3* = SKU 012, order 3 box")
o.append("Batas: *Senin 16 Nov 09:00*")
open(OUT+'WA_Command_W46_GT_Tangerang.txt','w').write('\n'.join(o)+'\n')
print('\n'.join(o)[:1500])
```

## fillcheck.py
```python
import sys, zipfile
from lxml import etree
A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
order=['lnL','lnR','lnT','lnB','lnTlToBr','lnBlToTr','cell3D','noFill','solidFill','gradFill','blipFill','pattFill','grpFill','headers','extLst']
for fn in sys.argv[1:]:
    z=zipfile.ZipFile(fn); bad=0
    for n in z.namelist():
        if n.startswith('ppt/slides/slide') and n.endswith('.xml'):
            t=etree.fromstring(z.read(n))
            for tc in t.iter(A+'tcPr'):
                tags=[e.tag.split('}')[1] for e in tc]; idx=[order.index(x) for x in tags if x in order]
                if idx!=sorted(idx) or sum(x in order[7:13] for x in tags)>1: bad+=1
    print(fn.split('/')[-1],'bad tcPr:',bad)
```

## stripfonts.py
```python
"""Remove embedded fonts (ppt/fonts/*.fntdata, embeddedFontLst, rels, content type) from a pptx in place."""
import sys, zipfile, re, shutil, os
for fn in sys.argv[1:]:
    tmp=fn+'.tmp'; zin=zipfile.ZipFile(fn); zout=zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        n=it.filename; d=zin.read(n)
        if n.startswith('ppt/fonts/'): continue
        if n=='ppt/presentation.xml':
            d=re.sub(rb'<p:embeddedFontLst>.*?</p:embeddedFontLst>',b'',d,flags=re.S)
            d=d.replace(b' embedTrueTypeFonts="1"',b'').replace(b' saveSubsetFonts="1"',b'')
        if n=='ppt/_rels/presentation.xml.rels':
            d=re.sub(rb'<Relationship [^>]*Target="fonts/[^"]*"/>',b'',d)
        if n=='[Content_Types].xml':
            d=re.sub(rb'<Default Extension="fntdata"[^>]*/>',b'',d)
        zout.writestr(it,d)
    zout.close(); zin.close(); shutil.move(tmp,fn); print('stripped',fn)
```

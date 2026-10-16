import json, glob, os, datetime as dt
from lxml import etree
B = __import__('os').environ.get('PIPE_DATA', 'dados')+'/nvda/dl/'
F = json.load(open(B + 'facts.json'))['facts']
P = [f'{q}Q{y:02d}' for y in range(20, 27) for q in range(1, 5)][:26]   # 1Q20..2Q26 calendario
def d(s): return dt.date.fromisoformat(s)
def label(end):
    e = d(end); m, y = e.month, e.year
    if m in (1, 2): return f'4Q{(y - 1) % 100:02d}'
    if m in (4, 5): return f'1Q{y % 100:02d}'
    if m in (7, 8): return f'2Q{y % 100:02d}'
    if m in (10, 11): return f'3Q{y % 100:02d}'
def fq(lab):   # trimestre fiscal Apple: cal 4Q -> FQ1, 1Q -> FQ2, 2Q -> FQ3, 3Q -> FQ4
    return {'4': 1, '1': 2, '2': 3, '3': 4}[lab[0]]
def kind_of(start, end):
    days = (d(end) - d(start)).days
    if 80 <= days <= 100: return '3M'
    if 170 <= days <= 195: return '6M'
    if 260 <= days <= 285: return '9M'
    if 350 <= days <= 380: return '12M'
def series(tag, kind, ns='us-gaap'):
    """kind: '3M','6M','9M','12M' (duracao) ou 'I' (instante). {label: (val, filed, accn)} - o filing mais recente vence."""
    t = F.get(ns, {}).get(tag)
    out = {}
    if not t: return out
    for u, facts in t['units'].items():
        for f in facts:
            if kind == 'I':
                if 'start' in f: continue
            else:
                if 'start' not in f or kind_of(f['start'], f['end']) != kind: continue
            lab = label(f['end'])
            if lab not in P and not (lab == '4Q19'): continue
            prev = out.get(lab)
            if prev is None or f['filed'] > prev[1]: out[lab] = (f['val'], f['filed'], f['accn'])
    return out
def first(tags, kind):
    out = {}
    for t in tags:
        for k, v in series(t, kind).items(): out.setdefault(k, v)
    return out
# ---- dimensional (instancias)
X = '{http://www.xbrl.org/2003/instance}'; XDI = '{http://xbrl.org/2006/xbrldi}'
FILED = {l.split()[0]: l.split()[2] for l in open(B + 'accns.txt') if l.strip()}
_dim = None
def dims():
    global _dim
    if _dim is not None: return _dim
    _dim = []
    for f in sorted(glob.glob(B + 'inst/*.xml')):
        a = os.path.basename(f)[:-4]
        t = etree.parse(f).getroot(); ctx = {}
        for c in t.iter(X + 'context'):
            p = c.find(X + 'period'); s = p.find(X + 'startDate'); e = p.find(X + 'endDate')
            if s is None: continue
            dd = tuple(sorted((m.get('dimension'), m.text) for m in c.iter(XDI + 'explicitMember')))
            ctx[c.get('id')] = (s.text, e.text, dd)
        for el in t:
            if not isinstance(el.tag, str): continue
            tag = etree.QName(el).localname
            if tag in ('Revenues', 'RevenueFromContractWithCustomerExcludingAssessedTax', 'CostOfRevenue', 'OperatingIncomeLoss') and el.get('contextRef') in ctx and el.text:
                s, e, dd = ctx[el.get('contextRef')]
                _dim.append((tag, s, e, dd, float(el.text), FILED.get(a, '0'), a))
    return _dim
def dseries(tag, member, kind, orig=False):
    out = {}
    for tg, s, e, dd, v, fl, a in dims():
        if tg != tag or kind_of(s, e) != kind: continue
        mem = [m for ax, m in dd if not (ax == 'srt:ConsolidationItemsAxis' and m == 'us-gaap:OperatingSegmentsMember')]
        if len(mem) != 1 or mem[0] != member: continue
        lab = label(e)
        if lab not in P: continue
        if lab not in out or (fl < out[lab][1] if orig else fl > out[lab][1]): out[lab] = (v, fl, a)
    return out
def inst_series(tag):
    """instante sem dimensão lido direto das instâncias XBRL (para tags ausentes do companyfacts). {label: (val, filed, accn)}"""
    out = {}
    for f in sorted(glob.glob(B + 'inst/*.xml')):
        a = os.path.basename(f)[:-4]
        t = etree.parse(f).getroot(); ctx = {}
        for c in t.iter(X + 'context'):
            p = c.find(X + 'period'); i = p.find(X + 'instant')
            if i is None or c.find('.//' + XDI + 'explicitMember') is not None: continue
            ctx[c.get('id')] = i.text
        for el in t:
            if isinstance(el.tag, str) and etree.QName(el).localname == tag and el.get('contextRef') in ctx and el.text:
                lab = label(ctx[el.get('contextRef')])
                if lab in P and (lab not in out or FILED.get(a, '0') > out[lab][1]): out[lab] = (float(el.text), FILED.get(a, '0'), a)
    return out
def qser(tag, member, orig=False):
    """trimestre-calendário: 3M do 10-Q; 4T (nov-jan, FQ4) = ano fiscal (10-K) - 9M (10-Q de out). orig=True usa o valor do filing original (não o reapresentado)."""
    s3, s9, s12 = dseries(tag, member, '3M', orig), dseries(tag, member, '9M', orig), dseries(tag, member, '12M', orig)
    out = {}
    for p in P:
        if p[0] == '4':
            pv = P[P.index(p) - 1]
            if p in s12 and pv in s9: out[p] = s12[p][0] - s9[pv][0]
        elif p in s3: out[p] = s3[p][0]
    return out
# ---- balanço de face: cada trimestre lido da instância do próprio 10-Q/10-K (sem misturar refilings)
RD = {l.split()[0]: l.split()[4] for l in open(B + 'accns.txt') if len(l.split()) == 5}
_face = None
# 4T25: o 10-K juntou títulos de dívida e ações listadas numa linha só; usa a coluna comparativa (jan/26) do 10-Q de abr/26, que separa as duas
OVR = {'4Q25': ('0001045810-26-000052', '2026-01-25')}
def face():
    """{label: [(tag, valor), ...]} na ordem da instância, até LiabilitiesAndStockholdersEquity."""
    global _face
    if _face is not None: return _face
    _face = {}
    items = list(RD.items()) + [(a, rd) for a, rd in OVR.values()]
    for a, rd in items:
        if label(rd) in OVR and (a, rd) != OVR[label(rd)]: continue
        fn = B + 'inst/' + a + '.xml'
        if not os.path.exists(fn): continue
        t = etree.parse(fn).getroot(); ctx = set()
        for c in t.iter(X + 'context'):
            i = c.find(X + 'period').find(X + 'instant')
            if i is not None and i.text == rd and c.find('.//' + XDI + 'explicitMember') is None: ctx.add(c.get('id'))
        out, seen = [], set()
        for el in t:
            if isinstance(el.tag, str) and el.get('contextRef') in ctx and el.text:
                n = etree.QName(el).localname
                if n in seen: continue
                seen.add(n)
                try: out.append((n, float(el.text)))
                except ValueError: continue
                if n == 'LiabilitiesAndStockholdersEquity': break
        lab = label(rd)
        if lab in P: _face[lab] = (out, a)
    return _face

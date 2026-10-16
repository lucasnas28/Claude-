import csv, datetime as dt, collections
B = __import__('os').environ.get('PIPE_DATA', 'dados')+'/aure/'
ROWS = [r for r in csv.DictReader(open(B + 'aure_cvm.csv')) if r['CD_CVM'] == '026620' and r['ORDEM_EXERC'] == 'ÚLTIMO']
def lab(d):
    d = dt.date.fromisoformat(d); return f'{(d.month - 1) // 3 + 1}Q{d.year % 100:02d}'
def st(n):
    for k in ('BPA', 'BPP', 'DRE', 'DFC_MI', 'DFC_MD', 'DMPL', 'DVA'):
        if k in n: return k
# keep only latest version per (statement, DT_REFER)
best = {}
for r in ROWS:
    k = (st(r['ARQ']), r['DT_REFER'])
    best[k] = max(best.get(k, 0), int(r['VERSAO']))
ROWS = [r for r in ROWS if int(r['VERSAO']) == best[(st(r['ARQ']), r['DT_REFER'])]]
V = collections.defaultdict(dict)   # (stmt, label, kind) -> {code: (desc, val)}
for r in ROWS:
    s = st(r['ARQ']); fim = r['DT_FIM_EXERC']
    if r['DT_REFER'] != fim: continue
    if s in ('BPA', 'BPP'): kind = 'I'
    else:
        days = (dt.date.fromisoformat(fim) - dt.date.fromisoformat(r['DT_INI_EXERC'])).days
        kind = '3M' if days < 100 else ('YTD')
    V[(s, lab(fim), kind)][r['CD_CONTA']] = (r['DS_CONTA'], float(r['VL_CONTA']))
def q(s, code, p):
    """valor trimestral isolado; DRE/DFC: 3M se houver, senão YTD - YTD anterior."""
    if s in ('BPA', 'BPP'): return V[(s, p, 'I')].get(code, (None, None))[1]
    if (s, p, '3M') in V and code in V[(s, p, '3M')]: return V[(s, p, '3M')][code][1]
    y = V.get((s, p, 'YTD'), {}).get(code)
    if y is None: return None
    if p[0] == '1': return y[1]
    pv = f'{int(p[0]) - 1}Q{p[2:]}'
    y0 = V.get((s, pv, 'YTD'), {}).get(code) or V.get((s, pv, '3M'), {}).get(code)
    return None if y0 is None else y[1] - y0[1]
import re as _re
PP = [f'{q_}Q{y:02d}' for y in range(21, 27) for q_ in range(1, 5)][:22]
def dfc_ytd(p, pat, codepref):
    """soma das linhas da DFC (MI ou MD) cuja descrição casa com pat, acumulado no ano até p."""
    for s in ('DFC_MI', 'DFC_MD'):
        for kind in ('YTD', '3M'):
            d = V.get((s, p, kind))
            if d:
                hits = [v for c, (ds, v) in d.items() if c.startswith(codepref) and _re.search(pat, ds, _re.I)]
                return sum(hits)
    return None
def dfc_q(p, pat, codepref):
    y = dfc_ytd(p, pat, codepref)
    if y is None: return None
    if p[0] == '1': return y
    y0 = dfc_ytd(f'{int(p[0]) - 1}Q{p[2:]}', pat, codepref)
    return None if y0 is None else y - y0
DA_PAT, DA_PRE = r'deprecia|amortiza..o de mais valia|amortiza..o de .gio', '6.01.01'
DIV_PAT, DIV_PRE = r'pagamento de dividendos', '6.03'
def bs(s, code, p):
    """saldo da conta reconstruído: se a conta tem filhas no período, usa a soma das filhas (cobre ITR com conta-mãe ausente)."""
    d = V.get((s, p, 'I'), {})
    kids = [c for c in d if c.startswith(code + '.') and c.count('.') == code.count('.') + 1]
    ks = sum(bs(s, c, p) for c in kids) if kids else None
    if code in d:
        own = d[code][1]
        return ks if ks is not None and abs(ks) > abs(own) + 1 else own
    return ks or 0.0

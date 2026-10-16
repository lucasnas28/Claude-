import sys, re, copy, json
sys.path.insert(0, __import__('os').environ.get('PIPE_DATA', 'dados')+'/work')
from segdata import seg
from sec import PERIODS as P
import openpyxl
from openpyxl.utils import get_column_letter as CL
from openpyxl.formula.translate import Translator
from openpyxl.styles import Font, Alignment

SRC, OUT = sys.argv[1], sys.argv[2]
LAYF = __import__('os').environ.get('PIPE_DATA', 'dados')+'/work/layout.json'
lay = json.load(open(LAYF))
wb = openpyxl.load_workbook(SRC)
inp, w2, m = wb['Inputs'], wb['Inputs (2)'], wb['Model']
PC = {p: 11 + i for i, p in enumerate(P)}
col = lambda p: CL(PC[p])
OLD_END, R0, HDR = lay['END'], lay['R0'], 4
pos = lay['pos']
bold = Font(bold=True)

BU = [  # key, member, code, PT label, EN label
 ('dc',   'DataCenter', 45, 'Data Center', 'Data Center'),
 ('comp', 'Compute', 46, 'Data Center - Compute (aberto desde o FY24)', 'Compute'),
 ('net',  'Networking', 47, 'Data Center - Networking (aberto desde o FY24)', 'Networking'),
 ('gam',  'Gaming', 48, 'Gaming', 'Gaming'),
 ('pv',   'ProfessionalVisualization', 49, 'Professional Visualization', 'Professional Visualization'),
 ('auto', 'Automotive', 50, 'Automotive', 'Automotive'),
 ('oem',  'OEMAndOther', 51, 'OEM & Other', 'OEM and Other'),
 ('hyp',  'Hyperscale', 52, 'Data Center - Hyperscale (abertura FY27; reapresentada desde 1T FY26)', 'Hyperscale'),
 ('aic',  'AICloudsIndustrialEnterprise', 53, 'Data Center - AI Clouds, Industrial & Enterprise (abertura FY27)', 'AI Clouds, Industrial and Enterprise'),
 ('edge', 'EdgeComputing', 54, 'Edge Computing (abertura FY27 = Gaming + ProViz + Auto + OEM & Other)', 'Edge Computing'),
]
data = {k: {p: v[0] for p, v in seg(mem).items()} for k, mem, *_ in BU}

r = lay['LASTROW'] + 2
def put_row(rr, cods, unid, pt, en, origem):
    for i, (c, s) in enumerate(cods): inp.cell(rr, 1 + 2 * i, c); inp.cell(rr, 2 + 2 * i, s)
    inp.cell(rr, 7, unid); inp.cell(rr, 8, pt); inp.cell(rr, 9, en); inp.cell(rr, 10, origem)
    for p in P: inp.cell(rr, PC[p]).number_format = '#,##0;(#,##0)'

inp.cell(r, 8, 'RECEITA POR MERCADO - BRUTO COMO DIVULGADO (USD unidades; nota de segmentos dos 10-Q/10-K, eixo srt:ProductOrServiceAxis). 1T-3T = 3M; coluna 4T = ANO FISCAL 12M').font = bold
r += 1
raw = {}
for k, mem, code, pt, en in BU:
    put_row(r, [], 'moeda', pt + ' [3M; 4T = 12M]', f'{en} revenue [us-gaap:Revenues, nvda:{mem}Member]', 'memoria')
    for p, x in data[k].items(): inp.cell(r, PC[p], x)
    raw[k] = r; r += 1
r += 1
inp.cell(r, 8, 'RECEITA POR MERCADO TRIMESTRALIZADA (derivada; 4T = ano fiscal - 9M; so onde 1T-3T existem). Codigos 45-54 = memoria, nao entram na receita do cod 1').font = bold
r += 1
der = {}
for k, mem, code, pt, en in BU:
    put_row(r, [(code, 1)], 'moeda', pt + ' (trimestre)', en + ', quarterly', 'derivada')
    rr = raw[k]
    for p in P:
        if p not in data[k]: continue
        q = int(p[0]); yy = p[2:]
        if q < 4: inp.cell(r, PC[p], f'={col(p)}{rr}')
        else:
            g = [f'{i}Q{yy}' for i in (1, 2, 3)]
            if all(x in data[k] for x in g):
                inp.cell(r, PC[p], f'={col(p)}{rr}-SUM({col(g[0])}{rr}:{col(g[2])}{rr})')
    der[k] = r; r += 1
def chk(label, terms, minus, only):
    global r
    put_row(r, [], 'moeda', label, '', 'derivada')
    for p in P:
        if p in only:
            inp.cell(r, PC[p], '=' + '+'.join(f'{col(p)}{der[t]}' for t in terms) + f'-{col(p)}{minus(p)}')
    r += 1
qrev = pos['q_rev']
oldp = [p for p in P if all(p in data[k] for k in ('dc', 'gam', 'pv', 'auto', 'oem'))]
cnp = [p for p in P if p in data['comp'] and inp.cell(der['comp'], PC[p]).value is not None]
newp = [p for p in P if p in data['hyp']]
chk('CHECK: DC + Gaming + ProViz + Auto + OEM - receita total (deve ser 0)', ['dc', 'gam', 'pv', 'auto', 'oem'], lambda p: qrev, oldp)
chk('CHECK: Compute + Networking - Data Center (deve ser 0)', ['comp', 'net'], lambda p: der['dc'], cnp)
chk('CHECK: Hyperscale + AI Clouds/Ind./Ent. - Data Center (deve ser 0)', ['hyp', 'aic'], lambda p: der['dc'], newp)
chk('CHECK: Data Center + Edge Computing - receita total (deve ser 0)', ['dc', 'edge'], lambda p: qrev, newp)
inp.cell(r, 8, '[vazio] Compute/Networking antes do 1T FY24: NVIDIA so abriu o Data Center a partir do FY24 (o FY23 aparece apenas como total anual na coluna 4Q22 do bruto)')
r += 1
inp.cell(r, 8, '[vazio] Gaming, ProViz, Automotive, OEM & Other a partir do 1T FY27: NVIDIA passou a divulgar so Data Center (Hyperscale / AI Clouds) e Edge Computing; 3T-4T FY26 nao foram reapresentados no formato novo')
LAST = r
NEW_END = ((LAST + 30 + 9) // 10) * 10

# move the conventions block and widen the Inputs (2) formula range
oldconv = OLD_END + 3
for rr in range(oldconv, oldconv + 8):
    for c in range(1, 12): inp.cell(rr, c).value = None
conv = NEW_END + 3
inp.cell(conv, 8, 'CONVENCOES').font = bold
notes = [
 f'Linhas de origem: {R0} a {LAST}. As formulas da Inputs (2) varrem as linhas {R0} a {NEW_END}; as linhas {LAST + 1} a {NEW_END} estao livres para novas contas.',
 f'Periodos: colunas {col(P[0])} a {col(P[-1])}, rotulo de trimestre-calendario. Cada trimestre fiscal NVIDIA vai para o trimestre-calendario em que majoritariamente cai (linha 2 mostra o trimestre fiscal; linha 3 a data de encerramento). Ex.: 1Q25 = 1T FY26 (fev-abr/2025).',
 'Brutos em USD unidades e acoes em unidades; a Inputs (2) multiplica pelo fator da celula B1 as linhas com UNID = moeda ou acoes. LPA e fatores passam sem fator.',
 'SINAL converte para a convencao do modelo (custos, despesas, impostos, D&A, juros e dividendos negativos). COD2/SINAL2 usam a mesma linha num segundo destino (reclassificacao/estorno).',
 'ORIGEM: fonte = copiado do XBRL; derivada = formula nesta aba (trimestralizacao, ajuste de desdobramento, checks); memoria = informativa, sem codigo.',
 'Valores do XBRL: para cada periodo usa-se o fato arquivado mais recentemente (ja reflete reapresentacoes e desdobramentos nas comparacoes posteriores).',
 'Codigos 45-54 (receita por mercado) sao memoria: alimentam o bloco de mercados no fim do Model, nao a receita total (cod 1).',
]
for i, t in enumerate(notes): inp.cell(conv + 1 + i, 8, t)

rx = re.compile(r'(Inputs!\$[A-Z]{1,2}\$\d+:\$[A-Z]{1,2}\$)' + str(OLD_END) + r'\b')
nre = 0
for row in w2.iter_rows():
    for c in row:
        if isinstance(c.value, str) and c.value.startswith('=') and 'Inputs!' in c.value:
            nv = rx.sub(lambda g: g.group(1) + str(NEW_END), c.value)
            if nv != c.value: c.value = nv; nre += 1

# Inputs (2) rows for codes 45-54
w2['C52'] = 'Receita por mercado (memo, como divulgado) - codigos 45-54 [Model linhas 363-377]'
w2['C52'].font = bold
annual = {8: ('D', 'G'), 13: ('I', 'L'), 18: ('N', 'Q'), 23: ('S', 'V'), 28: ('X', 'AA'), 33: ('AC', 'AF'), 38: ('AH', 'AK')}
row2 = {}
for i, (k, mem, code, pt, en) in enumerate(BU):
    rr = 53 + i; row2[k] = rr
    w2.cell(rr, 2, code); w2.cell(rr, 3, en + ' revenue')
    for c in range(1, 39): w2.cell(rr, c)._style = copy.copy(w2.cell(4, c)._style)
    for c in range(4, 38):
        if c in annual: continue
        w2.cell(rr, c, Translator(w2.cell(48, c).value, origin=f'{CL(c)}48').translate_formula(f'{CL(c)}{rr}'))
    for c, (a, b) in annual.items():
        w2.cell(rr, c, f'=SUM({a}{rr}:{b}{rr})')

# Model block at the bottom
B0 = 361
mann = {8: ('D', 'G'), 13: ('I', 'L'), 18: ('N', 'Q'), 23: ('S', 'V'), 28: ('X', 'AA'), 33: ('AC', 'AF')}
for rr in range(B0, B0 + 20):
    for c in range(1, 35): assert m.cell(rr, c).value is None, (rr, c)
m.cell(B0, 3, 'Receita por mercado - como divulgado pela NVIDIA (USD milhões)')._style = copy.copy(m['C13']._style)
for c in range(4, 34):
    m.cell(B0, c, f'={CL(c)}7')._style = copy.copy(m.cell(13, c)._style)
lines = [  # Model row, key or None, label, kind
 (B0 + 2, 'dc', 'Data Center', 'bu'),
 (B0 + 3, 'comp', '   Compute', 'bu'),
 (B0 + 4, 'net', '   Networking', 'bu'),
 (B0 + 5, 'gam', 'Gaming', 'bu'),
 (B0 + 6, 'pv', 'Professional Visualization', 'bu'),
 (B0 + 7, 'auto', 'Automotive', 'bu'),
 (B0 + 8, 'oem', 'OEM & Other', 'bu'),
 (B0 + 9, None, 'Net Revenues (soma dos mercados)', 'sum'),
 (B0 + 10, None, 'Check vs Net Revenues (linha 15)', 'chk'),
 (B0 + 12, None, 'Abertura nova (FY27 em diante; reapresentada so no 1Q25 e 2Q25 desta grade)', 'hdr'),
 (B0 + 13, 'hyp', '   Data Center - Hyperscale', 'bu'),
 (B0 + 14, 'aic', '   Data Center - AI Clouds, Industrial & Enterprise', 'bu'),
 (B0 + 15, 'edge', 'Edge Computing', 'bu'),
 (B0 + 17, None, 'Mix: Data Center % da receita', 'mix'),
]
SUMR = B0 + 9
nm = 0
for rr, k, lab, kind in lines:
    lc = m.cell(rr, 3, lab)
    lc._style = copy.copy(m['C15']._style)
    if kind in ('sum', 'hdr'): lc.font = Font(bold=True, name=m['C15'].font.name, size=m['C15'].font.sz)
    if kind == 'hdr': continue
    for c in range(4, 34):
        L = CL(c); cell = m.cell(rr, c)
        cell._style = copy.copy(m.cell(15, c)._style)
        if kind == 'bu':
            if c in mann:
                a, b = mann[c]; cell.value = f'=IF(COUNT({a}{rr}:{b}{rr})=4,SUM({a}{rr}:{b}{rr}),"")'
            else:
                cell.value = f"=IFERROR(INDEX('Inputs (2)'!$D${row2[k]}:$AL${row2[k]},MATCH({L}$7,'Inputs (2)'!$D$2:$AL$2,0)),\"\")"
        elif kind == 'sum':
            cell.value = f'=SUM({L}{B0 + 2},{L}{B0 + 5}:{L}{B0 + 8})'
        elif kind == 'chk':
            cell.value = f'={L}15-{L}{SUMR}'
        elif kind == 'mix':
            cell.value = f'=IFERROR({L}{B0 + 2}/{L}15,"")'
            cell.number_format = '0.0%'
        nm += 1
# row 17 "Check Receita" next to Net Revenues
for c in range(4, 34):
    assert m.cell(17, c).value is None
    m.cell(17, c, f'={CL(c)}15-{CL(c)}{SUMR}')._style = copy.copy(m.cell(15, c)._style)
    nm += 1
m['C17'] = f'Check Receita (linha 15 - soma dos mercados, linha {SUMR})'

lay['END'] = NEW_END; lay['LASTROW'] = LAST
lay['bu_raw'] = raw; lay['bu_der'] = der
json.dump(lay, open(LAYF, 'w'))
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print('Inputs new rows to', LAST, 'END', OLD_END, '->', NEW_END, 'Inputs(2) refs widened', nre, 'Model cells', nm)

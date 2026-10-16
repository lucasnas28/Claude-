import sys, re, copy, datetime as dt, json
import openpyxl
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI
from openpyxl.formula.translate import Translator
from openpyxl.styles import Font

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)
m, v, w2, inp = wb['Model'], wb['Valuation'], wb['Inputs (2)'], wb['Inputs']
lay = json.load(open(__import__('os').environ.get('PIPE_DATA', 'dados')+'/work/layout.json'))
log = []

def tr(f, src, dst): return Translator(f, origin=src).translate_formula(dst)

# ---------------------------------------------------------------- 1. missing historical formulas
gaps = {'I': [134], 'AF': [16, 21, 25, 37, 41, 98, 105, 112, 116, 122, 134, 143, 156, 165, 169, 172, 181, 202, 207]}
for col, rws in gaps.items():
    ref = 'N' if col == 'I' else 'AE'
    for r in rws:
        dst = f'{col}{r}'
        assert m[dst].value is None, dst
        m[dst] = tr(m[f'{ref}{r}'].value, f'{ref}{r}', dst)
        log.append(f'Model!{dst} = {m[dst].value}')

# ---------------------------------------------------------------- 2. Valuation column alignment
SHIFT = 10   # two annual blocks of 5 columns: Valuation D (2023) must read Model W, not M
pat = re.compile(r"Model!(\$?)([A-Z]{1,3})(\$?)(\d+)")
def shift(f):
    return pat.sub(lambda g: f"Model!{g.group(1)}{CL(CI(g.group(2)) + SHIFT)}{g.group(3)}{g.group(4)}", f)
nshift = 0
for row in v.iter_rows(min_row=16, max_row=56, min_col=4, max_col=18):
    for c in row:
        if isinstance(c.value, str) and 'Model!' in c.value:
            c.value = shift(c.value); nshift += 1
v['G24'] = v['G24'].value.replace('Model!AL99', 'Model!AL98')        # template typo (row 99 is empty)
# columns S..V (2038-2041) lie beyond the Model horizon (ends 2037): clear their series
ncl = 0
for row in v.iter_rows(min_row=18, max_row=56, min_col=19, max_col=22):
    for c in row:
        if c.value is not None: c.value = None; ncl += 1
grey = Font(color='FF808080')
for a in ('S16', 'T16', 'U16', 'V16'): v[a].font = grey
v['S15'] = 'Fora do horizonte do Model (projecao vai ate 2037)'; v['S15'].font = grey
# terminal year moves from V to R (2037)
v['D64'] = '=R50'; v['D67'] = '=D66*R55'; v['D68'] = '=D67*R54'
v['G57'] = '=SUM(G56:R56)'
v['D59'] = dt.datetime(2037, 12, 31)
v['Y18'] = '=((R18/L18)^(1/6))-1'; v['Y22'] = '=((R22/L22)^(1/6))-1'
v['X16'] = '=F16&"-"&L16'; v['Y16'] = '=L16&"-"&R16'
v['X49'] = '=F16'; v['Y49'] = '=L16'; v['Z49'] = '=R16'; v['Z50'] = '=R50'

# ---------------------------------------------------------------- 3. IS check against reported net income
pos = lay['pos']
qn = pos['q_ni']
assert inp.cell(qn, 1).value is None
inp.cell(qn, 1, 44); inp.cell(qn, 2, 1)
inp.cell(qn, 8, 'Lucro liquido (trimestre) - COD 44 alimenta o check IS do Model (linha 4)')
w2['B50'] = 44; w2['C50'] = 'Net income reported (SEC) - memo p/ check IS [Model linha 4]'
annual = {8, 13, 18, 23, 28, 33, 38}
for c in range(1, 39):
    w2.cell(50, c)._style = copy.copy(w2.cell(13, c)._style)
for c in range(4, 38):
    if c in annual: continue
    w2.cell(50, c, tr(w2.cell(48, c).value, f'{CL(c)}48', f'{CL(c)}50'))
for c, (a, b) in {8: ('D', 'G'), 13: ('I', 'L'), 18: ('N', 'Q'), 23: ('S', 'V'), 28: ('X', 'AA'), 33: ('AC', 'AF'), 38: ('AH', 'AK')}.items():
    w2.cell(50, c, f'=SUM({a}50:{b}50)')
m['C4'] = 'Net Income reportado (SEC) - base do check IS'
mann = {8: ('D', 'G'), 13: ('I', 'L'), 18: ('N', 'Q'), 23: ('S', 'V'), 28: ('X', 'AA'), 33: ('AC', 'AF')}
for c in range(4, 34):
    cell = m.cell(4, c); assert cell.value is None
    if c in mann: cell.value = f'=SUM({mann[c][0]}4:{mann[c][1]}4)'
    else: cell.value = f"=INDEX('Inputs (2)'!$D$50:$AL$50,MATCH({CL(c)}$7,'Inputs (2)'!$D$2:$AL$2,0))"
    cell._style = copy.copy(m.cell(103, c)._style)
m['C4']._style = copy.copy(m['C103']._style)
n3 = 0
for c in range(4, m.max_column + 1):
    L = CL(c)
    if c >= 34 and m.cell(3, c).value is None: continue
    m.cell(3, c).value = f'=IF(COUNT({L}4)=0,"",ABS({L}103-{L}4))'; n3 += 1
    if c < 34 and c not in mann: m.cell(3, c)._style = copy.copy(m['H3']._style)

# ---------------------------------------------------------------- 4. US macro block (FRED)
macro = {
 8:  ('CPI EUA (dez/dez) - FRED CPIAUCNS', [0.0136, 0.0704, 0.0645, 0.0335, 0.0289, 0.0268]),
 9:  ('PIB real EUA (var. anual) - FRED GDPC1', [-0.0208, 0.0627, 0.0243, 0.0292, 0.0295, 0.0234]),
 10: ('Fed Funds efetiva (media anual) - FRED FEDFUNDS', [0.0038, 0.0008, 0.0168, 0.0502, 0.0514, 0.0421]),
}
for r, (lab, vals) in macro.items():
    m.cell(r, 3).value = lab
    for c, x in zip(('H', 'M', 'R', 'W', 'AB', 'AG'), vals):
        m[f'{c}{r}'] = x
        m[f'{c}{r}'].number_format = '0.00%'
    for c in range(34, m.max_column + 1):
        assert m.cell(r, c).value is None

# ---------------------------------------------------------------- 5. Inputs (2) header typo
assert w2['AG2'].value == 2025
w2['AG2'] = 2026

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print('\n'.join(log)); print('valuation refs shifted', nshift, 'cleared S..V', ncl, 'row3 checks', n3)

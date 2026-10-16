"""Fixes for the standard model mask (Modelo_2.0_proprio_Padrao_AJUSTADO-2028).
usage: fix_mask.py IN OUT [--only-new]   (--only-new skips the fixes already applied to filled models)"""
import sys, re, copy, datetime as dt
import openpyxl
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI
from openpyxl.formula.translate import Translator
from openpyxl.styles import Font

IN, OUT = sys.argv[1], sys.argv[2]
ONLY_NEW = '--only-new' in sys.argv
wb = openpyxl.load_workbook(IN)
m, val, w2 = wb['Model'], wb['Valuation'], wb['Inputs (2)']
log = []
ANNUAL = set(range(8, 108, 5))            # H, M, R, ... CY
HISTQ = [c for c in range(4, 33) if c not in ANNUAL]     # D..AF quarters
PROJQ = [c for c in range(34, 108) if c not in ANNUAL]   # AH..CX quarters
def prevq(c):
    p = c - 1
    return p - 1 if p in ANNUAL else p
def setf(coord, f, why):
    old = m[coord].value
    if old != f:
        m[coord] = f; log.append((coord, why))
def tr(f, s, d): return Translator(f, origin=s).translate_formula(d)

if not ONLY_NEW:
    # 1. missing historical formulas (capex in 1Q21/4Q25, ratios in 4Q25)
    for colx, rws in {'I': [134], 'AF': [16, 21, 25, 37, 41, 98, 105, 112, 116, 122, 134, 143, 156, 165, 169, 172, 181, 202, 207]}.items():
        ref = 'N' if colx == 'I' else 'AE'
        for r in rws:
            if m[f'{colx}{r}'].value is None:
                setf(f'{colx}{r}', tr(m[f'{ref}{r}'].value, f'{ref}{r}', f'{colx}{r}'), 'formula historica faltando')
    # 2. Valuation columns aligned with Model years; horizon ends 2037 (end of the Model)
    pat = re.compile(r"Model!(\$?)([A-Z]{1,3})(\$?)(\d+)")
    for row in val.iter_rows(min_row=16, max_row=56, min_col=4, max_col=18):
        for c in row:
            if isinstance(c.value, str) and 'Model!' in c.value:
                c.value = pat.sub(lambda g: f"Model!{g.group(1)}{CL(CI(g.group(2)) + 10)}{g.group(3)}{g.group(4)}", c.value)
    val['G24'] = val['G24'].value.replace('Model!AL99', 'Model!AL98')
    for row in val.iter_rows(min_row=18, max_row=56, min_col=19, max_col=22):
        for c in row: c.value = None
    for a in ('S16', 'T16', 'U16', 'V16'): val[a].font = Font(color='FF808080')
    val['S15'] = 'Fora do horizonte do Model (projecao vai ate 2037)'; val['S15'].font = Font(color='FF808080')
    val['D64'] = '=R50'; val['D67'] = '=D66*R55'; val['D68'] = '=D67*R54'; val['G57'] = '=SUM(G56:R56)'
    val['D59'] = dt.datetime(2037, 12, 31)
    val['Y18'] = '=((R18/L18)^(1/6))-1'; val['Y22'] = '=((R22/L22)^(1/6))-1'
    val['X16'] = '=F16&"-"&L16'; val['Y16'] = '=L16&"-"&R16'
    val['X49'] = '=F16'; val['Y49'] = '=L16'; val['Z49'] = '=R16'; val['Z50'] = '=R50'
    log.append(('Valuation', 'colunas alinhadas aos anos do Model; horizonte ate 2037; G24 linha 99->98'))
    # 3. IS check: row 4 = reported net income (input or link); row 3 compares and stays quiet when row 4 is empty
    m['C4'] = 'Net Income reportado (preencher ou ligar) - base do check IS'
    m['C4']._style = copy.copy(m['C103']._style)
    for c in range(4, 34):
        m.cell(4, c)._style = copy.copy(m.cell(103, c)._style)
        if c in ANNUAL:
            a, b = CL(c - 4), CL(c - 1); m.cell(4, c).value = f'=IF(COUNT({a}4:{b}4)=4,SUM({a}4:{b}4),"")'
    for c in range(4, m.max_column + 1):
        L = CL(c)
        if c >= 34 and m.cell(3, c).value is None: continue
        m.cell(3, c).value = f'=IF(COUNT({L}4)=0,"",ABS({L}103-{L}4))'
        if c < 34 and c not in ANNUAL: m.cell(3, c)._style = copy.copy(m['H3']._style)
    log.append(('Model!3:4', 'check IS compara com o lucro reportado da linha 4'))
    # 5. Inputs (2) header typo
    if w2['AG2'].value == 2025: w2['AG2'] = 2026; log.append(('Inputs (2)!AG2', '2025 -> 2026'))

# ---------------------------------------------------------------- new fixes
# 6. leverage ratios
for c in HISTQ + PROJQ:
    L = CL(c)
    if m.cell(300, c).value is not None:
        setf(f'{L}300', f'=SUM({L}178,{L}179,{L}180,{L}183,{L}201,{L}204)/({L}35*4)', 'divida/EBITDA dividia pela linha 16 (% YoY)')
    if m.cell(302, c).value is not None:
        setf(f'{L}302', f'=(SUM({L}178,{L}179,{L}180,{L}183,{L}201,{L}204)-{L}110)/({L}35*4)', 'divida liquida com linhas erradas (184, 205; sem 178 e 201)')
    if m.cell(303, c).value is not None:
        setf(f'{L}303', f'=(SUM({L}178,{L}179,{L}183,{L}204)-{L}110)/({L}35*4)', 'divida liquida ex IFRS16 com linhas erradas (175, 176, 181, 202, 106, 28)')
# 7. total assets #REF! in the first projected quarter of each year
for c in PROJQ:
    L = CL(c)
    if '#REF!' in str(m.cell(163, c).value):
        setf(f'{L}163', tr(m.cell(163, c + 1).value, f'{CL(c + 1)}163', f'{L}163'), 'total do ativo com #REF! (intangivel fora da soma)')
# 8. intangibles roll-forward mirror: historical row 150 pointed to goodwill (145)
for c in HISTQ:
    L = CL(c)
    if m.cell(150, c).value is not None or c == 32:
        setf(f'{L}150', f'={L}146', 'linha 150 (intangivel) espelhava o goodwill (145)')
# 9. roll-forwards must start from the previous quarter, not from an empty annual column or two quarters back
first = PROJQ[0]   # AH
for c in PROJQ:
    L, P = CL(c), CL(prevq(c))
    setf(f'{L}137', f'={P}137+SUM({L}134:{L}136)', 'imobilizado: partia de coluna anual vazia / pulava trimestre')
    setf(f'{L}150', f'={P}150+SUM({L}148:{L}149)', 'intangivel: partia de coluna anual vazia / pulava trimestre')
    seed189 = f'{P}183' if c == first else f'{P}189'
    setf(f'{L}189', f'={seed189}+SUM({L}187:{L}188)', 'divida LP: partia de coluna anual vazia / AI189 era 0 digitado')
    if c != first:     # AH keeps AG225 / AG230 as the seed input (current price, current shares)
        setf(f'{L}225', f'={P}225*(1+{L}226)', 'preco medio: pulava trimestre / partia de coluna anual vazia')
        setf(f'{L}230', f'={P}230+{L}229-{L}228', 'acoes no fim do periodo: pulava trimestre / partia de coluna anual vazia')
        setf(f'{L}231', f'=AVERAGE({P}230,{L}230)', 'media de acoes: pulava trimestre / partia de coluna anual vazia')
# 10. free cash flow before debt: "(X258,X265)" is a union of two cells (gives #VALUE!) and points one row above CFO / CFI
for c in PROJQ:
    L = CL(c)
    if str(m.cell(196, c).value).startswith('=('):
        setf(f'{L}196', f'=SUM({L}259,{L}266)+SUM({L}268:{L}274)', 'FCF antes da divida: uniao (X258,X265) dava #VALUE!; agora CFO (259) + CFI (266) + fluxos de capital')
# seed inputs flagged in the sheet
for coord, note in (('AG197', 'INPUT: caixa minimo como % da receita (usado em AH197:CX197)'),
                    ('AG225', 'INPUT: preco medio da acao atual (semente da linha 225)'),
                    ('AG230', 'INPUT: acoes em circulacao no fim do ultimo trimestre (semente da linha 230)')):
    if m[coord].value is None:
        m[coord].fill = openpyxl.styles.PatternFill('solid', fgColor='FFFFF2CC')
        log.append((coord, note))

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
from collections import Counter
print(len(log), 'alteracoes'); print(Counter(w for _, w in log).most_common())

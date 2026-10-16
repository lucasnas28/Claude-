"""Dupla checagem de caixa e receita de um modelo convertido contra o modelo de origem (máscara antiga).
uso: dupla_checagem.py ORIGEM.xlsx CONVERTIDO.xlsx   (os dois com cache de cálculo)
Por trimestre de 1T21 a 2T26 confere: receita líquida (Model e soma dos segmentos), caixa do balanço, linhas de check
(BS e DRE), lucro líquido e o plug de reconciliação de caixa da DFC (diferença entre a DFC indireta e a variação do caixa)."""
import sys, openpyxl
from openpyxl.utils import get_column_letter as L
a = openpyxl.load_workbook(sys.argv[1], data_only=True)['Model']
b = openpyxl.load_workbook(sys.argv[2], data_only=True)['Model']
def row(ws, label, start=8):
    # entre as linhas com o rótulo, a que tem mais números de 1T21 a 2T26 (ex.: o caixa do balanço, não o da receita financeira)
    rs = [r for r in range(start, ws.max_row + 1) if str(ws.cell(r, 3).value or '').strip().lower().startswith(label.lower())]
    return max(rs, key=lambda r: sum(isinstance(ws.cell(r, c).value, (int, float)) and ws.cell(r, c).value != 0 for c in range(9, 36)), default=None)
cols = [c for c in range(9, 36) if (c - 9) % 5 != 4]          # só trimestres (pula M, R, W, AB, AG)
def per(c):
    q = (c - 9) - (c - 9) // 5; return f'{q % 4 + 1}T{21 + q // 4}'
num = lambda x: x if isinstance(x, (int, float)) else None
linhas = {'receita': 'Net Revenues', 'caixa': 'Cash and equivalents', 'lucro': 'Net Income'}
ra = {k: row(a, v) for k, v in linhas.items()}; rb = {k: row(b, v) for k, v in linhas.items()}
seg = 17; plug = row(b, 'Others / cash reconciliation plug'); dcx = row(b, 'Total increase/decrease of cash')
prob = []
print('per    receita(conv)  dif.orig  seg.chk | caixa(conv)  dif.orig | BS chk  DRE chk | plug DFC')
for c in cols:
    v = {k: (num(a.cell(ra[k], c).value), num(b.cell(rb[k], c).value)) for k in linhas}
    d = {k: (None if None in v[k] else round(v[k][1] - v[k][0], 2)) for k in linhas}
    s = num(b.cell(seg, c).value) if seg else None
    bs, isc = num(b.cell(2, c).value), num(b.cell(3, c).value)
    pl = num(b.cell(plug, c).value) if plug else None
    print(f'{per(c):5}', f'{v["receita"][1] or 0:14,.1f}', f'{d["receita"] if d["receita"] is not None else "-":>9}', f'{s if s is not None else "-":>8} |',
          f'{v["caixa"][1] or 0:11,.1f}', f'{d["caixa"] if d["caixa"] is not None else "-":>9} |', f'{bs if bs is not None else "-":>6}', f'{isc if isc is not None else "-":>8} |',
          f'{round(pl, 1) if pl is not None else "-":>8}')
    for k in linhas:
        if d[k] is not None and abs(d[k]) > 0.5: prob.append(f'{per(c)} {k} difere da origem em {d[k]}')
        if v[k][0] is not None and v[k][1] is None: prob.append(f'{per(c)} {k} vazio no convertido')
    if s is not None and abs(s) > 0.5: prob.append(f'{per(c)} soma dos segmentos difere da receita em {s}')
    if bs is None or isc is None: prob.append(f'{per(c)} check {"BS" if bs is None else "DRE"} vazio')
    elif bs > 0.5 or isc > 0.5: prob.append(f'{per(c)} check BS {bs} DRE {isc}')
print('RESULTADO:', 'OK' if not prob else 'VER: ' + '; '.join(prob[:12]))

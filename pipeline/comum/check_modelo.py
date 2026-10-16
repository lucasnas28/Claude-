"""Confere um modelo na máscara nova: checks do Model (linhas 1 a 3) zerados de 1T21 a 2T26 e nenhum erro de fórmula.
uso: check_modelo.py MODELO.xlsx  (com cache de cálculo: rode depois do recalc + inject)"""
import sys, openpyxl
v = openpyxl.load_workbook(sys.argv[1], data_only=True); m = v['Model']
bad = [(r, c) for r in (1, 2, 3) for c in range(9, 36) if isinstance(m.cell(r, c).value, (int, float)) and abs(m.cell(r, c).value) > 0.5]
err = sum(1 for ws in v.worksheets for row in ws.iter_rows() for x in row if isinstance(x.value, str) and x.value[:2] in ('#D', '#V', '#R', '#N'))
print(sys.argv[1].split('/')[-1].ljust(34), 'checks OK' if not bad else f'CHECK {bad[:4]}', '| erros', err)
sys.exit(1 if bad or err else 0)

#!/bin/bash
# uso: post.sh PREFIX  (PREFIX_a.xlsx -> PREFIX_b.xlsx realizado + recalc + IFERROR nos #DIV/0! históricos)
H="$(cd "$(dirname "$0")" && pwd)"; W="$H/../comum"; X=$1
python3 $W/generic_post.py ${X}_a.xlsx ${X}_b.xlsx --realize >/dev/null && timeout 900 python3 $W/recalc.py ${X}_b.xlsx ${X}_b_recalc.xlsx 2>&1 | grep -v javaldx
python3 - "$X" <<'PY' 2>&1 | grep -v -i warn
import openpyxl, sys
X=sys.argv[1]
v=openpyxl.load_workbook(X+'_b_recalc.xlsx',data_only=True)['Model']
rows=sorted(set(r for r in range(1,400) for c in range(4,36) if isinstance(v.cell(r,c).value,str) and v.cell(r,c).value.startswith('#')))
wb=openpyxl.load_workbook(X+'_b.xlsx'); m=wb['Model']; n=0
for rw in rows:
    for c in range(4,108):
        f=m.cell(rw,c).value
        if isinstance(f,str) and f.startswith('=') and 'IFERROR' not in f: m.cell(rw,c).value=f'=IFERROR({f[1:]},"")'; n+=1
wb.calculation.fullCalcOnLoad=True; wb.save(X+'_b.xlsx'); print('rows with errors', rows, 'wrapped', n)
PY
timeout 900 python3 $W/recalc.py ${X}_b.xlsx ${X}_b_recalc.xlsx 2>&1 | grep -v javaldx

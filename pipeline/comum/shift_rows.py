"""Insert N blank rows in a sheet at row AT, updating every formula reference in the workbook,
conditional formatting, row dimensions and freeze panes.  usage: shift_rows.py IN OUT SHEET AT N"""
import sys, re, copy
import openpyxl
from openpyxl.formula import Tokenizer
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.worksheet.dimensions import RowDimension

IN, OUT, SHEET, AT, N = sys.argv[1], sys.argv[2], sys.argv[3], int(sys.argv[4]), int(sys.argv[5])
MAXR = 1048576
wb = openpyxl.load_workbook(IN)
target = wb[SHEET]
cellre = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")

def shift_ref(ref):
    def sub(mm):
        row = int(mm.group(4))
        if row >= AT: row = min(row + N, MAXR)
        return f'{mm.group(1)}{mm.group(2)}{mm.group(3)}{row}'
    return cellre.sub(sub, ref)

def shift_formula(f, own_sheet):
    tok = Tokenizer(f)
    changed = False
    for t in tok.items:
        if t.type == 'OPERAND' and t.subtype == 'RANGE':
            v = t.value
            if '!' in v:
                sh, ref = v.rsplit('!', 1)
                if sh.strip("'") != SHEET: continue
                nv = sh + '!' + shift_ref(ref)
            else:
                if own_sheet != SHEET or not cellre.search(v): continue
                nv = shift_ref(v)
            if nv != v: t.value = nv; changed = True
    return tok.render() if changed else f

n_f = 0
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('='):
                nf = shift_formula(c.value, ws.title)
                if nf != c.value: c.value = nf; n_f += 1

# conditional formatting of the target sheet
old_cf = target.conditional_formatting
new_cf = ConditionalFormattingList()
for cf in old_cf:
    sq = ' '.join(shift_ref(str(rg)) for rg in str(cf.sqref).split())
    for rule in cf.rules:
        rule = copy.copy(rule)
        if rule.formula: rule.formula = [shift_formula('=' + x, SHEET)[1:] for x in rule.formula]
        new_cf.add(sq, rule)
target.conditional_formatting = new_cf

# row dimensions
old_rd = {k: copy.copy(v) for k, v in target.row_dimensions.items()}
target.row_dimensions.clear()
target.insert_rows(AT, N)
for k, v in old_rd.items():
    nk = k + N if k >= AT else k
    v.index = nk
    target.row_dimensions[nk] = v
for r in range(AT, AT + N):
    target.row_dimensions[r] = RowDimension(target, index=r)

pn = target.sheet_view.pane
if pn is not None and pn.ySplit and pn.ySplit + 1 >= AT:      # só desloca se a linha congelada estiver abaixo da inserção
    target.freeze_panes = shift_ref(target.freeze_panes)        # (freeze_panes devolve o topLeftCell, que é a posição de rolagem, não o split)
for ws in wb.worksheets:
    for n, d in list(ws.defined_names.items()):
        if SHEET in str(d.attr_text): d.attr_text = shift_formula('=' + d.attr_text, ws.title)[1:]
for n, d in list(wb.defined_names.items()):
    if SHEET in str(d.attr_text): d.attr_text = shift_formula('=' + d.attr_text, None)[1:]
wb.save(OUT)
print('formulas updated', n_f, '| rows inserted', N, 'at', AT)

"""Corrige o congelamento de painéis da aba Model direto no XML, sem tocar no resto do arquivo.

Dois sintomas, os dois impedem a planilha de rolar:
- ySplit inflado pelo shift_rows (linha ~350)
- ySplit já em 11, mas topLeftCell ainda em D336 (herdado da máscara)

uso: fix_panes.py ARQ.xlsx [...]"""
import sys, re, zipfile, shutil, os
NEW = ('<sheetViews><sheetView showGridLines="0" tabSelected="1" zoomScale="85" zoomScaleNormal="85" workbookViewId="0">'
       '<pane xSplit="3" ySplit="11" topLeftCell="D12" activePane="bottomRight" state="frozen"/>'
       '<selection pane="topRight" activeCell="D1" sqref="D1"/><selection pane="bottomLeft" activeCell="A12" sqref="A12"/>'
       '<selection pane="bottomRight" activeCell="D12" sqref="D12"/></sheetView></sheetViews>')
def model_path(z):
    wb = z.read('xl/workbook.xml').decode()
    rid = re.search(r'<sheet [^>]*name="Model"[^>]*r:id="(rId\d+)"', wb).group(1)
    rels = z.read('xl/_rels/workbook.xml.rels').decode()
    tgt = re.search(r'<Relationship [^>]*Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or re.search(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    t = tgt.group(1).lstrip('/')
    return t if t.startswith('xl/') else 'xl/' + t
for f in sys.argv[1:]:
    with zipfile.ZipFile(f) as z:
        mp = model_path(z); x = z.read(mp).decode('utf-8')
        pane = re.search(r'<pane [^>]*/?>', x)
        y = re.search(r'ySplit="(\d+)"', pane.group(0) if pane else '')
        tl = re.search(r'topLeftCell="([^"]+)"', pane.group(0) if pane else '')
        if pane and y and int(y.group(1)) == 11 and tl and tl.group(1) == 'D12':
            print(f, 'ok', y.group(1), tl.group(1)); continue
        x2 = re.sub(r'<sheetViews>.*?</sheetViews>', NEW, x, count=1, flags=re.S)
        tmp = f + '.tmp'
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as out:
            for it in z.infolist():
                out.writestr(it, x2.encode('utf-8') if it.filename == mp else z.read(it.filename))
    shutil.move(tmp, f)
    print(f, 'corrigido', (y.group(1) if y else '?'), (tl.group(1) if tl else '?'), '-> 11 D12')

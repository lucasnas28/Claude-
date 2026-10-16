import sys, zipfile, shutil, openpyxl, re
from lxml import etree
out, calc, final = sys.argv[1:4]
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
vals = openpyxl.load_workbook(calc, data_only=True)
wbo = openpyxl.load_workbook(out)           # to map sheet names -> part paths
zin = zipfile.ZipFile(out)
rels = etree.fromstring(zin.read('xl/_rels/workbook.xml.rels'))
wbx = etree.fromstring(zin.read('xl/workbook.xml'))
rid2t = {r.get('Id'): r.get('Target') for r in rels}
sheetpart = {}
for s in wbx.find(f'{{{NS}}}sheets'):
    rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    t = rid2t[rid].lstrip('/'); t = t if t.startswith('xl/') else 'xl/' + t
    sheetpart[t] = s.get('name')
ERR = {'#DIV/0!', '#N/A', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#NULL!'}
stats = {}
zout = zipfile.ZipFile(final, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename in sheetpart:
        name = sheetpart[item.filename]; ws = vals[name]
        root = etree.fromstring(data)
        n = miss = 0
        for c in root.iter(f'{{{NS}}}c'):
            f = c.find(f'{{{NS}}}f')
            if f is None: continue
            v = c.find(f'{{{NS}}}v')
            if v is None: v = etree.SubElement(c, f'{{{NS}}}v')
            x = ws[c.get('r')].value
            if x is None:
                c.set('t', 'str'); v.text = ''; miss += 1
                continue
            if isinstance(x, bool): c.set('t', 'b'); v.text = '1' if x else '0'
            elif isinstance(x, (int, float)): c.attrib.pop('t', None); v.text = repr(float(x)) if isinstance(x, float) else str(x)
            elif isinstance(x, str) and x in ERR: c.set('t', 'e'); v.text = x
            elif hasattr(x, 'hour') and not hasattr(x, 'year'):
                c.attrib.pop('t', None); v.text = repr((x.hour*3600+x.minute*60+x.second)/86400)
            elif hasattr(x, 'year'):
                from openpyxl.utils.datetime import to_excel
                c.attrib.pop('t', None); v.text = repr(float(to_excel(x)))
            else: c.set('t', 'str'); v.text = str(x)
            n += 1
        stats[name] = (n, miss)
        data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    zout.writestr(item, data)
zout.close()
print(stats)

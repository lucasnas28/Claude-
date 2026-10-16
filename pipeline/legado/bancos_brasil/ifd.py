"""Leitura do IF.data (BCB, API Olinda) ja baixado em ifd/raw. Valores em R$ -> devolve R$ milhoes."""
import json, glob, re, os
D = __import__('os').environ.get('PIPE_DATA', 'dados')+'/bank/ifd/raw/'
P = [f'{q}Q{y:02d}' for y in range(20, 27) for q in range(1, 5)][:26]
def am(p): return f'20{p[2:]}{int(p[0]) * 3:02d}'
_cache = {}
def rows(amv, tipo, rel, cod):
    k = (amv, tipo, rel, cod)
    if k not in _cache:
        f = f'{D}{amv}_{tipo}_{rel}.json'
        out = {}
        if os.path.exists(f):
            for x in json.load(open(f))['value']:
                if x['CodInst'] == cod:
                    nm = re.sub(r'\s+', ' ', x['NomeColuna']).strip()
                    out[((x['Grupo'] or '').strip(), nm)] = x['Saldo']
        _cache[k] = out
    return _cache[k]
def new(p): return am(p) >= '202503'
def letter(p, rel, cod, let, tipo=1):
    """valor (R$ mi) da linha cujo nome termina em '(let)' ou contem '(let) ='"""
    for (g, nm), v in rows(am(p), tipo, rel, cod).items():
        m = re.search(r'\(([a-z]{1,2}\d?)\)', nm)
        if m and m.group(1) == let:
            return None if v is None else v / 1e6
    return None
def name(p, rel, cod, pat, grupo=None, tipo=1):
    for (g, nm), v in rows(am(p), tipo, rel, cod).items():
        if re.search(pat, nm) and (grupo is None or re.search(grupo, g)):
            return None if v is None else v / 1e6
    return None
def lsum(p, rel, cod, lets, tipo=1):
    vs = [letter(p, rel, cod, l, tipo) for l in lets]
    if all(v is None for v in vs): return None
    return sum(v for v in vs if v is not None)

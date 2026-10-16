import json, os, sys, urllib.request, urllib.parse, concurrent.futures as cf
D = __import__('os').environ.get('PIPE_DATA', 'dados')+'/bank/ifd/raw/'
PRU = ['C0080336', 'C0080312', 'C0080178', 'C0083694', 'C0088022', 'C0080934']   # BTG, ABC, BMG, Agibank, PicPay, BR Partners
FIN = ['C0049944', 'C0041856', 'C0030290', 'C0051956', 'C0052632']
AMS = [f'{y}{m:02d}' for y in range(2019, 2027) for m in (3, 6, 9, 12)]
AMS = [a for a in AMS if '201912' <= a <= '202606']
def get(am, tipo, rel, codes):
    f = f'{D}{am}_{tipo}_{rel}.json'
    if os.path.exists(f) and os.path.getsize(f) > 20: return f
    flt = ' or '.join(f"CodInst eq '{c}'" for c in codes)
    q = f"@AnoMes={am}&@TipoInstituicao={tipo}&@Relatorio='{rel}'&$filter=" + urllib.parse.quote(flt, safe="'") + "&$format=json"
    url = 'https://olinda.bcb.gov.br/olinda/servico/IFDATA/versao/v1/odata/IfDataValores(AnoMes=@AnoMes,TipoInstituicao=@TipoInstituicao,Relatorio=@Relatorio)?' + q
    for k in range(4):
        try:
            data = urllib.request.urlopen(url, timeout=240).read()
            json.loads(data); open(f, 'wb').write(data); return f
        except Exception as e:
            err = e
    return f'ERR {am} {tipo} {rel} {err}'
jobs = [(am, 1, rel, PRU) for am in AMS for rel in (1, 2, 3, 4, 5, 6)] + [(am, t, rel, PRU if t == 1 else FIN) for am in AMS for t in (1, 2) for rel in (11, 13, 14, 16)]
with cf.ThreadPoolExecutor(6) as ex:
    for r in ex.map(lambda j: get(*j), jobs):
        if r.startswith('ERR'): print(r)
print('done', len(jobs))

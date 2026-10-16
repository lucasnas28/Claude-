import os, sys
sys.path.insert(0, os.path.join(os.environ.get('PIPE_SEC', 'pipeline/sec'), '..', 'empresas', 'shel'))
from book import load
from cfg import BOOK, DL
D = load(BOOK)
TITLE = ('DRIVERS DE RECEITA (codigos 45+): receita de terceiros por segmento (Integrated Gas, Upstream, '
         'Marketing, Chemicals & Products, Renewables and Energy Solutions) e Adjusted Earnings por segmento — '
         'Quarterly Databook Q2 2026 do RI (USD milhões, fator 1).')
NOTE = ('Notas: receita de terceiros as reported (aba de cada segmento, linha Third-party abaixo de Revenue:). '
        'A soma dos cinco segmentos difere da receita do grupo em até US$ 17 mi (1T22); no 2T26 a diferença é de US$ −2 mi. '
        'Não interpolar o residual — a Shell não abre essa linha na face de segmentos (Corporate não tem receita de terceiros própria no databook). '
        'Intersegmento fica de fora (é eliminado na consolidação). Adjusted Earnings é a abertura de resultado que a Shell usa no release, não EBIT IFRS. '
        'Série: 1T21-2T26. 1T20-4T20 vazios: o databook vigente começa em Q1 2021.')
C17 = 'Check Receita (linha 15 − soma dos segmentos de terceiros)'
SEGS = [('ig', 'Integrated Gas'), ('up', 'Upstream'), ('mk', 'Marketing'),
        ('cp', 'Chemicals & Products'), ('res', 'Renewables and Energy Solutions')]
def items(co):
    out = []
    for k, nm in SEGS:
        out.append(('s_' + k, f'{nm} - receita de terceiros (USD mi)', 'moeda', D['tp'][k], '#,##0', 'sum'))
    for k, nm in SEGS + [('co', 'Corporate')]:
        out.append(('a_' + k, f'{nm} - Adjusted Earnings (USD mi)', 'moeda', D['ae'][k], '#,##0', 'sum'))
    return out
BLOCK = r'''
hdr_row(rr, 'Receita de terceiros por segmento (USD milhões, as reported)'); rr += 1
for k, nm in SEGS: line('s_' + k, '   ' + nm)
line('tot', 'Receita de terceiros (soma dos segmentos)', annual='calc', fn=sumf(*['s_' + k for k, nm in SEGS]), b=True)
line('chk', C17, annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["tot"]})=0,"",ROUND({CL(c)}15-{CL(c)}{R["tot"]},1))')
for k, nm in SEGS[:3]: line('y_' + k, '   % YoY ' + nm, '0.0%', 'calc', yoy(R['s_' + k]))
rr += 1
hdr_row(rr, 'Adjusted Earnings por segmento (USD milhões; métrica da Shell, pós-itens identificados e CCS)'); rr += 1
for k, nm in SEGS + [('co', 'Corporate')]: line('a_' + k, '   ' + nm)
line('ae_t', 'Adjusted Earnings (soma dos segmentos)', annual='calc', fn=sumf(*['a_' + k for k, nm in SEGS + [('co', 'Corporate')]]), b=True)
'''
MONTHMAP = {}

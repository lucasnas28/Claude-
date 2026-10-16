import sys; sys.path.insert(0, __import__('os').environ.get('PIPE_SEC', 'pipeline/sec')); from segspec import make; from rpo import inst
globals().update(make(__import__('os').environ.get('PIPE_DATA', 'dados')+'/crwv/dl',
 'DRIVERS DE RECEITA (codigos 45+): receita EUA x fora dos EUA e RPO (carteira contratada) - SEC XBRL (USD, fator 0,000001).',
 'Notas: CoreWeave tem um único segmento. IPO em mar/2025: os trimestres de 2024 vêm das colunas comparativas dos 10-Q de 2025 e do 10-K de 2025 (4T24 = ano - 9M); o balanço começa no 4T24 (comparativo do 10-Q do 1T25). Não há dado trimestral publicado antes de 2024.',
 [('us', 'Estados Unidos', ['country:US'], False), ('xus', 'Fora dos EUA', ['us-gaap:NonUsMember'], False)],
 RV=('RevenueFromContractWithCustomerExcludingAssessedTax', 'Revenues')))
_items = items
def items(co):
    return _items(co) + [('rpo', 'RPO - obrigações de desempenho remanescentes, saldo no fim do período (USD)', 'moeda', inst(co, ['RevenueRemainingPerformanceObligation']), '#,##0', 'last')]
BLOCK += '''
rr += 1
hdr_row(rr, 'Carteira contratada (USD milhões)'); rr += 1
line('rpo', '   RPO - carteira contratada (saldo)', annual='last')
line('rpo_y', '   % YoY RPO', '0.0%', 'calc', yoy(R['rpo']))
'''

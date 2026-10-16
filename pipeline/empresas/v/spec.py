DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/v/dl'
TITLE = 'DRIVERS DE RECEITA (codigos 45+): linhas de receita da Visa (serviço, processamento de dados, transações internacionais, outras, incentivos a clientes) e EUA x exterior - SEC XBRL (USD, fator 0,000001).'
NOTE = 'Notas: valores dimensionais das instâncias XBRL; ano fiscal termina em setembro. Incentivos a clientes entram negativos. No balanço, a linha "PL sem tag própria" de 2020-22 são as ações preferenciais séries A/B (etiquetadas por série).'
C17 = 'Check Receita (linha 15 - soma das linhas)'
RC, RV = 'RevenueFromContractWithCustomerExcludingAssessedTax', 'Revenues'
def items(co):
    def q(mem):
        out = co.dq(RC, [mem])
        for k, v in co.dq(RV, [mem]).items(): out.setdefault(k, v)
        return out
    ci = {k: -abs(v) for k, v in q('v:ClientIncentivesMember').items()}
    return [
     ('r_svc', 'Receita de serviços (USD)', 'moeda', q('us-gaap:ServiceMember'), '#,##0', 'sum'),
     ('r_dp', 'Receita de processamento de dados (USD)', 'moeda', q('v:DataProcessingRevenuesMember'), '#,##0', 'sum'),
     ('r_it', 'Receita de transações internacionais (USD)', 'moeda', q('v:InternationalTransactionRevenuesMember'), '#,##0', 'sum'),
     ('r_ot', 'Outras receitas (USD)', 'moeda', q('us-gaap:ServiceOtherMember'), '#,##0', 'sum'),
     ('r_ci', 'Incentivos a clientes (USD, negativo)', 'moeda', ci, '#,##0', 'sum'),
     ('g_us', 'Receita EUA (USD)', 'moeda', q('country:US'), '#,##0', 'sum'),
     ('g_xu', 'Receita fora dos EUA (USD)', 'moeda', q('us-gaap:NonUsMember'), '#,##0', 'sum'),
    ]
BLOCK = r"""
hdr_row(rr, 'Receita - como a Visa divulga (USD milhões)'); rr += 1
line('r_svc', '   Serviços'); line('r_dp', '   Processamento de dados'); line('r_it', '   Transações internacionais'); line('r_ot', '   Outras'); line('r_ci', '   (-) Incentivos a clientes')
line('tot', 'Receita líquida (soma)', annual='calc', fn=sumf('r_svc', 'r_dp', 'r_it', 'r_ot', 'r_ci'), b=True)
line('chk', 'Check vs Net Revenues (linha 15)', annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["tot"]})=0,"",ROUND({CL(c)}15-{CL(c)}{R["tot"]},1))')
line('p_ci', '   Incentivos / receita bruta', '0.0%', 'calc', lambda c: f'=IFERROR(-{CL(c)}{R["r_ci"]}/({CL(c)}{R["tot"]}-{CL(c)}{R["r_ci"]}),"")')
line('y_it', '   % YoY transações internacionais', '0.0%', 'calc', yoy(R['r_it']))
rr += 1
hdr_row(rr, 'Região (USD milhões)'); rr += 1
line('g_us', '   EUA'); line('g_xu', '   Fora dos EUA')
line('p_xu', '   % fora dos EUA', '0.0%', 'calc', ratio(R['g_xu'], R['tot']))
"""

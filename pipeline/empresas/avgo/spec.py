DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/avgo/dl'
TITLE = 'DRIVERS DE RECEITA (codigos 45+): segmentos Semicondutores x Software de Infraestrutura (receita e lucro operacional), produtos x assinaturas e regiões - SEC XBRL (USD, fator 0,000001).'
NOTE = 'Notas: valores dimensionais das instâncias XBRL; ano fiscal termina no início de novembro (fechamento de fev -> 4T, mai -> 1T, ago -> 2T, nov -> 3T). VMware consolidada a partir de nov/23.'
C17 = 'Check Receita (linha 15 - semicondutores + software)'
RC, OI = 'RevenueFromContractWithCustomerExcludingAssessedTax', 'OperatingIncomeLoss'
SEMI, SOFT = 'avgo:SemiconductorSolutionsMember', 'avgo:InfrastructureSoftwareMember'
def items(co):
    q = co.dq
    def reg(m):
        a, b = q(RC, [(m, 'us-gaap:ProductMember')]), q(RC, [(m, 'avgo:SubscriptionsandServicesMember')])
        return {p: a[p] + b[p] for p in a if p in b}
    return [
     ('s_semi', 'Semiconductor Solutions - receita (USD)', 'moeda', q(RC, [SEMI]), '#,##0', 'sum'),
     ('s_soft', 'Infrastructure Software - receita (USD)', 'moeda', q(RC, [SOFT]), '#,##0', 'sum'),
     ('o_semi', 'Semiconductor Solutions - lucro operacional (USD)', 'moeda', q(OI, [SEMI]), '#,##0', 'sum'),
     ('o_soft', 'Infrastructure Software - lucro operacional (USD)', 'moeda', q(OI, [SOFT]), '#,##0', 'sum'),
     ('t_prd', 'Receita de produtos (USD)', 'moeda', q(RC, ['us-gaap:ProductMember']), '#,##0', 'sum'),
     ('t_sub', 'Receita de assinaturas e serviços (USD)', 'moeda', q(RC, ['avgo:SubscriptionsandServicesMember']), '#,##0', 'sum'),
     ('g_am', 'Receita Américas (USD)', 'moeda', reg('srt:AmericasMember'), '#,##0', 'sum'),
     ('g_ap', 'Receita Ásia-Pacífico (USD)', 'moeda', reg('srt:AsiaPacificMember'), '#,##0', 'sum'),
     ('g_em', 'Receita EMEA (USD)', 'moeda', reg('us-gaap:EMEAMember'), '#,##0', 'sum'),
    ]
BLOCK = r"""
hdr_row(rr, 'Segmentos (USD milhões)'); rr += 1
line('s_semi', '   Semiconductor Solutions'); line('s_soft', '   Infrastructure Software')
line('tot', 'Receita (soma)', annual='calc', fn=sumf('s_semi', 's_soft'), b=True)
line('chk', 'Check vs Net Revenues (linha 15)', annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["tot"]})=0,"",ROUND({CL(c)}15-{CL(c)}{R["tot"]},1))')
line('y_semi', '   % YoY semicondutores', '0.0%', 'calc', yoy(R['s_semi']))
line('y_soft', '   % YoY software', '0.0%', 'calc', yoy(R['s_soft']))
line('o_semi', '   Lucro operacional semicondutores'); line('o_soft', '   Lucro operacional software')
line('m_semi', '   Margem operacional semicondutores', '0.0%', 'calc', ratio(R['o_semi'], R['s_semi']))
line('m_soft', '   Margem operacional software', '0.0%', 'calc', ratio(R['o_soft'], R['s_soft']))
rr += 1
hdr_row(rr, 'Tipo e região (USD milhões)'); rr += 1
line('t_prd', '   Produtos'); line('t_sub', '   Assinaturas e serviços')
line('g_am', '   Américas'); line('g_ap', '   Ásia-Pacífico'); line('g_em', '   EMEA')
"""

DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/ge/dl'
TITLE = 'DRIVERS DE RECEITA (codigos 45+): produtos x serviços e segmentos da GE Aerospace (CES e DPT: receita e lucro) - SEC XBRL (USD, fator 0,000001).'
NOTE = 'Notas: valores dimensionais das instâncias XBRL. Segmentos CES e DPT só a partir de 2023 (reapresentados desde 4T22). A DRE do modelo segue a última apresentação de cada período: 2020-21 é a GE inteira, 2022 já sem a GE HealthCare (cisão jan/23) e 2023 em diante sem a GE Vernova (cisão abr/24); o balanço é o de cada data (até mar/24 inclui a Vernova).'
C17 = 'Receita de seguros e quebras de perímetro (linha 15 - produtos - serviços)'
RV, OI = 'Revenues', 'OperatingIncomeLoss'
CES, DPT = 'ge:CommercialEnginesAndServicesReportableSegmentMember', 'ge:DefenseAndPropulsionTechnologiesReportableSegmentMember'
def items(co):
    q = co.dq
    return [
     ('t_prd', 'Receita de produtos (USD)', 'moeda', q(RV, ['us-gaap:ProductMember']), '#,##0', 'sum'),
     ('t_srv', 'Receita de serviços (USD)', 'moeda', q(RV, ['us-gaap:ServiceMember']), '#,##0', 'sum'),
     ('s_ces', 'Commercial Engines & Services - receita (USD)', 'moeda', q(RV, [CES]), '#,##0', 'sum'),
     ('s_dpt', 'Defense & Propulsion Technologies - receita (USD)', 'moeda', q(RV, [DPT]), '#,##0', 'sum'),
     ('o_ces', 'Commercial Engines & Services - lucro do segmento (USD)', 'moeda', q(OI, [CES]), '#,##0', 'sum'),
     ('o_dpt', 'Defense & Propulsion Technologies - lucro do segmento (USD)', 'moeda', q(OI, [DPT]), '#,##0', 'sum'),
     ('c_srv', 'CES - receita de serviços (USD)', 'moeda', q(RV, [(CES, 'us-gaap:ServiceMember')]), '#,##0', 'sum'),
    ]
BLOCK = r"""
hdr_row(rr, 'Receita por tipo (USD milhões)'); rr += 1
line('t_prd', '   Produtos'); line('t_srv', '   Serviços')
line('tot', 'Receita (soma)', annual='calc', fn=sumf('t_prd', 't_srv'), b=True)
line('chk', 'Diferença vs linha 15: receita de seguros (run-off) e quebras de perímetro (cisões HealthCare 2023 e Vernova 2024)', annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["tot"]})=0,"",ROUND({CL(c)}15-{CL(c)}{R["tot"]},1))')
line('p_srv', '   % serviços na receita', '0.0%', 'calc', ratio(R['t_srv'], R['tot']))
rr += 1
hdr_row(rr, 'Segmentos GE Aerospace (USD milhões; desde 4T22)'); rr += 1
line('s_ces', '   Receita CES'); line('c_srv', '      dos quais serviços'); line('s_dpt', '   Receita DPT')
line('y_ces', '   % YoY CES', '0.0%', 'calc', yoy(R['s_ces']))
line('o_ces', '   Lucro CES'); line('o_dpt', '   Lucro DPT')
line('m_ces', '   Margem CES', '0.0%', 'calc', ratio(R['o_ces'], R['s_ces']))
line('m_dpt', '   Margem DPT', '0.0%', 'calc', ratio(R['o_dpt'], R['s_dpt']))
"""

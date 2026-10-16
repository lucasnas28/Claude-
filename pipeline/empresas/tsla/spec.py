DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/tsla/dl'
TITLE = 'DRIVERS DE RECEITA (codigos 45+): receita e lucro bruto por segmento (Automotivo x Energia), linhas de receita automotiva e regiões - SEC XBRL (USD, fator 0,000001).'
NOTE = 'Notas: valores dimensionais das instâncias XBRL. A abertura da receita automotiva mudou em 2021-22 (vendas com e sem garantia de recompra -> vendas; créditos regulatórios separados); os membros foram costurados. O 10-Q do 1T23 trocou as etiquetas de Automotivo e Energia (1T23 e coluna comparativa do 1T22); os valores foram destrocados.'
C17 = 'Check Receita (linha 15 - soma dos segmentos)'
RV, RC, GP = 'Revenues', 'RevenueFromContractWithCustomerExcludingAssessedTax', 'GrossProfit'
AUT, ENE = 'tsla:AutomotiveSegmentMember', 'tsla:EnergyGenerationAndStorageSegmentMember'
def items(co):
    q = co.dq
    seg_a = q(RV, [AUT]); seg_a.update({k: v for k, v in q(RC, [AUT]).items() if k not in seg_a})
    seg_e = q(RV, [ENE]); seg_e.update({k: v for k, v in q(RC, [ENE]).items() if k not in seg_e})
    for p in list(seg_a):                     # o 10-Q do 1T23 trocou os dois segmentos (no 1T23 e na coluna comparativa do 1T22)
        if p in seg_e and seg_e[p] > seg_a[p]: seg_a[p], seg_e[p] = seg_e[p], seg_a[p]
    sales = q(RC, ['tsla:AutomotiveSalesMember', ('tsla:AutomotiveMember', 'tsla:AutomotiveSalesMember')])
    w, wo = q(RC, [('tsla:AutomotiveMember', 'tsla:AutomotiveSalesWithResaleValueGuaranteeMember')]), q(RC, [('tsla:AutomotiveMember', 'tsla:AutomotiveSalesWithoutResaleValueGuaranteeMember')])
    for p in wo:
        if p not in sales: sales[p] = wo[p] + w.get(p, 0)
    return [
     ('s_aut', 'Segmento Automotivo - receita (USD)', 'moeda', seg_a, '#,##0', 'sum'),
     ('s_ene', 'Segmento Geração e Armazenamento de Energia - receita (USD)', 'moeda', seg_e, '#,##0', 'sum'),
     ('gp_aut', 'Segmento Automotivo - lucro bruto (USD)', 'moeda', q(GP, [AUT]), '#,##0', 'sum'),
     ('gp_ene', 'Segmento Energia - lucro bruto (USD)', 'moeda', q(GP, [ENE]), '#,##0', 'sum'),
     ('a_sal', 'Vendas de veículos (USD; até 2021 com + sem garantia de recompra)', 'moeda', sales, '#,##0', 'sum'),
     ('a_crd', 'Créditos regulatórios (USD)', 'moeda', q(RC, ['tsla:AutomotiveRegulatoryCreditsMember', ('tsla:AutomotiveMember', 'tsla:AutomotiveRegulatoryCreditsMember')]), '#,##0', 'sum'),
     ('a_lea', 'Leasing automotivo (USD)', 'moeda', q(RC, ['tsla:AutomotiveLeasingMember']), '#,##0', 'sum'),
     ('a_srv', 'Serviços e outros (USD)', 'moeda', q(RC, ['tsla:ServicesAndOtherMember']), '#,##0', 'sum'),
     ('g_us', 'Receita Estados Unidos (USD)', 'moeda', q(RV, ['country:US']), '#,##0', 'sum'),
     ('g_cn', 'Receita China (USD)', 'moeda', q(RV, ['country:CN']), '#,##0', 'sum'),
    ]
BLOCK = r"""
hdr_row(rr, 'Segmentos - como a Tesla divulga (USD milhões)'); rr += 1
line('s_aut', 'Automotivo (inclui serviços e outros)'); line('s_ene', 'Geração e armazenamento de energia')
line('tot', 'Receita (soma dos segmentos)', annual='calc', fn=sumf('s_aut', 's_ene'), b=True)
line('chk', 'Check vs Net Revenues (linha 15)', annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["tot"]})=0,"",ROUND({CL(c)}15-{CL(c)}{R["tot"]},1))')
line('y_aut', '   % YoY Automotivo', '0.0%', 'calc', yoy(R['s_aut']))
line('y_ene', '   % YoY Energia', '0.0%', 'calc', yoy(R['s_ene']))
line('gp_aut', '   Lucro bruto Automotivo'); line('gp_ene', '   Lucro bruto Energia')
line('m_aut', '   Margem bruta Automotivo', '0.0%', 'calc', ratio(R['gp_aut'], R['s_aut']))
line('m_ene', '   Margem bruta Energia', '0.0%', 'calc', ratio(R['gp_ene'], R['s_ene']))
rr += 1
hdr_row(rr, 'Receita automotiva por linha (USD milhões)'); rr += 1
line('a_sal', '   Vendas de veículos'); line('a_crd', '   Créditos regulatórios'); line('a_lea', '   Leasing'); line('a_srv', '   Serviços e outros')
line('p_crd', '   % créditos regulatórios na receita', '0.0%', 'calc', ratio(R['a_crd'], R['tot']))
rr += 1
hdr_row(rr, 'Receita por região (USD milhões)'); rr += 1
line('g_us', '   Estados Unidos'); line('g_cn', '   China')
line('g_ot', '   Outros países (receita total - EUA - China)', annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["g_us"]})=0,"",{CL(c)}15-N({CL(c)}{R["g_us"]})-N({CL(c)}{R["g_cn"]}))')
line('p_cn', '   % China na receita', '0.0%', 'calc', ratio(R['g_cn'], R['tot']))
"""

DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/amd/dl'
TITLE = 'DRIVERS DE RECEITA (codigos 45+): segmentos Data Center, Client, Gaming e Embedded (receita e lucro operacional) e segmentos antigos (até 1T22) - SEC XBRL (USD, fator 0,000001).'
NOTE = 'Notas: valores dimensionais das instâncias XBRL. Segmentos atuais a partir de 2022 (reapresentados desde 2020 para parte deles); antes, Computing & Graphics e Enterprise, Embedded & Semi-Custom. Em 2025 Client e Gaming passaram a ser um segmento (Client and Gaming). Xilinx consolidada a partir de fev/22. O 10-K da AMD não etiqueta a receita por segmento no 4T: o Data Center do 4T é calculado como receita total menos Client and Gaming e Embedded; o lucro por segmento do 4T fica vazio. O 4T22 de Client/Gaming/Data Center fica vazio (abertura do 10-K de 2022 inconsistente com os 9M).'
C17 = 'Check Receita (linha 15 - DC + Client/Gaming + Embedded)'
RC, OI = 'RevenueFromContractWithCustomerExcludingAssessedTax', 'OperatingIncomeLoss'
def items(co):
    q = co.dq
    cg_r, cg_o = q(RC, ['amd:ClientAndGamingMember']), q(OI, ['amd:ClientAndGamingMember'])
    cl_r, ga_r, cl_o, ga_o = q(RC, ['amd:ClientMember']), q(RC, ['amd:GamingMember']), q(OI, ['amd:ClientMember']), q(OI, ['amd:GamingMember'])
    for p in cl_r:
        if p in ga_r: cg_r.setdefault(p, cl_r[p] + ga_r[p])
    for p in cl_o:
        if p in ga_o: cg_o.setdefault(p, cl_o[p] + ga_o[p])
    for d_ in (cg_r, cg_o, cl_r, ga_r, cl_o, ga_o): d_.pop('4Q22', None)   # 4T22: Client/Gaming do 10-K de 2022 não fecham com o 9M original
    dc, em = q(RC, ['amd:DataCenterMember']), q(RC, ['amd:EmbeddedMember'])
    tot = co.quarter([RC, 'Revenues'])
    for p in tot:                                 # 4T: o 10-K não etiqueta Data Center -> receita total - Client and Gaming - Embedded
        if p not in dc and p in cg_r and p in em and p >= '1Q22'[0] + p[1:2] and p[2:] >= '22': dc[p] = tot[p] - cg_r[p] - em[p]
    return [
     ('s_dc', 'Data Center - receita (USD; 4T = receita total - demais segmentos)', 'moeda', dc, '#,##0', 'sum'),
     ('s_cg', 'Client and Gaming - receita (USD; até 2024 = Client + Gaming)', 'moeda', cg_r, '#,##0', 'sum'),
     ('s_cl', 'Client - receita (USD; até 2024)', 'moeda', cl_r, '#,##0', 'sum'),
     ('s_ga', 'Gaming - receita (USD; até 2024)', 'moeda', ga_r, '#,##0', 'sum'),
     ('s_em', 'Embedded - receita (USD)', 'moeda', em, '#,##0', 'sum'),
     ('o_dc', 'Data Center - lucro operacional (USD)', 'moeda', q(OI, ['amd:DataCenterMember']), '#,##0', 'sum'),
     ('o_cg', 'Client and Gaming - lucro operacional (USD)', 'moeda', cg_o, '#,##0', 'sum'),
     ('o_em', 'Embedded - lucro operacional (USD)', 'moeda', q(OI, ['amd:EmbeddedMember']), '#,##0', 'sum'),
     ('x_cgs', 'Computing and Graphics - receita (USD; até 1T22)', 'moeda', q(RC, ['amd:ComputingandGraphicsMember']), '#,##0', 'sum'),
     ('x_ees', 'Enterprise, Embedded and Semi-Custom - receita (USD; até 1T22)', 'moeda', q(RC, ['amd:EnterpriseEmbeddedandSemiCustomMember']), '#,##0', 'sum'),
    ]
BLOCK = r"""
hdr_row(rr, 'Segmentos (USD milhões)'); rr += 1
line('s_dc', '   Data Center'); line('s_cg', '   Client and Gaming'); line('s_cl', '      Client'); line('s_ga', '      Gaming'); line('s_em', '   Embedded')
line('tot', 'Receita (soma dos segmentos)', annual='calc', fn=sumf('s_dc', 's_cg', 's_em'), b=True)
line('chk', 'Check vs Net Revenues (linha 15)', annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["tot"]})=0,"",ROUND({CL(c)}15-{CL(c)}{R["tot"]},1))')
line('y_dc', '   % YoY Data Center', '0.0%', 'calc', yoy(R['s_dc']))
line('p_dc', '   % Data Center na receita', '0.0%', 'calc', ratio(R['s_dc'], R['tot']))
line('o_dc', '   Lucro operacional Data Center'); line('o_cg', '   Lucro operacional Client and Gaming'); line('o_em', '   Lucro operacional Embedded')
line('m_dc', '   Margem operacional Data Center', '0.0%', 'calc', ratio(R['o_dc'], R['s_dc']))
rr += 1
hdr_row(rr, 'Segmentos antigos (USD milhões; até 1T22)'); rr += 1
line('x_cgs', '   Computing and Graphics'); line('x_ees', '   Enterprise, Embedded and Semi-Custom')
"""

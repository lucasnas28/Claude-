DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/msft/dl'
TITLE = 'DRIVERS DE RECEITA (codigos 45+): segmentos (receita e lucro operacional), Microsoft Cloud, linhas de produto e EUA x exterior - SEC XBRL (USD, fator 0,000001).'
NOTE = 'Notas: valores dimensionais das instâncias XBRL; ano fiscal termina em junho, trimestre sem 3M publicado = acumulado menos o acumulado anterior. Linhas de produto mudam de taxonomia (Office -> Microsoft 365 a partir de 2023; Windows + Devices -> Windows and Devices); a soma das linhas de produto não é usada como check. Segmentos: a Microsoft reapresentou os segmentos em ago/24 só a partir de jul/22; trimestres de fechamento de ano fiscal anteriores ficam na base original (cada trimestre soma a receita total).'
C17 = 'Check Receita (linha 15 - soma dos segmentos)'
RV, OI = 'RevenueFromContractWithCustomerExcludingAssessedTax', 'OperatingIncomeLoss'
PB, IC, PC_ = 'msft:ProductivityAndBusinessProcessesMember', 'msft:IntelligentCloudMember', 'msft:MorePersonalComputingMember'
def items(co):
    q = co.dq
    m365c, m365k, off = q(RV, ['msft:MicrosoftThreeSixFiveCommercialProductsAndCloudServicesMember']), q(RV, ['msft:MicrosoftThreeSixFiveConsumerProductsAndCloudServicesMember']), q(RV, ['msft:OfficeProductsAndCloudServicesMember'])
    m365 = {p: m365c[p] + m365k.get(p, 0) for p in m365c}
    for p, v in off.items(): m365.setdefault(p, v)
    win = q(RV, ['msft:WindowsAndDevicesMember'])
    w0, dv = q(RV, ['msft:WindowsMember']), q(RV, ['msft:DevicesMember'])
    for p in w0: win.setdefault(p, w0[p] + dv.get(p, 0))
    return [
     ('s_pb', 'Productivity and Business Processes - receita (USD)', 'moeda', q(RV, [PB]), '#,##0', 'sum'),
     ('s_ic', 'Intelligent Cloud - receita (USD)', 'moeda', q(RV, [IC]), '#,##0', 'sum'),
     ('s_pc', 'More Personal Computing - receita (USD)', 'moeda', q(RV, [PC_]), '#,##0', 'sum'),
     ('o_pb', 'Productivity and Business Processes - lucro operacional (USD)', 'moeda', q(OI, [PB]), '#,##0', 'sum'),
     ('o_ic', 'Intelligent Cloud - lucro operacional (USD)', 'moeda', q(OI, [IC]), '#,##0', 'sum'),
     ('o_pc', 'More Personal Computing - lucro operacional (USD)', 'moeda', q(OI, [PC_]), '#,##0', 'sum'),
     ('mc', 'Microsoft Cloud - receita (USD; até 2021 "Commercial cloud")', 'moeda', q(RV, ['msft:MicrosoftCloudMember', 'msft:CommercialCloudMember']), '#,##0', 'sum'),
     ('p_srv', 'Server products and cloud services (Azure) - receita (USD)', 'moeda', q(RV, ['msft:ServerProductsAndCloudServicesMember']), '#,##0', 'sum'),
     ('p_m365', 'Microsoft 365 comercial + consumo (até 2022: Office products and cloud services) - receita (USD)', 'moeda', m365, '#,##0', 'sum'),
     ('p_li', 'LinkedIn - receita (USD)', 'moeda', q(RV, ['msft:LinkedInCorporationMember']), '#,##0', 'sum'),
     ('p_dyn', 'Dynamics - receita (USD)', 'moeda', q(RV, ['msft:DynamicsProductsAndCloudServicesMember', 'msft:DynamicsMember']), '#,##0', 'sum'),
     ('p_gam', 'Gaming - receita (USD)', 'moeda', q(RV, ['msft:GamingMember']), '#,##0', 'sum'),
     ('p_win', 'Windows and Devices (até 2022: Windows + Devices) - receita (USD)', 'moeda', win, '#,##0', 'sum'),
     ('p_ads', 'Search and news advertising - receita (USD)', 'moeda', q(RV, ['msft:SearchAndNewsAdvertisingMember', 'msft:SearchAdvertisingMember']), '#,##0', 'sum'),
     ('g_us', 'Receita Estados Unidos (USD)', 'moeda', q(RV, ['country:US']), '#,##0', 'sum'),
     ('g_xu', 'Receita fora dos EUA (USD)', 'moeda', q(RV, ['us-gaap:NonUsMember']), '#,##0', 'sum'),
    ]
BLOCK = r"""
hdr_row(rr, 'Segmentos - como a Microsoft divulga (USD milhões)'); rr += 1
line('s_pb', 'Productivity and Business Processes'); line('s_ic', 'Intelligent Cloud'); line('s_pc', 'More Personal Computing')
line('tot', 'Receita (soma dos segmentos)', annual='calc', fn=sumf('s_pb', 's_ic', 's_pc'), b=True)
line('chk', 'Check vs Net Revenues (linha 15)', annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["tot"]})=0,"",ROUND({CL(c)}15-{CL(c)}{R["tot"]},1))')
line('y_ic', '   % YoY Intelligent Cloud', '0.0%', 'calc', yoy(R['s_ic']))
line('o_pb', '   Lucro operacional P&BP'); line('o_ic', '   Lucro operacional Intelligent Cloud'); line('o_pc', '   Lucro operacional MPC')
line('m_pb', '   Margem operacional P&BP', '0.0%', 'calc', ratio(R['o_pb'], R['s_pb']))
line('m_ic', '   Margem operacional Intelligent Cloud', '0.0%', 'calc', ratio(R['o_ic'], R['s_ic']))
line('m_pc', '   Margem operacional MPC', '0.0%', 'calc', ratio(R['o_pc'], R['s_pc']))
rr += 1
hdr_row(rr, 'Linhas de produto (USD milhões)'); rr += 1
line('mc', 'Microsoft Cloud', b=True)
line('y_mc', '   % YoY Microsoft Cloud', '0.0%', 'calc', yoy(R['mc']))
line('p_srv', '   Server products and cloud services (Azure)'); line('y_srv', '   % YoY Server products and cloud', '0.0%', 'calc', yoy(R['p_srv']))
line('p_m365', '   Microsoft 365 / Office'); line('p_li', '   LinkedIn'); line('p_dyn', '   Dynamics'); line('p_gam', '   Gaming'); line('p_win', '   Windows and Devices'); line('p_ads', '   Search and news advertising')
rr += 1
hdr_row(rr, 'Receita por região (USD milhões)'); rr += 1
line('g_us', '   Estados Unidos'); line('g_xu', '   Fora dos EUA')
line('p_us', '   % EUA na receita', '0.0%', 'calc', ratio(R['g_us'], R['tot']))
"""

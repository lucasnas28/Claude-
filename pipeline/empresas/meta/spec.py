DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/meta/dl'
TITLE = 'DRIVERS DE RECEITA (codigos 45+): publicidade x outras receitas, segmentos Family of Apps e Reality Labs (receita e lucro operacional) e regiões - SEC XBRL (USD, fator 0,000001).'
NOTE = 'Notas: valores dimensionais das instâncias XBRL. Segmentos (FoA e RL) só existem a partir do 4T20; antes a Meta abria só publicidade e outras receitas. Membros fb: (até 2021) e meta: costurados.'
C17 = 'Check Receita (linha 15 - publicidade + outras + Reality Labs)'
RV, OI = 'RevenueFromContractWithCustomerExcludingAssessedTax', 'OperatingIncomeLoss'
FOA, RL = 'meta:FamilyOfAppsMember', 'meta:RealityLabsMember'
def items(co):
    q = co.dq
    ads = q(RV, [(FOA, 'us-gaap:AdvertisingMember'), 'us-gaap:AdvertisingMember'])
    oth = q(RV, [(FOA, 'us-gaap:ServiceOtherMember')])
    rl = q(RV, [RL])
    so = q(RV, ['us-gaap:ServiceOtherMember'])          # até 3T20: outras receitas incluíam hardware (Oculus/Portal)
    for p, v in so.items():
        if p not in oth and p not in rl: oth[p] = v
    return [
     ('ads', 'Receita de publicidade (USD)', 'moeda', ads, '#,##0', 'sum'),
     ('oth', 'Outras receitas Family of Apps (USD; até 3T20 inclui hardware, antes da criação do segmento Reality Labs)', 'moeda', oth, '#,##0', 'sum'),
     ('s_rl', 'Reality Labs - receita (USD)', 'moeda', rl, '#,##0', 'sum'),
     ('s_foa', 'Family of Apps - receita (USD)', 'moeda', q(RV, [FOA]), '#,##0', 'sum'),
     ('o_foa', 'Family of Apps - lucro operacional (USD)', 'moeda', q(OI, [FOA]), '#,##0', 'sum'),
     ('o_rl', 'Reality Labs - lucro operacional (USD)', 'moeda', q(OI, [RL]), '#,##0', 'sum'),
     ('g_usc', 'Receita EUA e Canadá (USD)', 'moeda', q(RV, ['meta:USCanadaMember', 'fb:USCanadaMember']), '#,##0', 'sum'),
     ('g_eu', 'Receita Europa (USD)', 'moeda', q(RV, ['srt:EuropeMember']), '#,##0', 'sum'),
     ('g_ap', 'Receita Ásia-Pacífico (USD)', 'moeda', q(RV, ['srt:AsiaPacificMember']), '#,##0', 'sum'),
     ('g_rw', 'Receita resto do mundo (USD)', 'moeda', q(RV, ['meta:RestOfWorldMember', 'fb:RestOfWorldMember']), '#,##0', 'sum'),
    ]
BLOCK = r"""
hdr_row(rr, 'Receita - como a Meta divulga (USD milhões)'); rr += 1
line('ads', 'Publicidade', b=True); line('oth', '   Outras receitas (Family of Apps)'); line('s_rl', '   Reality Labs')
line('tot', 'Receita (soma)', annual='calc', fn=sumf('ads', 'oth', 's_rl'), b=True)
line('chk', 'Check vs Net Revenues (linha 15)', annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["tot"]})=0,"",ROUND({CL(c)}15-{CL(c)}{R["tot"]},1))')
line('y_ads', '   % YoY publicidade', '0.0%', 'calc', yoy(R['ads']))
line('p_ads', '   % publicidade na receita', '0.0%', 'calc', ratio(R['ads'], R['tot']))
rr += 1
hdr_row(rr, 'Segmentos (USD milhões; a partir do 4T20)'); rr += 1
line('s_foa', '   Receita Family of Apps')
line('o_foa', '   Lucro operacional Family of Apps'); line('o_rl', '   Lucro (prejuízo) operacional Reality Labs')
line('m_foa', '   Margem operacional Family of Apps', '0.0%', 'calc', ratio(R['o_foa'], R['s_foa']))
rr += 1
hdr_row(rr, 'Receita por região (USD milhões; pela localização do usuário)'); rr += 1
line('g_usc', '   EUA e Canadá'); line('g_eu', '   Europa'); line('g_ap', '   Ásia-Pacífico'); line('g_rw', '   Resto do mundo')
line('y_ap', '   % YoY Ásia-Pacífico', '0.0%', 'calc', yoy(R['g_ap']))
"""

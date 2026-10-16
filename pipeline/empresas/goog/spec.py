DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/goog/dl'
TITLE = 'DRIVERS DE RECEITA (codigos 45+): receita por tipo (Search, YouTube, Network, Subscriptions, Cloud, Other Bets), segmentos (receita e lucro operacional) e regiões - SEC XBRL (USD, fator 0,000001).'
NOTE = 'Notas: valores dimensionais das instâncias XBRL; trimestre sem 3M publicado = acumulado do ano menos o acumulado anterior. Os nomes dos membros mudaram ao longo do tempo (GoogleInc, GoogleServicesMember, GoogleServicesSegmentMember) e foram costurados.'
C17 = 'Check Receita (linha 15 - soma dos tipos de receita)'
RV, OI = 'RevenueFromContractWithCustomerExcludingAssessedTax', 'OperatingIncomeLoss'
GS = ['goog:GoogleServicesSegmentMember', 'goog:GoogleServicesMember', 'goog:GoogleInc.Member']
def T(prod): return [(prod, s) for s in GS]
def items(co):
    q = co.dq
    return [
     ('t_srch', 'Receita Google Search & other (USD)', 'moeda', q(RV, T('goog:GoogleSearchOtherMember')), '#,##0', 'sum'),
     ('t_yt', 'Receita YouTube ads (USD)', 'moeda', q(RV, T('goog:YouTubeAdvertisingRevenueMember')), '#,##0', 'sum'),
     ('t_net', 'Receita Google Network (USD)', 'moeda', q(RV, T('goog:GoogleNetworkMember') + T('goog:GoogleNetworkMembersPropertiesMember')), '#,##0', 'sum'),
     ('t_sub', 'Receita Google subscriptions, platforms and devices (USD; até 2021 "Google other")', 'moeda', q(RV, T('goog:SubscriptionsPlatformsAndDevicesRevenueMember') + T('goog:OtherRevenuesMember') + ['goog:SubscriptionsPlatformsAndDevicesRevenueMember', 'goog:OtherRevenuesMember']), '#,##0', 'sum'),
     ('s_gs', 'Segmento Google Services - receita (USD)', 'moeda', q(RV, GS[:2]), '#,##0', 'sum'),
     ('s_gc', 'Segmento Google Cloud - receita (USD)', 'moeda', q(RV, ['goog:GoogleCloudSegmentMember', 'goog:GoogleCloudMember']), '#,##0', 'sum'),
     ('s_ob', 'Segmento Other Bets - receita (USD)', 'moeda', q(RV, ['us-gaap:AllOtherSegmentsMember']), '#,##0', 'sum'),
     ('o_gs', 'Google Services - lucro operacional (USD)', 'moeda', q(OI, GS[:2]), '#,##0', 'sum'),
     ('o_gc', 'Google Cloud - lucro operacional (USD)', 'moeda', q(OI, ['goog:GoogleCloudSegmentMember', 'goog:GoogleCloudMember']), '#,##0', 'sum'),
     ('o_ob', 'Other Bets - lucro operacional (USD)', 'moeda', q(OI, ['us-gaap:AllOtherSegmentsMember']), '#,##0', 'sum'),
     ('o_ab', 'Alphabet-level activities - resultado operacional (USD)', 'moeda', q(OI, ['us-gaap:CorporateNonSegmentMember']), '#,##0', 'sum'),
     ('g_us', 'Receita Estados Unidos (USD)', 'moeda', q(RV, ['country:US']), '#,##0', 'sum'),
     ('g_em', 'Receita EMEA (USD)', 'moeda', q(RV, ['us-gaap:EMEAMember']), '#,##0', 'sum'),
     ('g_ap', 'Receita APAC (USD)', 'moeda', q(RV, ['srt:AsiaPacificMember']), '#,##0', 'sum'),
     ('g_oa', 'Receita Outras Américas (USD)', 'moeda', q(RV, ['goog:AmericasExcludingUnitedStatesMember']), '#,##0', 'sum'),
    ]
BLOCK = r"""
hdr_row(rr, 'Receita por tipo - como a Alphabet divulga (USD milhões)'); rr += 1
line('t_srch', '   Google Search & other'); line('t_yt', '   YouTube ads'); line('t_net', '   Google Network')
line('adv', 'Google advertising', annual='calc', fn=sumf('t_srch', 't_yt', 't_net'), b=True)
line('t_sub', '   Subscriptions, platforms and devices')
line('s_gc', 'Google Cloud'); line('s_ob', 'Other Bets')
line('tot', 'Receita (soma dos tipos; a diferença para a linha 15 são os ganhos de hedge)', annual='calc', fn=sumf('adv', 't_sub', 's_gc', 's_ob'), b=True)
line('chk', 'Hedge / diferença vs Net Revenues (linha 15)', annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["tot"]})=0,"",ROUND({CL(c)}15-{CL(c)}{R["tot"]},1))')
line('y_srch', '   % YoY Search', '0.0%', 'calc', yoy(R['t_srch']))
line('y_yt', '   % YoY YouTube ads', '0.0%', 'calc', yoy(R['t_yt']))
line('y_gc', '   % YoY Google Cloud', '0.0%', 'calc', yoy(R['s_gc']))
rr += 1
hdr_row(rr, 'Segmentos - receita e lucro operacional (USD milhões)'); rr += 1
line('s_gs', '   Receita Google Services')
line('o_gs', '   Lucro operacional Google Services'); line('o_gc', '   Lucro operacional Google Cloud'); line('o_ob', '   Lucro operacional Other Bets'); line('o_ab', '   Alphabet-level activities')
line('m_gs', '   Margem operacional Google Services', '0.0%', 'calc', ratio(R['o_gs'], R['s_gs']))
line('m_gc', '   Margem operacional Google Cloud', '0.0%', 'calc', ratio(R['o_gc'], R['s_gc']))
rr += 1
hdr_row(rr, 'Receita por região (USD milhões)'); rr += 1
line('g_us', '   Estados Unidos'); line('g_em', '   EMEA'); line('g_ap', '   APAC'); line('g_oa', '   Outras Américas')
line('p_us', '   % EUA na receita', '0.0%', 'calc', ratio(R['g_us'], R['tot']))
"""

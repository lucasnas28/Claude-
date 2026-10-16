DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/ko/dl'
TITLE = 'DRIVERS DE RECEITA (codigos 45+): segmentos operacionais (receita e lucro operacional), concentrado x produto acabado e EUA x internacional - SEC XBRL (USD, fator 0,000001).'
NOTE = 'Notas: valores dimensionais das instâncias XBRL. Segmentos EMEA/América Latina/América do Norte com membros novos a partir de 2018 (antes us-gaap:EMEA etc.); Global Ventures até 2024; Ásia-Pacífico ("Pacific"). Eliminações e Corporate fecham a receita. Segmentos pelo valor originalmente divulgado: em 2025 a Coca-Cola extinguiu Global Ventures e reapresentou 2024 sem ele.'
C17 = 'Check Receita (linha 15 - soma dos segmentos + corporate + eliminações)'
RV, OI = 'Revenues', 'OperatingIncomeLoss'
SEG = [('emea', 'EMEA', ['ko:EuropeMiddleEastAfricaMember', 'us-gaap:EMEAMember']), ('lat', 'América Latina', ['ko:LatinAmericaSegmentMember', 'srt:LatinAmericaMember']),
       ('na', 'América do Norte', ['ko:NorthAmericaSegmentMember', 'srt:NorthAmericaMember']), ('ap', 'Ásia-Pacífico', ['ko:A.PacificMember', 'ko:PacificMember', 'srt:AsiaPacificMember']),
       ('gv', 'Global Ventures', ['ko:GlobalVenturesMember']), ('bi', 'Bottling Investments', ['ko:BottlingInvestmentsMember']),
       ('co', 'Corporate', ['us-gaap:CorporateMember', 'us-gaap:CorporateNonSegmentMember']), ('el', 'Eliminações', ['ko:EliminationsMember', 'srt:ConsolidationEliminationsMember', 'us-gaap:IntersegmentEliminationMember'])]
def items(co):
    out = []
    for k, nm, mm in SEG:
        out.append(('s_' + k, f'Receita {nm} (USD)', 'moeda', co.dq(RV, mm, orig=True), '#,##0', 'sum'))
        if k != 'el': out.append(('o_' + k, f'Lucro operacional {nm} (USD)', 'moeda', co.dq(OI, mm, orig=True), '#,##0', 'sum'))
    out += [('t_con', 'Receita operações de concentrado (USD; desde 2021)', 'moeda', co.dq(RV, ['ko:ConcentrateOperationsMember']), '#,##0', 'sum'),
            ('t_fin', 'Receita operações de produto acabado (USD; desde 2021)', 'moeda', co.dq(RV, ['ko:FinishedProductOperationsMember']), '#,##0', 'sum'),
            ('g_us', 'Receita EUA (USD)', 'moeda', co.dq(RV, ['country:US']), '#,##0', 'sum')]
    return out
BLOCK = r"""
hdr_row(rr, 'Segmentos - receita (USD milhões)'); rr += 1
for k, nm, mm in SEG: line('s_' + k, '   ' + nm)
line('tot', 'Receita (soma)', annual='calc', fn=sumf(*['s_' + k for k, nm, mm in SEG]), b=True)
line('chk', 'Check vs Net Revenues (linha 15)', annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["tot"]})=0,"",ROUND({CL(c)}15-{CL(c)}{R["tot"]},1))')
rr += 1
hdr_row(rr, 'Segmentos - lucro operacional (USD milhões)'); rr += 1
for k, nm, mm in SEG[:-1]: line('o_' + k, '   ' + nm)
line('m_na', '   Margem operacional América do Norte', '0.0%', 'calc', ratio(R['o_na'], R['s_na']))
line('m_lat', '   Margem operacional América Latina', '0.0%', 'calc', ratio(R['o_lat'], R['s_lat']))
rr += 1
hdr_row(rr, 'Modelo de negócio e região (USD milhões)'); rr += 1
line('t_con', '   Operações de concentrado'); line('t_fin', '   Operações de produto acabado')
line('p_con', '   % concentrado na receita', '0.0%', 'calc', ratio(R['t_con'], R['tot']))
line('g_us', '   Receita EUA'); line('p_us', '   % EUA na receita', '0.0%', 'calc', ratio(R['g_us'], R['tot']))
"""

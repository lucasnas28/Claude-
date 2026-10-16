DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/mu/dl'
TITLE = 'DRIVERS DE RECEITA (codigos 45+): receita por tecnologia (DRAM, NAND, outros) e por unidade de negócio (antigas até 2T25 e novas desde 3T23) - SEC XBRL (USD, fator 0,000001).'
NOTE = 'Notas: valores dimensionais das instâncias XBRL; ano fiscal termina no fim de agosto/início de setembro (cada trimestre vai para o fim de trimestre-calendário mais próximo). Em 2025 a Micron trocou as unidades de negócio (CNBU/MBU/SBU/EBU -> CMBU/CDBU/MCBU/AEBU, reapresentadas desde 3T23).'
C17 = 'Check Receita (linha 15 - DRAM + NAND + outros)'
RV, OI = 'RevenueFromContractWithCustomerExcludingAssessedTax', 'OperatingIncomeLoss'
def items(co):
    q = co.dq
    out = [
     ('t_dram', 'Receita DRAM (USD)', 'moeda', q(RV, ['mu:DRAMProductsMember']), '#,##0', 'sum'),
     ('t_nand', 'Receita NAND (USD)', 'moeda', q(RV, ['mu:NANDProductsMember']), '#,##0', 'sum'),
     ('t_oth', 'Receita outros produtos (USD)', 'moeda', q(RV, ['mu:OtherProductSalesMember']), '#,##0', 'sum')]
    for k, m, nm in [('cnbu', 'mu:CNBUMember', 'CNBU (compute & networking, antiga)'), ('mbu', 'mu:MBUMember', 'MBU (mobile, antiga)'), ('sbu', 'mu:SBUMember', 'SBU (storage, antiga)'), ('ebu', 'mu:EBUMember', 'EBU (embedded, antiga)'),
                     ('cmbu', 'mu:CMBUMember', 'CMBU (cloud memory)'), ('cdbu', 'mu:CDBUMember', 'CDBU (core data center)'), ('mcbu', 'mu:MCBUMember', 'MCBU (mobile & client)'), ('aebu', 'mu:AEBUMember', 'AEBU (automotive & embedded)')]:
        out.append(('b_' + k, f'Receita {nm} (USD)', 'moeda', q(RV, [m]), '#,##0', 'sum'))
        if k in ('cnbu', 'mbu', 'sbu', 'ebu'): out.append(('o_' + k, f'Lucro operacional {nm} (USD)', 'moeda', q(OI, [m]), '#,##0', 'sum'))
    return out
BLOCK = r"""
hdr_row(rr, 'Receita por tecnologia (USD milhões)'); rr += 1
line('t_dram', '   DRAM'); line('t_nand', '   NAND'); line('t_oth', '   Outros')
line('tot', 'Receita (soma)', annual='calc', fn=sumf('t_dram', 't_nand', 't_oth'), b=True)
line('chk', 'Check vs Net Revenues (linha 15)', annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["tot"]})=0,"",ROUND({CL(c)}15-{CL(c)}{R["tot"]},1))')
line('p_dram', '   % DRAM na receita', '0.0%', 'calc', ratio(R['t_dram'], R['tot']))
line('y_dram', '   % YoY DRAM', '0.0%', 'calc', yoy(R['t_dram']))
line('y_nand', '   % YoY NAND', '0.0%', 'calc', yoy(R['t_nand']))
rr += 1
hdr_row(rr, 'Unidades de negócio - novas (desde 3T23, USD milhões)'); rr += 1
line('b_cmbu', '   Cloud Memory (CMBU)'); line('b_cdbu', '   Core Data Center (CDBU)'); line('b_mcbu', '   Mobile and Client (MCBU)'); line('b_aebu', '   Automotive and Embedded (AEBU)')
rr += 1
hdr_row(rr, 'Unidades de negócio - antigas (até 2T25, USD milhões)'); rr += 1
line('b_cnbu', '   Compute and Networking (CNBU)'); line('b_mbu', '   Mobile (MBU)'); line('b_sbu', '   Storage (SBU)'); line('b_ebu', '   Embedded (EBU)')
line('o_cnbu', '   Lucro operacional CNBU'); line('o_mbu', '   Lucro operacional MBU'); line('o_sbu', '   Lucro operacional SBU'); line('o_ebu', '   Lucro operacional EBU')
line('m_cnbu', '   Margem operacional CNBU', '0.0%', 'calc', ratio(R['o_cnbu'], R['b_cnbu']))
"""

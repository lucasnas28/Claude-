DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/ma/dl'
TITLE = 'DRIVERS DE RECEITA (codigos 45+): Payment Network x Value-Added Services and Solutions (desde 2020) e regiões (América do Norte x Internacional) - SEC XBRL (USD, fator 0,000001).'
NOTE = 'Notas: valores dimensionais das instâncias XBRL. Até 2022 a Mastercard também abria a receita por tipo de tarifa (domestic assessments, cross-border, processamento, outras), mantida como memória.'
C17 = 'Check Receita (linha 15 - Payment Network + VAS)'
RV, RC = 'Revenues', 'RevenueFromContractWithCustomerExcludingAssessedTax'
def items(co):
    def q(mem):
        out = co.dq(RV, [mem])
        for k, v in co.dq(RC, [mem]).items(): out.setdefault(k, v)
        return out
    return [
     ('s_pn', 'Payment Network - receita (USD)', 'moeda', q('ma:PaymentNetworkMember'), '#,##0', 'sum'),
     ('s_vas', 'Value-Added Services and Solutions - receita (USD)', 'moeda', q('ma:ValueAddedServicesAndSolutionsMember'), '#,##0', 'sum'),
     ('g_na', 'Receita América do Norte (USD)', 'moeda', q('srt:NorthAmericaMember'), '#,##0', 'sum'),
     ('g_im', 'Receita International Markets (USD)', 'moeda', q('ma:InternationalMarketsMember'), '#,##0', 'sum'),
     ('f_dom', 'Domestic assessments (USD; até 2022)', 'moeda', q('ma:DomesticAssessmentsMember'), '#,##0', 'sum'),
     ('f_xb', 'Cross-border volume fees (USD; até 2022)', 'moeda', q('ma:CrossborderVolumeFeesMember'), '#,##0', 'sum'),
     ('f_tp', 'Transaction processing (USD; até 2022)', 'moeda', q('ma:TransactionProcessingMember'), '#,##0', 'sum'),
    ]
BLOCK = r"""
hdr_row(rr, 'Receita por linha (USD milhões; desde 2020)'); rr += 1
line('s_pn', '   Payment Network'); line('s_vas', '   Value-Added Services and Solutions')
line('tot', 'Receita (soma)', annual='calc', fn=sumf('s_pn', 's_vas'), b=True)
line('chk', 'Check vs Net Revenues (linha 15)', annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["tot"]})=0,"",ROUND({CL(c)}15-{CL(c)}{R["tot"]},1))')
line('y_vas', '   % YoY VAS', '0.0%', 'calc', yoy(R['s_vas']))
line('p_vas', '   % VAS na receita', '0.0%', 'calc', ratio(R['s_vas'], R['tot']))
rr += 1
hdr_row(rr, 'Região e tarifas (USD milhões)'); rr += 1
line('g_na', '   América do Norte'); line('g_im', '   International Markets')
line('f_dom', '   Domestic assessments (até 2022)'); line('f_xb', '   Cross-border volume fees (até 2022)'); line('f_tp', '   Transaction processing (até 2022)')
"""

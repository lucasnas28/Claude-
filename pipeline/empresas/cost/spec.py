DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/cost/dl'
MONTHMAP = {11: '4', 12: '4', 2: '1', 3: '1', 5: '2', 6: '2', 8: '3', 9: '3'}
TITLE = 'DRIVERS DE RECEITA (codigos 45+): vendas líquidas x taxas de associação e segmentos geográficos (EUA, Canadá, outros internacionais: receita e lucro operacional) - SEC XBRL (USD, fator 0,000001).'
NOTE = 'Notas: ano fiscal da Costco termina no fim de agosto/início de setembro, com trimestres de 12 semanas e o último de 16 semanas (fica no 3T calendário). Colunas: fechamento de nov -> 4T, fev -> 1T, mai -> 2T, ago/set -> 3T. Membros dos segmentos mudaram de nome em 2023-24 e foram costurados.'
C17 = 'Check Receita (linha 15 - vendas + associação)'
RC, RV, OI = 'RevenueFromContractWithCustomerExcludingAssessedTax', 'Revenues', 'OperatingIncomeLoss'
def items(co):
    def m(tag_list, mems):
        out = {}
        for t in tag_list:
            for k, v in co.dq(t, mems).items(): out.setdefault(k, v)
        return out
    return [
     ('t_sal', 'Vendas líquidas (USD)', 'moeda', m([RC], ['us-gaap:ProductMember']), '#,##0', 'sum'),
     ('t_mem', 'Taxas de associação (USD)', 'moeda', m([RC], ['us-gaap:MembershipMember']), '#,##0', 'sum'),
     ('s_us', 'Receita EUA (USD)', 'moeda', m([RV, RC], ['cost:UnitedStatesMember', 'country:US']), '#,##0', 'sum'),
     ('s_ca', 'Receita Canadá (USD)', 'moeda', m([RV, RC], ['cost:CanadaMember', 'country:CA']), '#,##0', 'sum'),
     ('s_oi', 'Receita outros internacionais (USD)', 'moeda', m([RV, RC], ['cost:OtherInternationalMember', 'cost:OtherInternationalOperationsMember']), '#,##0', 'sum'),
     ('o_us', 'Lucro operacional EUA (USD)', 'moeda', m([OI], ['cost:UnitedStatesMember', 'country:US']), '#,##0', 'sum'),
     ('o_ca', 'Lucro operacional Canadá (USD)', 'moeda', m([OI], ['cost:CanadaMember', 'country:CA']), '#,##0', 'sum'),
     ('o_oi', 'Lucro operacional outros internacionais (USD)', 'moeda', m([OI], ['cost:OtherInternationalMember', 'cost:OtherInternationalOperationsMember']), '#,##0', 'sum'),
    ]
BLOCK = r"""
hdr_row(rr, 'Receita - como a Costco divulga (USD milhões)'); rr += 1
line('t_sal', '   Vendas líquidas'); line('t_mem', '   Taxas de associação')
line('tot', 'Receita total (soma)', annual='calc', fn=sumf('t_sal', 't_mem'), b=True)
line('chk', 'Check vs Net Revenues (linha 15)', annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["tot"]})=0,"",ROUND({CL(c)}15-{CL(c)}{R["tot"]},1))')
line('y_mem', '   % YoY taxas de associação', '0.0%', 'calc', yoy(R['t_mem']))
line('p_mem', '   % associação na receita', '0.0%', 'calc', ratio(R['t_mem'], R['tot']))
rr += 1
hdr_row(rr, 'Segmentos geográficos (USD milhões)'); rr += 1
line('s_us', '   Receita EUA'); line('s_ca', '   Receita Canadá'); line('s_oi', '   Receita outros internacionais')
line('o_us', '   Lucro operacional EUA'); line('o_ca', '   Lucro operacional Canadá'); line('o_oi', '   Lucro operacional outros internacionais')
line('m_us', '   Margem operacional EUA', '0.0%', 'calc', ratio(R['o_us'], R['s_us']))
"""

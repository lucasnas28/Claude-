DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/xom/dl'
TITLE = 'DRIVERS DE RECEITA (codigos 45+): composição da receita (vendas, equivalência, outras) e lucro por segmento (Upstream, Energy Products, Chemical Products, Specialty Products), EUA + exterior - SEC XBRL (USD, fator 0,000001).'
NOTE = 'Notas: valores dimensionais das instâncias XBRL. A linha "Revenues" da Exxon inclui equivalência patrimonial e outras receitas; o lucro operacional do modelo é receita total - custos totais + juros. Segmentos novos (Energy/Chemical/Specialty Products) a partir do 4T20; Upstream desde sempre. Lucro por segmento = EUA + exterior.'
C17 = 'Check Receita (linha 15 - vendas + equivalência + outras)'
RV, NI = 'Revenues', 'NetIncomeLoss'
def items(co):
    def seg(tag, mem, alt_us=None, alt_xu=None):
        out = {}
        us, xu = co.dq(tag, [(mem, 'country:US')]), co.dq(tag, [(mem, 'us-gaap:NonUsMember')])
        for p in us:
            if p in xu: out[p] = us[p] + xu[p]
        if alt_us:
            a, b = co.dq(tag, [alt_us]), co.dq(tag, [alt_xu])
            for p in a:
                if p in b: out.setdefault(p, a[p] + b[p])
        return out
    q = co.dq
    return [
     ('r_sal', 'Vendas e outras receitas operacionais (USD)', 'moeda', q(RV, ['xom:SalesAndOtherOperatingRevenueMember']), '#,##0', 'sum'),
     ('r_eq', 'Receita de equivalência patrimonial (USD)', 'moeda', q(RV, ['xom:IncomeFromEquityAffiliatesMember']), '#,##0', 'sum'),
     ('r_oth', 'Outras receitas (USD)', 'moeda', q(RV, ['xom:OtherRevenueMember']), '#,##0', 'sum'),
     ('e_up', 'Lucro Upstream (EUA + exterior, USD)', 'moeda', seg(NI, 'xom:UpstreamMember', 'xom:UpstreamUnitedStatesMember', 'xom:UpstreamNonUsMember'), '#,##0', 'sum'),
     ('e_ep', 'Lucro Energy Products (USD)', 'moeda', seg(NI, 'xom:EnergyProductsMember'), '#,##0', 'sum'),
     ('e_cp', 'Lucro Chemical Products (USD)', 'moeda', seg(NI, 'xom:ChemicalProductsMember'), '#,##0', 'sum'),
     ('e_sp', 'Lucro Specialty Products (USD)', 'moeda', seg(NI, 'xom:SpecialtyProductsMember'), '#,##0', 'sum'),
     ('e_co', 'Corporate e financeiro (USD)', 'moeda', q(NI, ['us-gaap:CorporateNonSegmentMember', 'xom:CorporateAndFinancingMember']), '#,##0', 'sum'),
     ('s_up', 'Receita Upstream (USD; desde 4T22)', 'moeda', seg(RV, 'xom:UpstreamMember'), '#,##0', 'sum'),
     ('s_ep', 'Receita Energy Products (USD; desde 4T22)', 'moeda', seg(RV, 'xom:EnergyProductsMember'), '#,##0', 'sum'),
    ]
BLOCK = r"""
hdr_row(rr, 'Composição da receita (USD milhões; a partir do 4T20)'); rr += 1
line('r_sal', '   Vendas e outras receitas operacionais'); line('r_eq', '   Equivalência patrimonial'); line('r_oth', '   Outras receitas')
line('tot', 'Receita total (soma)', annual='calc', fn=sumf('r_sal', 'r_eq', 'r_oth'), b=True)
line('chk', 'Check vs Net Revenues (linha 15)', annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["tot"]})=0,"",ROUND({CL(c)}15-{CL(c)}{R["tot"]},1))')
line('s_up', '   Receita Upstream'); line('s_ep', '   Receita Energy Products')
rr += 1
hdr_row(rr, 'Lucro por segmento (USD milhões)'); rr += 1
line('e_up', '   Upstream'); line('e_ep', '   Energy Products'); line('e_cp', '   Chemical Products'); line('e_sp', '   Specialty Products'); line('e_co', '   Corporate e financeiro')
line('etot', 'Lucro líquido (soma dos segmentos; completa só a partir de 2022 - antes a Exxon usava Downstream/Chemical)', annual='calc', fn=sumf('e_up', 'e_ep', 'e_cp', 'e_sp', 'e_co'), b=True)
line('p_up', '   % Upstream no lucro', '0.0%', 'calc', ratio(R['e_up'], R['etot']))
"""

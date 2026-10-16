DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/lly/dl'
TITLE = 'DRIVERS DE RECEITA (codigos 45+): principais produtos (Mounjaro, Zepbound, Trulicity, Verzenio, Jardiance, Taltz), áreas terapêuticas e EUA x exterior - SEC XBRL (USD, fator 0,000001).'
NOTE = 'Notas: valores dimensionais das instâncias XBRL. "Diabetes" (até 2023) virou "Cardiometabolic Health" (com obesidade); os membros foram costurados. Jardiance inclui pagamentos de marco da parceria com a Boehringer quando divulgados no mesmo membro.'
C17 = 'Check Receita (linha 15 - EUA + exterior)'
RV = 'Revenues'
def items(co):
    def q(tag, mems):
        out = {}
        for mm in mems:                              # produto sozinho; no 10-K só vem EUA e exterior separados -> soma
            a = co.dq(tag, [mm]); us = co.dq(tag, [(mm, 'country:US')]); xu = co.dq(tag, [(mm, 'us-gaap:NonUsMember')])
            for p in set(a) | set(us):
                if p in out: continue
                if p in a: out[p] = a[p]
                elif p in us and p in xu: out[p] = us[p] + xu[p]
        return out
    P_ = lambda nm, mem: (nm, f'Receita {nm} (USD)', 'moeda', q(RV, mem), '#,##0', 'sum')
    return [
     P_('Mounjaro', ['lly:MounjaroMember']), P_('Zepbound', ['lly:ZepboundMember']), P_('Trulicity', ['lly:TrulicityMember', 'lly:TrulicityMemberMember']),
     P_('Verzenio', ['lly:VerzenioMember']), P_('Jardiance', ['lly:JardianceMember']), P_('Taltz', ['lly:TaltzMember']), P_('Humalog', ['lly:HumalogMember']),
     ('a_cm', 'Cardiometabolic Health / Diabetes - receita (USD)', 'moeda', q(RV, ['lly:CardiometabolicHealthMember', 'lly:DiabetesMember']), '#,##0', 'sum'),
     ('a_onc', 'Oncologia - receita (USD)', 'moeda', q(RV, ['lly:OncologyMember']), '#,##0', 'sum'),
     ('a_imm', 'Imunologia - receita (USD)', 'moeda', q(RV, ['lly:ImmunologyMember']), '#,##0', 'sum'),
     ('a_neu', 'Neurociência - receita (USD)', 'moeda', q(RV, ['lly:NeuroscienceMember']), '#,##0', 'sum'),
     ('g_us', 'Receita EUA (USD)', 'moeda', co.dq(RV, ['country:US']), '#,##0', 'sum'),
     ('g_xu', 'Receita fora dos EUA (USD)', 'moeda', co.dq(RV, ['us-gaap:NonUsMember']), '#,##0', 'sum'),
     ('t_prod', 'Receita de produtos (USD)', 'moeda', co.dq(RV, ['us-gaap:ProductMember']), '#,##0', 'sum'),
     ('t_col', 'Receita de colaboração e outras (USD)', 'moeda', co.dq(RV, ['lly:CollaborationandOtherRevenueMember']), '#,##0', 'sum'),
    ]
BLOCK = r"""
hdr_row(rr, 'Receita por região e tipo (USD milhões)'); rr += 1
line('g_us', '   EUA'); line('g_xu', '   Fora dos EUA')
line('tot', 'Receita (EUA + exterior)', annual='calc', fn=sumf('g_us', 'g_xu'), b=True)
line('chk', 'Check vs Net Revenues (linha 15)', annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["tot"]})=0,"",ROUND({CL(c)}15-{CL(c)}{R["tot"]},1))')
line('t_prod', '   Receita de produtos'); line('t_col', '   Colaboração e outras')
rr += 1
hdr_row(rr, 'Principais produtos (USD milhões)'); rr += 1
line('Mounjaro', '   Mounjaro (tirzepatida, diabetes)'); line('Zepbound', '   Zepbound (tirzepatida, obesidade)')
line('tirz', 'Tirzepatida (Mounjaro + Zepbound)', annual='calc', fn=sumf('Mounjaro', 'Zepbound'), b=True)
line('p_tirz', '   % tirzepatida na receita', '0.0%', 'calc', ratio(R['tirz'], R['tot']))
line('Trulicity', '   Trulicity'); line('Verzenio', '   Verzenio'); line('Jardiance', '   Jardiance'); line('Taltz', '   Taltz'); line('Humalog', '   Humalog (até 2024)')
rr += 1
hdr_row(rr, 'Áreas terapêuticas (USD milhões)'); rr += 1
line('a_cm', '   Cardiometabolic Health (até 2022: Diabetes)'); line('a_onc', '   Oncologia'); line('a_imm', '   Imunologia'); line('a_neu', '   Neurociência')
line('y_cm', '   % YoY Cardiometabolic', '0.0%', 'calc', yoy(R['a_cm']))
"""

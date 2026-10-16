DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/nflx/dl'
TITLE = 'DRIVERS DE RECEITA (codigos 45+): receita, assinantes pagos e receita média por assinante por região (UCAN, EMEA, LATAM, APAC) - SEC XBRL (receita em USD com fator 0,000001; assinantes em milhões).'
NOTE = 'Notas: valores dimensionais das instâncias XBRL. Assinantes pagos = saldo no fim do trimestre (a Netflix parou de divulgar assinantes em 2025). ARM = receita do trimestre / média de assinantes do trimestre / 3 (USD por mês). DVD até 3T23.'
C17 = 'Check Receita (linha 15 - soma das regiões + DVD)'
RV = 'Revenues'
REG = [('ucan', 'EUA e Canadá', 'nflx:UnitedStatesAndCanadaMember'), ('emea', 'EMEA', 'us-gaap:EMEAMember'), ('lat', 'América Latina', 'srt:LatinAmericaMember'), ('apac', 'Ásia-Pacífico', 'srt:AsiaPacificMember')]
def items(co):
    from sx import label, P
    def inst(tag, mem):
        out = {}
        for (n, ns, s, e, dd, v, u, a, fl, i) in co.facts:
            if n != tag or s is not None: continue
            ms = tuple(sorted(m for ax, m in dd))
            if ms not in ((mem,), tuple(sorted(('nflx:StreamingMember', mem)))): continue
            lab = label(e)
            if lab in P and (lab not in out or fl > out[lab][1]): out[lab] = (v / 1e6, fl)
        return {k: v[0] for k, v in out.items()}
    out = []
    for k, nm, mm in REG:
        out.append(('r_' + k, f'Receita streaming {nm} (USD)', 'moeda', co.dq(RV, [('nflx:StreamingMember', mm), mm]), '#,##0', 'sum'))
        out.append(('m_' + k, f'Assinantes pagos {nm} (milhões, fim do trimestre)', 'un', inst('NumberOfPaidMemberships', mm), '#,##0.0', 'last'))
    dvd = co.dq(RV, ['nflx:DVDMember'])
    tot = co.quarter([RV]); regs = [o[3] for o in out if o[0].startswith('r_')]
    for p in P:                                  # 4T sem DVD publicado isolado: receita total - regiões (só enquanto o DVD existia)
        if p not in dvd and p <= '3Q23'[0] + p[1:] and p[2:] <= '23' and p in tot and all(p in r for r in regs) and (p[2:] < '23'):
            dvd[p] = tot[p] - sum(r[p] for r in regs)
    out.append(('r_dvd', 'Receita DVD (USD; até 3T23; 4T = receita total - regiões)', 'moeda', dvd, '#,##0', 'sum'))
    return out
BLOCK = r"""
hdr_row(rr, 'Receita por região (USD milhões)'); rr += 1
for k, nm, mm in REG: line('r_' + k, '   ' + nm)
line('r_dvd', '   DVD')
line('tot', 'Receita (soma)', annual='calc', fn=sumf(*(['r_' + k for k, nm, mm in REG] + ['r_dvd'])), b=True)
line('chk', 'Check vs Net Revenues (linha 15)', annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["tot"]})=0,"",ROUND({CL(c)}15-{CL(c)}{R["tot"]},1))')
line('y_tot', '   % YoY receita', '0.0%', 'calc', yoy(R['tot']))
rr += 1
hdr_row(rr, 'Assinantes pagos (milhões, fim do trimestre)'); rr += 1
for k, nm, mm in REG: line('m_' + k, '   ' + nm, '#,##0.0', 'last')
line('mtot', 'Total de assinantes pagos', '#,##0.0', 'calc', fn=lambda c: '=IF(COUNT(' + ','.join(f'{CL(c)}{R["m_" + k]}' for k, nm, mm in REG) + ')=0,"",' + '+'.join(f'N({CL(c)}{R["m_" + k]})' for k, nm, mm in REG) + ')', b=True)
def arm(k):
    def f(c):
        pc = prevq(c)
        if pc is None or c in ANN: return '=""'
        a_, b_ = f'{CL(c)}{R["m_" + k]}', f'{CL(pc)}{R["m_" + k]}'
        return f'=IF(COUNT({a_},{b_})<2,"",IFERROR({CL(c)}{R["r_" + k]}/AVERAGE({a_},{b_})/3,""))'
    return f
rr += 1
hdr_row(rr, 'Receita média por assinante - ARM (USD por mês)'); rr += 1
for k, nm, mm in REG: line('a_' + k, '   ' + nm, '#,##0.00', 'blank', arm(k))
"""

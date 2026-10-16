# gerador de spec padrão: segmentos (receita e lucro) + lista livre de linhas; usado por várias companhias
def make(DL, TITLE, NOTE, SEGS, EXTRA=(), RV=('Revenues', 'RevenueFromContractWithCustomerExcludingAssessedTax', 'RevenueFromContractWithCustomerIncludingAssessedTax'),
         OI=('OperatingIncomeLoss',), SUMCHK=True, MONTHMAP=None, ORIG=False, DROP=()):
    g = {'DL': DL, 'TITLE': TITLE, 'NOTE': NOTE, 'MONTHMAP': MONTHMAP or {}, 'C17': 'Check Receita (linha 15 - soma dos segmentos)' if SUMCHK else 'Receita não alocada aos segmentos (linha 15 - soma)'}
    def items(co):
        def q(tags, mems):
            out = {}
            for t in tags:
                for k, v in co.dq(t, mems, orig=ORIG).items(): out.setdefault(k, v)
            return out
        out = []
        for k, nm, mems, has_oi in SEGS:
            out.append(('s_' + k, f'{nm} - receita (USD)', 'moeda', q(RV, mems), '#,##0', 'sum'))
            if has_oi: out.append(('o_' + k, f'{nm} - lucro do segmento (USD)', 'moeda', q(OI if has_oi is True else has_oi, mems), '#,##0', 'sum'))
        for k, nm, tags, mems in EXTRA:
            out.append(('x_' + k, f'{nm} (USD)', 'moeda', q(tags, mems), '#,##0', 'sum'))
        for o in out:                               # pontos descartados por inconsistência da fonte (ver nota)
            for p in DROP: o[3].pop(p, None)
        return out
    g['items'] = items
    g['SEGS'], g['EXTRA'] = SEGS, EXTRA
    g['BLOCK'] = r'''
if SEGS:
    hdr_row(rr, 'Segmentos - receita (USD milhões)'); rr += 1
    for k, nm, mems, has_oi in SEGS: line('s_' + k, '   ' + nm)
    line('tot', 'Receita (soma dos segmentos)', annual='calc', fn=sumf(*['s_' + k for k, nm, mems, has_oi in SEGS]), b=True)
    line('chk', C17, annual='calc', fn=lambda c: f'=IF(COUNT({CL(c)}{R["tot"]})=0,"",ROUND({CL(c)}15-{CL(c)}{R["tot"]},1))')
    for k, nm, mems, has_oi in SEGS[:3]: line('y_' + k, '   % YoY ' + nm, '0.0%', 'calc', yoy(R['s_' + k]))
else:
    line('chk', '(sem abertura de receita por segmento)', annual='calc', fn=lambda c: '=""')
if any(h for *_, h in SEGS):
    rr += 1
    hdr_row(rr, 'Segmentos - lucro e margem (USD milhões)'); rr += 1
    for k, nm, mems, has_oi in SEGS:
        if has_oi: line('o_' + k, '   ' + nm)
    for k, nm, mems, has_oi in SEGS:
        if has_oi: line('m_' + k, '   Margem ' + nm, '0.0%', 'calc', ratio(R['o_' + k], R['s_' + k]))
if EXTRA:
    rr += 1
    hdr_row(rr, 'Outras aberturas (USD milhões)'); rr += 1
    for k, nm, tags, mems in EXTRA: line('x_' + k, '   ' + nm)
'''
    return g

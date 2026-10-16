import sys; sys.path.insert(0, __import__('os').environ.get('PIPE_SEC', 'pipeline/sec')); from segspec import make; from sx import kind_of
IE, UN = 'us-gaap:IntersegmentEliminationMember', 'fcx:UnaffiliatedCustomersMember'
SG = [('mor', 'Morenci', 'fcx:MorenciMember', ('srt:NorthAmericaMember',)), ('cv', 'Cerro Verde', 'fcx:CerroVerdeMember', ('srt:SouthAmericaMember',)),
      ('gras', 'Grasberg (Indonésia)', 'fcx:GrasbergSegmentMember', ('country:ID',)), ('rod', 'Rod & Refining', 'fcx:RodAndRefiningSegmentMember', ()),
      ('atl', 'Atlantic Copper', 'fcx:AtlanticCopperSmeltingAndRefiningSegmentMember', ())]
globals().update(make(__import__('os').environ.get('PIPE_DATA', 'dados')+'/fcx/dl',
 'DRIVERS DE RECEITA (codigos 45+): segmentos Morenci, Cerro Verde, Grasberg (Indonésia), Rod & Refining e Atlantic Copper - receita total do segmento (terceiros + intersegmento) e lucro operacional - SEC XBRL (USD, fator 0,000001).',
 'Notas: valores dimensionais das instâncias XBRL. Até 3T25 a Freeport etiquetava a receita de cada mina separando vendas a terceiros (membro do segmento) e vendas intersegmento (membro de eliminação); a receita do segmento aqui é a soma das duas. A partir de 4T25 usa o total do segmento. As minas norte-americanas vendem quase tudo à Rod & Refining, por isso a soma dos segmentos supera a receita consolidada (diferença = eliminações e demais minas/molibdênio).',
 [(k, nm, [tuple(sorted((m,) + ex)), m], True) for k, nm, m, ex in SG], SUMCHK=False))
_items = items
def items(co):
    out = _items(co)
    for k, nm, m, ex in SG:
        new = co.dq('Revenues', [(m, UN)])                       # nova etiquetagem (4T25+): existe o total do segmento
        tot = co.dq('Revenues', [m]); old_u = co.dq('Revenues', [tuple(sorted((m,) + ex))] if ex else [m]); old_i = co.dq('Revenues', [tuple(sorted((m, IE) + ex))])
        ser = {}
        for p in set(tot) | set(old_u):
            if p in new and p in tot: ser[p] = tot[p]
            elif p in old_u: ser[p] = old_u[p] + old_i.get(p, 0)
        for y in range(21, 27):                                  # 4T na virada de etiquetagem: anual (total novo) - 9M do próprio trimestre a trimestre
            q4 = f'4Q{y}'
            if q4 not in ser and all(f'{i}Q{y}' in ser for i in (1, 2, 3)):
                fy = [v for (n, ns, s0, e, dd, v, u, a, fl, i) in co.facts if n == 'Revenues' and s0 and e.startswith(f'20{y}-12') and kind_of(s0, e) == '12M'
                      and tuple(sorted(mm for ax, mm in dd if mm != 'us-gaap:OperatingSegmentsMember')) == (m,)]
                if fy: ser[q4] = fy[-1] / 1e6 * 1e6 - sum(ser[f'{i}Q{y}'] for i in (1, 2, 3))
        for o in out:
            if o[0] == 's_' + k: o[3].clear(); o[3].update(ser)
    return out

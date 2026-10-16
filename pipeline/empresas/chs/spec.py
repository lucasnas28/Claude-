import sys; sys.path.insert(0, __import__('os').environ.get('PIPE_SEC', 'pipeline/sec')); from segspec import make
PT = ('IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest', 'IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments')
globals().update(make(__import__('os').environ.get('PIPE_DATA', 'dados')+'/chs/dl',
 'DRIVERS DE RECEITA (codigos 45+): segmentos Energy, Ag (até o ano fiscal 2025), Grains e Agronomy (desde o ano fiscal 2026), Nitrogen Production e Corporate & Other — receita e lucro antes do IR por segmento - SEC XBRL (USD, fator 0,000001).',
 'Notas: valores dimensionais das instâncias XBRL. CHS é uma cooperativa agrícola; o ano fiscal termina em 31/ago (nov -> 4T, fev -> 1T, mai -> 2T, ago -> 3T do calendário). A CHS mede os segmentos pelo lucro antes do IR (inclui equivalência, como a da CF Nitrogen no segmento Nitrogen Production, que não tem receita). A partir do ano fiscal 2026 (4T25 do calendário) Ag foi dividido em Grains e Agronomy e Nitrogen Production passou para "demais segmentos". A diferença para a receita do modelo são eliminações e itens de reconciliação.',
 [('en', 'Energy', ['chscp:EnergyMember'], PT), ('ag', 'Ag (estrutura até FY25)', ['chscp:AgMember'], PT),
  ('gr', 'Grains (estrutura FY26)', ['chscp:GrainsMember'], PT), ('agro', 'Agronomy (estrutura FY26)', ['chscp:AgronomyMember'], PT),
  ('nit', 'Nitrogen Production (até FY25; sem receita, só equivalência)', ['chscp:NitrogenProductionMember'], PT),
  ('corp', 'Corporate & Other / demais segmentos', ['us-gaap:CorporateAndOtherMember', 'us-gaap:AllOtherSegmentsMember'], PT)],
 RV=('Revenues', 'RevenueFromContractWithCustomerExcludingAssessedTax'), SUMCHK=False, ORIG=True))
_items = items
def items(co):
    """estrutura antiga até 3T25 (ago/25) e nova a partir de 4T25; os comparativos reapresentados da nova estrutura no FY25 ficam de fora para não somar em dobro"""
    out = _items(co)
    key = lambda p: (int(p[2:]), int(p[0]))
    for o in out:
        k = o[0][2:]
        for p in list(o[3]):
            if k in ('gr', 'agro') and key(p) < (25, 4): o[3].pop(p)
            if k in ('ag', 'nit') and key(p) >= (25, 4): o[3].pop(p)
    return out
_items2 = items
def items(co):
    """Corporate & Other até 3T25; a partir de 4T25 'demais segmentos' (AllOtherSegments) — evita misturar o recast do FY25"""
    out = _items2(co)
    key = lambda p: (int(p[2:]), int(p[0]))
    for o in out:
        if o[0] in ('s_corp', 'o_corp'):
            tags = ('Revenues', 'RevenueFromContractWithCustomerExcludingAssessedTax') if o[0] == 's_corp' else PT
            old, new = {}, {}
            for t in tags:
                for k, v in co.dq(t, ['us-gaap:CorporateAndOtherMember'], orig=True).items(): old.setdefault(k, v)
                for k, v in co.dq(t, ['us-gaap:AllOtherSegmentsMember']).items(): new.setdefault(k, v)
            o[3].clear(); o[3].update({p: v for p, v in old.items() if key(p) < (25, 4)}); o[3].update({p: v for p, v in new.items() if key(p) >= (25, 4)})
    return out

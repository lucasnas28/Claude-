import sys; sys.path.insert(0, __import__('os').environ.get('PIPE_SEC', 'pipeline/sec')); from segspec import make
RV = ('Revenues', 'RevenueFromContractWithCustomerExcludingAssessedTax')
globals().update(make(__import__('os').environ.get('PIPE_DATA', 'dados')+'/bg/dl',
 'DRIVERS DE RECEITA (codigos 45+): receita a terceiros por segmento. Estrutura antiga (até 2T25, somada e conferida): Agribusiness, Refined & Specialty Oils, Milling, Sugar & Bioenergy, Corporate (Fertilizer incluído em Agribusiness, como na reapresentação da companhia); estrutura nova (pós-Viterra, com comparativos reapresentados desde 3T24): Soybean P&R, Softseed P&R, Tropical Oils & Specialty Ingredients, Grain Merchandising & Milling - SEC XBRL (USD, fator 0,000001).',
 'Notas: valores dimensionais das instâncias XBRL (receita de clientes externos; vendas intersegmento excluídas). Bunge Limited (CIK 1144519) até 3T23 e Bunge Global SA (CIK 1996862) depois. No 10-K a Bunge etiqueta Refined & Specialty Oils como "Edible Oil Products". A aquisição da Viterra (jul/2025) quebra a comparabilidade a partir do 3T25. Lucro bruto por segmento só nos 10-Q da estrutura antiga (sem 4T). A linha de diferença compara a receita do modelo com a soma da estrutura antiga; a partir do 3T25 só existe a nova.',
 [('agri', 'Agribusiness', ['bg:AgribusinessMember'], ['GrossProfit']), ('rso', 'Refined & Specialty Oils', ['bg:RefinedAndSpecialtyOilsMember', 'bg:EdibleOilProductsMember'], ['GrossProfit']),
  ('mill', 'Milling', ['bg:MillingProductsMember'], ['GrossProfit']), ('sugar', 'Sugar & Bioenergy', ['bg:SugarAndBioenergyMember'], ['GrossProfit']),
  ('corp', 'Corporate & Other', ['bg:CorporateAndReconcilingItemsMember', 'us-gaap:CorporateNonSegmentMember'], False)],
 EXTRA=[('soy', 'Nova estrutura - Soybean Processing & Refining', RV, ['bg:SoybeanProcessingAndRefiningMember']), ('soft', 'Nova estrutura - Softseed Processing & Refining', RV, ['bg:SoftseedProcessingAndRefiningMember']),
        ('trop', 'Nova estrutura - Tropical Oils & Specialty Ingredients', RV, ['bg:TropicalOilsAndSpecialtyIngredientsMember']), ('grain', 'Nova estrutura - Grain Merchandising & Milling', RV, ['bg:GrainMerchandisingAndMillingMember']),
        ('corpn', 'Nova estrutura - Corporate & Other', RV, ['us-gaap:CorporateNonSegmentMember', 'bg:CorporateAndReconcilingItemsMember'])],
 RV=RV))
_items = items
def items(co):
    out = _items(co)
    fert = co.dq('Revenues', ['bg:FertilizerMember'])          # FY20 do 10-K ainda separava Fertilizer; 9M20 já reapresentado dentro de Agribusiness
    for o in out:
        if o[0] == 's_agri' and '4Q20' in o[3]: o[3]['4Q20'] += sum(fert.get(f'{i}Q20', 0) for i in (1, 2, 3, 4))
    for o in out:                                   # estrutura antiga termina no 2T25 (depois só a nova)
        if o[0][2:] in ('agri', 'rso', 'mill', 'sugar', 'corp'):
            for p in ('3Q25', '4Q25', '1Q26', '2Q26'): o[3].pop(p, None)
        if o[0] in ('x_soy', 'x_soft', 'x_trop', 'x_grain', 'x_corpn'):
            for p in list(o[3]):
                if (int(p[2:]), int(p[0])) < (24, 3): o[3].pop(p)
    return out

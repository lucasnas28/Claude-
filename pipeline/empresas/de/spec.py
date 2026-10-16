import sys; sys.path.insert(0, __import__('os').environ.get('PIPE_SEC', 'pipeline/sec')); from segspec import make
globals().update(make(__import__('os').environ.get('PIPE_DATA', 'dados')+'/de/dl',
 'DRIVERS DE RECEITA (codigos 45+): segmentos Production & Precision Ag, Small Ag & Turf, Construction & Forestry e Financial Services (receita e lucro operacional) - SEC XBRL (USD, fator 0,000001).',
 'Notas: ano fiscal da Deere termina no fim de outubro/início de novembro (fechamento de jan -> 4T anterior, abr -> 1T, jul -> 2T, out -> 3T). Segmentos agrícolas atuais desde 2020. A diferença para a linha 15 são outras receitas e eliminações.',
 [('ppa', 'Production & Precision Agriculture', ['de:ProductionAndPrecisionAgricultureSegmentMember'], True),
  ('sat', 'Small Agriculture & Turf', ['de:SmallAgricultureAndTurfSegmentMember'], True),
  ('cf', 'Construction & Forestry', ['de:ConstructionAndForestrySegmentMember'], True),
  ('fs', 'Financial Services', ['de:FinancialServicesSegmentMember'], True)], SUMCHK=False, RV=('Revenues', 'RevenueFromContractWithCustomerExcludingAssessedTax', 'RevenueNotFromContractWithCustomer')))

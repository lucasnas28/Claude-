import sys; sys.path.insert(0, __import__('os').environ.get('PIPE_SEC', 'pipeline/sec')); from segspec import make
globals().update(make(__import__('os').environ.get('PIPE_DATA', 'dados')+'/unh/dl',
 'DRIVERS DE RECEITA (codigos 45+): UnitedHealthcare, Optum Health, Optum Insight, Optum Rx (receita total do segmento e lucro operacional), total Optum e produtos x serviços - SEC XBRL (USD, fator 0,000001).',
 'Notas: valores dimensionais das instâncias XBRL. A receita de cada segmento inclui vendas intersegmento (eliminadas no consolidado), por isso a soma supera a receita do modelo; a linha de diferença mostra as eliminações.',
 [('uhc', 'UnitedHealthcare', ['unh:UnitedhealthcareMember'], True), ('oh', 'Optum Health', ['unh:OptumHealthMember', 'unh:OptumhealthMember'], True),
  ('oi', 'Optum Insight', ['unh:OptumInsightMember', 'unh:OptuminsightMember'], True), ('orx', 'Optum Rx', ['unh:OptumRxMember', 'unh:OptumrxMember'], True)],
 EXTRA=[('opt', 'Optum total - receita', ['Revenues'], ['unh:TotalOptumMember']), ('opt_oi', 'Optum total - lucro operacional', ['OperatingIncomeLoss'], ['unh:TotalOptumMember']),
        ('prd', 'Receita de produtos', ['RevenueFromContractWithCustomerExcludingAssessedTax'], ['us-gaap:ProductMember']), ('srv', 'Receita de serviços (contratos com clientes)', ['RevenueFromContractWithCustomerExcludingAssessedTax'], ['us-gaap:ServiceMember'])], SUMCHK=False))

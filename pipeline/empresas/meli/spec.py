import sys; sys.path.insert(0, __import__('os').environ.get('PIPE_SEC', 'pipeline/sec')); from segspec import make
globals().update(make(__import__('os').environ.get('PIPE_DATA', 'dados')+'/meli/dl',
 'DRIVERS DE RECEITA (codigos 45+): receita por país (Brasil, Argentina, México, outros) e por unidade (Commerce x Fintech; crédito dentro de Fintech) - SEC XBRL (USD, fator 0,000001).',
 'Notas: valores dimensionais das instâncias XBRL (receita líquida). A companhia trocou a etiqueta de receita em 2023; as duas foram unidas.',
 [('br', 'Brasil', ['meli:BrazilSegmentMember'], False), ('ar', 'Argentina', ['meli:ArgentinaSegmentMember'], False), ('mx', 'México', ['meli:MexicoSegmentMember'], False), ('oth', 'Outros países', ['meli:OtherCountriesSegmentMember'], False)],
 EXTRA=[('com', 'Commerce - receita', ['Revenues', 'RevenueFromContractWithCustomerExcludingAssessedTax'], ['meli:CommerceMember']), ('fin', 'Fintech - receita', ['Revenues', 'RevenueFromContractWithCustomerExcludingAssessedTax'], ['meli:FintechMember']),
        ('cred', 'Fintech - receita de crédito', ['Revenues'], [('meli:CreditRevenuesMember', 'meli:FintechMember')]), ('prd', 'Vendas de produtos (1P)', ['Revenues', 'RevenueFromContractWithCustomerExcludingAssessedTax'], ['us-gaap:ProductMember'])]))

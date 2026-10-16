import sys; sys.path.insert(0, __import__('os').environ.get('PIPE_SEC', 'pipeline/sec')); from segspec import make
globals().update(make(__import__('os').environ.get('PIPE_DATA', 'dados')+'/wmt/dl',
 'DRIVERS DE RECEITA (codigos 45+): segmentos Walmart U.S., Walmart International e Sam\'s Club U.S. (receita e lucro operacional), categorias do Walmart U.S. e e-commerce - SEC XBRL (USD, fator 0,000001).',
 'Notas: valores dimensionais das instâncias XBRL. Ano fiscal termina em 31/jan: o 4T fiscal entra como 4T do calendário anterior (ex.: jan/26 = 4Q25). A receita de segmentos é a receita líquida de vendas; a receita total do modelo inclui também a receita de membership, por isso a linha de diferença mostra membership e outras receitas.',
 [('us', 'Walmart U.S.', ['wmt:WalmartUSSegmentMember', 'wmt:WalmartUSMember'], True), ('intl', 'Walmart International', ['wmt:WalmartInternationalSegmentMember', 'wmt:WalmartInternationalMember'], True),
  ('sams', "Sam's Club U.S.", ['wmt:SamsClubSegmentMember', 'wmt:SamsClubMember', 'wmt:SamsClubUSMember', 'wmt:SamsClubUSSegmentMember'], True)],
 EXTRA=[('gro', 'Walmart U.S. - Grocery', ['RevenueFromContractWithCustomerExcludingAssessedTax'], [('wmt:GroceryMember', 'wmt:WalmartUSSegmentMember'), ('wmt:GroceryMember', 'wmt:WalmartUSMember')]),
        ('gm', 'Walmart U.S. - General merchandise', ['RevenueFromContractWithCustomerExcludingAssessedTax'], [('wmt:GeneralMerchandiseMember', 'wmt:WalmartUSSegmentMember'), ('wmt:GeneralMerchandiseMember', 'wmt:WalmartUSMember')]),
        ('hw', 'Walmart U.S. - Health and wellness', ['RevenueFromContractWithCustomerExcludingAssessedTax'], [('wmt:HealthandWellnessMember', 'wmt:WalmartUSSegmentMember'), ('wmt:HealthandWellnessMember', 'wmt:WalmartUSMember')]),
        ('ecus', 'Walmart U.S. - eCommerce', ['RevenueFromContractWithCustomerExcludingAssessedTax'], [('wmt:ECommerceMember', 'wmt:WalmartUSSegmentMember'), ('wmt:ECommerceMember', 'wmt:WalmartUSMember')]),
        ('ecin', 'Walmart International - eCommerce', ['RevenueFromContractWithCustomerExcludingAssessedTax'], [('wmt:ECommerceMember', 'wmt:WalmartInternationalSegmentMember'), ('wmt:ECommerceMember', 'wmt:WalmartInternationalMember')]),
        ('mex', 'Walmart International - México e América Central', ['RevenueFromContractWithCustomerExcludingAssessedTax'], [('wmt:MexicoandCentralAmericaMember', 'wmt:WalmartInternationalSegmentMember'), ('wmt:MexicoandCentralAmericaMember', 'wmt:WalmartInternationalMember')])],
 RV=('RevenueFromContractWithCustomerExcludingAssessedTax', 'Revenues'), SUMCHK=False))

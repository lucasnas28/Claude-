import sys; sys.path.insert(0, __import__('os').environ.get('PIPE_SEC', 'pipeline/sec')); from segspec import make
globals().update(make(__import__('os').environ.get('PIPE_DATA', 'dados')+'/cvx/dl',
 'DRIVERS (codigos 45+): lucro por segmento Upstream x Downstream (até 3T24) e receita de petróleo e gás EUA x exterior - SEC XBRL (USD, fator 0,000001).',
 'Notas: valores dimensionais das instâncias XBRL. A Chevron não etiqueta a abertura por segmento nos períodos mais recentes do mesmo jeito; ficam em branco onde não há dado. O lucro operacional do modelo é receita total (inclui equivalência e outras receitas) - custos totais + juros.',
 [], EXTRA=[('up', 'Lucro Upstream', ['NetIncomeLoss'], ['cvx:UpstreamMember', 'us-gaap:UpstreamEquipmentMember']), ('down', 'Lucro Downstream', ['NetIncomeLoss'], ['cvx:DownstreamMember', 'us-gaap:DownstreamEquipmentMember']),
            ('og_us', 'Receita de óleo e gás EUA', ['RevenueFromContractWithCustomerExcludingAssessedTax', 'RevenueFromContractWithCustomerIncludingAssessedTax'], [('country:US', 'us-gaap:OilAndGasMember')]),
            ('og_xu', 'Receita de óleo e gás fora dos EUA', ['RevenueFromContractWithCustomerExcludingAssessedTax', 'RevenueFromContractWithCustomerIncludingAssessedTax'], [('us-gaap:NonUsMember', 'us-gaap:OilAndGasMember')])], SUMCHK=False))

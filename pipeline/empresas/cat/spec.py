import sys; sys.path.insert(0, __import__('os').environ.get('PIPE_SEC', 'pipeline/sec')); from segspec import make
RS = 'us-gaap:ReportableSubsegmentsMember'
globals().update(make(__import__('os').environ.get('PIPE_DATA', 'dados')+'/cat/dl',
 'DRIVERS DE RECEITA (codigos 45+): segmentos Construction Industries, Resource Industries, Energy & Transportation (Power & Energy desde 2025) e Financial Products; regiões - SEC XBRL (USD, fator 0,000001).',
 'Notas: valores dimensionais das instâncias XBRL. A soma dos segmentos inclui vendas entre segmentos; a diferença para a linha 15 são eliminações e corporate. O 4T21 dos segmentos fica vazio (10-K de 2021 etiquetado em base diferente dos 9M).',
 [('ci', 'Construction Industries', ['cat:ConstructionIndustriesMember', ('cat:ConstructionIndustriesMember', RS)], False),
  ('ri', 'Resource Industries', ['cat:ResourceIndustriesMember', ('cat:ResourceIndustriesMember', RS)], False),
  ('et', 'Energy & Transportation / Power & Energy', ['cat:EnergyandTransportationMember', 'cat:PowerEnergyMember', ('cat:EnergyandTransportationMember', RS)], False),
  ('fp', 'Financial Products', ['cat:FinancialProductsSegmentMember', ('cat:FinancialProductsSegmentMember', RS)], False)],
 EXTRA=[('na', 'Receita América do Norte', ['Revenues'], ['srt:NorthAmericaMember']), ('la', 'Receita América Latina', ['Revenues'], ['srt:LatinAmericaMember']),
        ('em', 'Receita EAME', ['Revenues'], ['us-gaap:EMEAMember']), ('ap', 'Receita Ásia-Pacífico', ['Revenues'], ['srt:AsiaPacificMember'])], SUMCHK=False, DROP=('4Q21',)))

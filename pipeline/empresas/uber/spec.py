import sys; sys.path.insert(0, __import__('os').environ.get('PIPE_SEC', 'pipeline/sec')); from segspec import make
globals().update(make(__import__('os').environ.get('PIPE_DATA', 'dados')+'/uber/dl',
 'DRIVERS DE RECEITA (codigos 45+): segmentos Mobility, Delivery e Freight; receita por região (EUA e Canadá, América Latina, EMEA, Ásia-Pacífico) - SEC XBRL (USD, fator 0,000001).',
 'Notas: valores dimensionais das instâncias XBRL. A Uber mede os segmentos por EBITDA ajustado, que não é etiquetado no XBRL; por isso só receita. Antes de 2020 havia também ATG e outros programas de tecnologia (receita residual).',
 [('mob', 'Mobility', ['uber:MobilityMember'], False), ('del', 'Delivery', ['uber:DeliveryMember'], False), ('frt', 'Freight', ['uber:FreightMember'], False)],
 EXTRA=[('usc', 'EUA e Canadá', ['Revenues'], ['uber:UnitedStatesAndCanadaMember']), ('lat', 'América Latina', ['Revenues'], ['srt:LatinAmericaMember']),
        ('emea', 'EMEA', ['Revenues'], ['us-gaap:EMEAMember']), ('apac', 'Ásia-Pacífico', ['Revenues'], ['srt:AsiaPacificMember'])]))

NAME = 'CHS Inc.'; TICKER = 'CHSCP'; CIK = '0000823277'
DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/chs/dl'
UNSKIP = ['PreferredStockValue']    # ações preferenciais (CHSCP, CHSCL…) no PL
BSMAP = {'PreferredStockValue': 40}
# Os 10-K de 2022+ trazem a instância fora da ordem do documento: o balanço de agosto (3T) vem da coluna comparativa do 10-Q de novembro.
# Os 10-Q de 2020 (fev e mai) foram arquivados em ordem alfabética, sem como reconstruir a face: balanço vazio nesses dois trimestres.
FACE_OVR = {'3Q22': ('0000823277-22-000006', '2022-08-31') if False else ('0000823277-23-000006', '2022-08-31'),
            '3Q23': ('0000823277-24-000006', '2023-08-31'), '3Q24': ('0000823277-25-000006', '2024-08-31'),
            '3Q25': ('0000823277-26-000006', '2025-08-31'), '1Q20': 'skip', '2Q20': 'skip'}
DA = ['DepreciationDepletionAndAmortizationincludingmajormaintenance', 'DepreciationDepletionAndAmortization']   # 10-K etiqueta a D&A com manutenção maior numa tag própria; nos 10-Q a tag padrão já a inclui

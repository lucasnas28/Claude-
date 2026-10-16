NAME = 'Visa Inc.'; TICKER = 'V'; CIK = '0001403161'
DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/v/dl'
BSMAP = {'RestrictedCashAndCashEquivalentsAtCarryingValue': 16, 'SettlementReceivable': 16, 'SettlementPayable': 26, 'CustomerCollateral': 26, 'AccruedLitigation': 26,
         'InvestmentSecuritiesAvailableForSaleCurrent': 11, 'LongTermInvestments': 22}

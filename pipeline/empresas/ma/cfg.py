NAME = 'Mastercard Incorporated'; TICKER = 'MA'; CIK = '0001141391'
DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/ma/dl'
BSMAP = {'RestrictedCashAndCashEquivalentsAtCarryingValue': 16, 'RestrictedCashAndRestrictedCashEquivalentsLitigationProvision': 16, 'SettlementDueFromCustomers': 16, 'SettlementDueToCustomers': 26,
         'RestrictedSecurityDeposits': 26, 'AccruedLitigationCurrent': 26, 'InvestmentsNoncurrent': 22, 'Investments': 11}

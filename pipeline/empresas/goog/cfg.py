NAME = 'Alphabet Inc.'; TICKER = 'GOOGL'; CIK = '0001652044'
DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/goog/dl'
DA = ['DepreciationAmortizationAndImpairmentOnDispositionOfPropertyAndEquipment', 'DepreciationAndImpairmentOnDispositionOfPropertyAndEquipment', 'Depreciation']
BSMAP = {'NonMarketableSecurities': 22, 'AccruedRevenueShare': 24, 'ContractWithCustomerLiabilityCurrent': 26, 'IncomeTaxesPayableNoncurrent': 34}

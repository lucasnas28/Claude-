NAME = 'Tesla Inc.'; TICKER = 'TSLA'; CIK = '0001318605'
DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/tsla/dl'
BSMAP = {'OperatingLeaseVehicles': 22, 'SolarEnergySystemsNet': 22, 'DigitalAssetsNoncurrent': 22, 'ContractWithCustomerLiabilityCurrent': 26, 'ContractWithCustomerLiabilityNoncurrent': 34,
         'LongTermDebtAndCapitalLeaseObligationsCurrent': 27, 'TemporaryEquityCarryingAmountAttributableToParent': 40}

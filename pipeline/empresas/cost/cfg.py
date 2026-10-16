NAME = 'Costco Wholesale Corporation'; TICKER = 'COST'; CIK = '0000909832'
DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/cost/dl'
MONTHMAP = {11: '4', 12: '4', 2: '1', 3: '1', 5: '2', 6: '2', 8: '3', 9: '3'}
COGS = ['CostOfMerchandiseSold', 'CostOfGoodsAndServicesSold', 'CostOfRevenue']
BSMAP = {'ReceivablesNetCurrent': 12, 'MerchandiseInventory': 13, 'InventoryNet': 13, 'MembershipFeesDeferredRevenueCurrent': 26}
# a partir do ano fiscal 2025 os 10-Q etiquetam o lucro antes do IR como IncomeLossAttributableToParent
PRETAX = ['IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest', 'IncomeLossAttributableToParent']

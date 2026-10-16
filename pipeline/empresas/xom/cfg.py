NAME = 'Exxon Mobil Corporation'; TICKER = 'XOM'; CIK = '0000034088'
DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/xom/dl'
REV = ['Revenues', 'RevenueFromContractWithCustomerIncludingAssessedTax']
BSMAP = {'NotesAndLoansPayable': 27, 'PostemploymentBenefitsLiabilityNoncurrent': 34, 'AssetRetirementObligationsNoncurrent': 34, 'InvestmentsAndAdvances': 17}
REV = ['Revenues']
COGS = ['CrudeOilAndProductPurchases']

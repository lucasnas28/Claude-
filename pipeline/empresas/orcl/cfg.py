NAME = 'Oracle Corporation'; TICKER = 'ORCL'; CIK = '0001341439'
DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/orcl/dl'
COGS_SUM = [['CloudServicesAndLicenseSupportExpenses', 'CloudAndSoftwareExpenses'], ['HardwareExpenses'], ['ServicesExpense']]   # custo = despesas de cloud e software + hardware + serviços
DA_SUM = [['Depreciation'], ['AmortizationOfIntangibleAssets']]
UNSKIP = ['PreferredStockValue']    # preferenciais conversíveis obrigatórias série D (2026) no PL
BSMAP = {'PreferredStockValue': 40}

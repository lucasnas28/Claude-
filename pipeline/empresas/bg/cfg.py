NAME = 'Bunge Global SA (antes Bunge Limited)'; TICKER = 'BG'; CIK = '0001996862'
DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/bg/dl'
UNSKIP = ['PreferredStockValue']    # ações preferenciais conversíveis (US$ 690 mi) no PL até a conversão em 2022
BSMAP = {'PreferredStockValue': 40}
def OP_FUNC(q):
    gp, sga = q(['GrossProfit']), q(['SellingGeneralAndAdministrativeExpense'])
    return {p: gp[p] - sga[p] for p in gp if p in sga}
OP_DESC = 'Lucro operacional DERIVADO = lucro bruto - despesas de vendas, gerais e administrativas [us-gaap:GrossProfit - SellingGeneralAndAdministrativeExpense] (a Bunge não divulga lucro operacional; câmbio, juros, outras e equivalência ficam no não operacional)'
DA_ADDBACK = 2

NAME = 'Advanced Micro Devices Inc.'; TICKER = 'AMD'; CIK = '0000002488'
DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/amd/dl'
def DA_FUNC(qser):
    od, am = qser(['OtherDepreciationAndAmortization', 'DepreciationDepletionAndAmortization']), qser(['AmortizationOfIntangibleAssets'])
    out = {}
    for p, v in od.items():
        a = am.get(p, 0)
        out[p] = v if v >= a else v + a      # até 2024 a linha "depreciação e amortização" já incluía a amortização de intangíveis; depois veio separada
    return out
DA_DESC = 'D&A = depreciação e amortização do fluxo de caixa (+ amortização de intangíveis de aquisição quando veio em linha separada, a partir de 2024)'

NAME = 'GE Aerospace (General Electric Company)'; TICKER = 'GE'; CIK = '0000040545'
DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/ge/dl'
COGS_DIMS = [('CostOfGoodsAndServicesSold', ['us-gaap:ProductMember', 'us-gaap:ServiceMember'])]

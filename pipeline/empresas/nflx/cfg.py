NAME = 'Netflix Inc.'; TICKER = 'NFLX'; CIK = '0001065280'
DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/nflx/dl'
BSMAP = {'ContentAssetsNet': 22, 'FilmCosts': 22, 'ContentLiabilitiesCurrent': 26, 'ContentLiabilitiesNoncurrent': 34, 'DeferredRevenueCurrent': 26}

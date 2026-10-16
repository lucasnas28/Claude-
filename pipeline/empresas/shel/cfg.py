NAME = 'Shell plc'; TICKER = 'SHEL'; CIK = '0001306965'
import os
DATA = os.environ.get('PIPE_DATA', 'dados') + '/shel'
BOOK = DATA + '/q2-2026-quarterly-databook.xlsx'
DL = DATA + '/dl'

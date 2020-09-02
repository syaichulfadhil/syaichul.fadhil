from scipy.stats import ttest_ind
from scipy.stats import ttest_1samp
from statistics import mean
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pandas import pivot_table 
import itertools
from datetime import datetime
import pandas as pd
from pandas import ExcelWriter


dfmp1 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\TOKPEDGSK - JUL20.xlsx', sheet_name='Sum')
dfmp2 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\BUKALAPAKGSK - JUL20.xlsx', sheet_name='Sum')
dfmp3 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\LAZADAGSK - JUL20.xlsx', sheet_name='Sum')
dfmp4 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\SHOPEEGSK - JUL20.xlsx', sheet_name='Sum')
dfmp5 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\blibliGSK - JUL20.xlsx', sheet_name='Sum')
dfmp6 = pd.concat([dfmp1, dfmp2, dfmp3, dfmp4, dfmp5], sort=False)\
       .drop_duplicates(subset=['MP', 'M'])
dfmp6 = dfmp6.fillna(0)
dfmp7 = dfmp6[(dfmp6['M'] == 'May-20') | (dfmp6['M'] == 'Jun-20') | (dfmp6['M'] == 'Apr-20') | (dfmp6['M'] == 'Jul-20')]
dfmp = dfmp6[(dfmp6['M'] == 'Jul-20')]
dfmp['revcont.'] = dfmp['1NETTREV'] / dfmp['1NETTREV'].sum()

pivsum = pivot_table(dfmp7, values=['1NETTREV', '2QTY', '3ORDER', '99BUYER', 'BASKET', '5PV', '99SC', 'PREV', 'Promo', 'Voucher'], index=['M'], aggfunc={'1NETTREV':np.sum, '2QTY':np.sum, '3ORDER':np.sum, '99BUYER':np.sum, 'BASKET':np.mean, '5PV':np.sum, '99SC':np.sum, 'PREV':np.sum, 'Promo':np.sum, 'Voucher':np.sum})
dfsum = pd.DataFrame(pivsum.to_records())
dfsum = dfsum.fillna(0)


dfb1 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\TOKPEDGSK - JUL20.xlsx', sheet_name='Brand')
dfb2 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\BUKALAPAKGSK - JUL20.xlsx', sheet_name='Brand')
dfb3 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\LAZADAGSK - JUL20.xlsx', sheet_name='Brand')
dfb4 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\SHOPEEGSK - JUL20.xlsx', sheet_name='Brand')
dfb5 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\blibliGSK - JUL20.xlsx', sheet_name='Brand')
dfb6 = pd.concat([dfb1, dfb2, dfb3, dfb4, dfb5], sort=False)
dfb6 = dfb6.fillna(0)

dfb7 = dfb6[(dfb6['M'] == 'Jul-20')]
pivb = pivot_table(dfb7, values=['1NETTREV', '2QTY', '3ORDER', '99BUYER', '5PV', 'PREV'], index=['BRAND'], aggfunc={'1NETTREV':np.sum, '2QTY':np.sum, '3ORDER':np.sum, '99BUYER':np.sum, '5PV':np.sum, 'PREV':np.sum})
dfb = pd.DataFrame(pivb.to_records())
dfb['revcont.'] = dfb['1NETTREV'] / dfb['1NETTREV'].sum()
dfb = dfb.fillna(0)


dfp1 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\TOKPEDGSK - JUL20.xlsx', sheet_name='Product')
dfp2 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\BUKALAPAKGSK - JUL20.xlsx', sheet_name='Product')
dfp3 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\LAZADAGSK - JUL20.xlsx', sheet_name='Product')
dfp4 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\SHOPEEGSK - JUL20.xlsx', sheet_name='Product')
dfp5 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\blibliGSK - JUL20.xlsx', sheet_name='Product')
dfp6 = pd.concat([dfp1, dfp2, dfp3, dfp4, dfp5], sort=False)
dfp6 = dfp6.fillna(0)

dfp7 = dfp6[(dfp6['M'] == 'Jul-20')]
pivp = pivot_table(dfp7, values=['1NETTREV', '2QTY', '3ORDER', '99BUYER', '5PV', 'PREV'], index=['PRODUCT'], aggfunc={'1NETTREV':np.sum, '2QTY':np.sum, '3ORDER':np.sum, '99BUYER':np.sum, '5PV':np.sum, 'PREV':np.sum})
dfp = pd.DataFrame(pivp.to_records())
dfp['revcont.'] = dfp['1NETTREV'] / dfp['1NETTREV'].sum()
dfp = dfp.fillna(0)

dfvs1 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\TOKPEDGSK - JUL20.xlsx', sheet_name='Sum')
dfvs2 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\BUKALAPAKGSK - JUL20.xlsx', sheet_name='Sum')
dfvs3 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\LAZADAGSK - JUL20.xlsx', sheet_name='Sum')
dfvs4 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\SHOPEEGSK - JUL20.xlsx', sheet_name='Sum')
dfvs5 = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\blibliGSK - JUL20.xlsx', sheet_name='Sum')
dfvs6 = pd.concat([dfmp1, dfmp2, dfmp3, dfmp4, dfmp5], sort=False)\
       .drop_duplicates(subset=['MP', 'M'])
dfvs6 = dfvs6.fillna(0)
dfvs7 = dfvs6[(dfvs6['M'] == 'May-20') | (dfvs6['M'] == 'Jun-20') | (dfvs6['M'] == 'Apr-20') | (dfvs6['M'] == 'Jul-20')]

pivsumvs = pivot_table(dfmp7, values=['1NETTREV'], index=['MP', 'M'], aggfunc={'1NETTREV':np.sum})
dfsumvs = pd.DataFrame(pivsumvs.to_records())
dfsumvs = dfsumvs.fillna(0)




writer = pd.ExcelWriter(r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\ALLMPGSK - JUL20.xlsx', engine='xlsxwriter')
dfsum.to_excel(writer, sheet_name='Sum', index=False)
dfsumvs.to_excel(writer, sheet_name='vs 2019', index=False)
dfmp.to_excel(writer, sheet_name='MP Cont.', index=False)
dfb.to_excel(writer, sheet_name='Brand Cont.', index=False)
dfp.to_excel(writer, sheet_name='Product', index=False)
writer.save()


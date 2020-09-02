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

df = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\lazadagsk update.xlsx')


#DATAFRAME (Detail Report)
df.columns = [c.replace(' ', '_') for c in df.columns]
df.columns = [c.replace('-', '_') for c in df.columns]
df = df.rename(columns={"Shipping_Fee":"99SC", "Order_Number": "3ORDER", "Order_Item_Id": "2QTY", "Paid_Price": "1NETTREV", "Item_Name": "PRODUCT", "Status": "STAT", "Customer_Name":"99BUYER"})
dfx = pd.DataFrame(df)
dfx = dfx.fillna(0)
print(dfx.dtypes)

# dfx['4PRICE'] = dfx['4PRICE'].str.replace(r'Rp ', '')
# dfx['4PRICE'] = dfx['4PRICE'].str.replace(r'.', '')
# dfx['4PRICE'] = dfx['4PRICE'].astype(int)
# dfx['5PV'] = dfx['5PV'].astype(int)
dfx['2QTY'] = dfx['2QTY'].astype(int)
dfx['STAT'] = dfx['STAT'].str.replace(r'\r', '')
dfx['STAT'] = dfx['STAT'].str.replace(r'\n', '')
dfx = dfx[(dfx['STAT'] == 'Delivered') | (dfx['STAT'] == 'Shipped')]
detail = pd.DataFrame(dfx)

#Brand Name
SE = ['Sensodyne']
PH = ['Physiogel']
PO = ['Polident']
SC = ['Scotts']
AC = ['Acne Aid']
filterBN = \
    [(detail.PRODUCT.str.contains('|'.join(SE))) | (detail.PRODUCT.isin(SE)), \
    (detail.PRODUCT.str.contains('|'.join(PH))) | (detail.PRODUCT.isin(PH)), \
    (detail.PRODUCT.str.contains('|'.join(PO))) | (detail.PRODUCT.isin(PO)), \
    (detail.PRODUCT.str.contains('|'.join(SC))) | (detail.PRODUCT.isin(SC)), \
    (detail.PRODUCT.str.contains('|'.join(AC))) | (detail.PRODUCT.isin(AC))]

choicesBN = ['Sensodyne', 'Physiogel', 'Polident', 'Scotts', 'Acne Aid']
detail['BRAND'] = np.select(filterBN, choicesBN, default='others')


#FILTER
month = detail[(detail['M'] == 'Jul-20')]
P3M = detail[(detail['M'] == 'May-20') | (detail['M'] == 'Jun-20') | (detail['M'] == 'Apr-20')]


#SUM
piv0 = pivot_table(detail, values=['2QTY', '1NETTREV', '3ORDER', '99BUYER'], index=['M'], aggfunc={'2QTY':pd.Series.nunique, '1NETTREV':np.sum, '3ORDER':pd.Series.nunique, '99BUYER':pd.Series.nunique})
df0 = pd.DataFrame(piv0.to_records())
df0 = df0.fillna(0)
df0['BASKET'] = df0['1NETTREV']/df0['3ORDER']

pivsc = pivot_table(detail, values=['99SC'], index=['3ORDER', 'M'], aggfunc={'99SC':np.mean})
pivsc1 = pivot_table(pivsc, values=['99SC'], index=['M'], aggfunc={'99SC':np.sum})
dfsc1 = pd.DataFrame(pivsc1.to_records())
df00 = pd.merge(df0, dfsc1, on='M', how='left')
df00['MP'] = 'Lazada GSK'

#SUM Product
pivpv = pivot_table(detail, values=['5PV', '2QTY', '1NETTREV', '3ORDER', '99BUYER'], index=['PRODUCT', 'BRAND', 'M'], aggfunc={'5PV':np.mean, '2QTY':pd.Series.nunique, '1NETTREV':np.sum, '3ORDER':pd.Series.nunique, '99BUYER':pd.Series.nunique})
dfpv = pd.DataFrame(pivpv.to_records())
df1 = dfpv[(dfpv['M'] == 'Jul-20')]
df1 = df1.fillna(0)
pivsum = pivot_table(dfpv, values=['5PV'], index=['M'], aggfunc={'5PV':np.sum}) 
dfsum1 = pd.DataFrame(pivsum.to_records())
dfsum = pd.merge(df00, dfsum1, on='M', how='left')


#SUM Brand
piv2 = pivot_table(month, values=['5PV', '2QTY', '1NETTREV', '3ORDER', '99BUYER'], index=['PRODUCT', 'BRAND', 'M'], aggfunc={'5PV':np.mean, '2QTY':pd.Series.nunique, '1NETTREV':np.sum, '3ORDER':pd.Series.nunique, '99BUYER':pd.Series.nunique})
df2 = pd.DataFrame(piv2.to_records())
df2 = df2.fillna(0)
piv3 = pivot_table(df2, values=['5PV', '2QTY', '1NETTREV', '3ORDER', '99BUYER'], index=['BRAND', 'M'], aggfunc={'5PV':np.sum, '2QTY':np.sum, '1NETTREV':np.sum, '3ORDER':np.sum, '99BUYER':np.sum})
df3 = pd.DataFrame(piv3.to_records())
df3 = df3.fillna(0)
df3['MP'] = 'Lazada GSK'

writer = pd.ExcelWriter(r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\LAZADAGSK - JUL20.xlsx', engine='xlsxwriter')
dfsum.to_excel(writer, sheet_name='Sum', index=False)
df1.to_excel(writer, sheet_name='Product', index=False)
df3.to_excel(writer, sheet_name='Brand', index=False)
detail.to_excel(writer, sheet_name='raw', index=False)
writer.save()
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

df = pd.read_excel (r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\shopeegsk update.xlsx')


#DATAFRAME (Detail Report)
df.columns = [c.replace(' ', '_') for c in df.columns]
df.columns = [c.replace('-', '_') for c in df.columns]
df = df.rename(columns={"Total_Pembayaran":"GREV", "Diskon_Dari_Penjual":"Promo",  "Voucher_Ditanggung_Penjual":"Voucher", "Ongkos_Kirim_Dibayar_oleh_Pembeli":"99SC", "No_Pesanan": "3ORDER", "Jumlah": "2QTY", "Harga_Setelah_Diskon": "4PRICE", "Nama_Produk": "PRODUCT", "Status_Pesanan": "STAT", "Username_(Pembeli)":"99BUYER"})
dfx = pd.DataFrame(df)
dfx = dfx.fillna(0)
print(dfx.dtypes)

# dfx['4PRICE'] = dfx['4PRICE'].str.replace(r'Rp ', '')
# dfx['4PRICE'] = dfx['4PRICE'].str.replace(r'.', '')
dfx['2QTY'] = dfx['2QTY'].astype(int)
dfx['4PRICE'] = dfx['4PRICE'].astype(int)
dfx['5PV'] = dfx['5PV'].astype(int)
dfx['Promo'] = dfx['Promo'].astype(int)
dfx['Voucher'] = dfx['Voucher'].astype(int)
dfx['99SC'] = dfx['99SC'].astype(int)
dfx = dfx[(dfx['STAT'] == 'Selesai') | (dfx['STAT'] == 'Sedang Dikirim')]
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
detail['PREV'] = detail['4PRICE']*detail['2QTY']

#FILTER
month = detail[(detail['M'] == 'Jul-20')]
P3M = detail[(detail['M'] == 'May-20') | (detail['M'] == 'Jun-20') | (detail['M'] == 'Apr-20')]


#SUM
piv0 = pivot_table(detail, values=['2QTY', '3ORDER', '99BUYER', '4PRICE', 'Promo', 'PREV'], index=['M'], aggfunc={'2QTY':np.sum, '3ORDER':pd.Series.nunique, '99BUYER':pd.Series.nunique, '4PRICE':np.mean, 'Promo':np.sum, 'PREV':np.sum})
df0 = pd.DataFrame(piv0.to_records())
df0 = df0.fillna(0)


pivsc = pivot_table(detail, values=['99SC', 'Voucher'], index=['3ORDER', 'M'], aggfunc={'99SC':np.mean, 'Voucher':np.mean})
pivsc1 = pivot_table(pivsc, values=['99SC', 'Voucher'], index=['M'], aggfunc={'99SC':np.sum, 'Voucher':np.sum})
dfsc1 = pd.DataFrame(pivsc1.to_records())
df00 = pd.merge(df0, dfsc1, on='M', how='left')
df00['MP'] = 'Shopee GSK'
df00['1NETTREV'] = df00['PREV'] - df00['Voucher']
df00['BASKET'] = df00['1NETTREV']/df00['3ORDER']

#SUM Product
pivpv = pivot_table(detail, values=['5PV', '2QTY', 'PREV', '3ORDER', '99BUYER', '4PRICE'], index=['PRODUCT', 'BRAND', 'M'], aggfunc={'5PV':np.mean, '2QTY':np.sum, 'PREV':np.sum, '3ORDER':pd.Series.nunique, '99BUYER':pd.Series.nunique, '4PRICE':np.mean}) 
dfpv = pd.DataFrame(pivpv.to_records())
df1 = dfpv[(dfpv['M'] == 'Jul-20')]
df1 = df1.fillna(0)
pivsum = pivot_table(dfpv, values=['5PV'], index=['M'], aggfunc={'5PV':np.sum}) 
dfsum1 = pd.DataFrame(pivsum.to_records())
dfsum = pd.merge(df00, dfsum1, on='M', how='left')

#SUM Brand
piv2 = pivot_table(month, values=['5PV', '2QTY', 'PREV', '3ORDER', '99BUYER', '4PRICE'], index=['PRODUCT', 'BRAND', 'M'], aggfunc={'5PV':np.mean, '2QTY':np.sum, 'PREV':np.sum, '3ORDER':pd.Series.nunique, '99BUYER':pd.Series.nunique, '4PRICE':np.mean})
df2 = pd.DataFrame(piv2.to_records())
df2 = df2.fillna(0)
piv3 = pivot_table(df2, values=['5PV', '2QTY', 'PREV', '3ORDER', '99BUYER', '4PRICE'], index=['BRAND', 'M'], aggfunc={'5PV':np.sum, '2QTY':np.sum, 'PREV':np.sum, '3ORDER':np.sum, '99BUYER':np.sum, '4PRICE':np.mean})
df3 = pd.DataFrame(piv3.to_records())
df3 = df3.fillna(0)
df3['MP'] = 'Shopee GSK'

# #SUM PROMO
# pivpromo = pivot_table(month, values=['2QTY', '3ORDER', '4BUYER', 'Promo', 'Voucher', 'PREV'], index=['M', 'PNAME'], aggfunc={'2QTY':np.sum,  '3ORDER':pd.Series.nunique, '4BUYER':pd.Series.nunique, 'Promo':np.sum, 'Voucher':np.sum, 'PREV':np.sum})
# dfpromo = pd.DataFrame(pivpromo.to_records())
# dfpromo = dfpromo.fillna(0)


# #SUM
# piv4 = pivot_table(month, values=['2QTY', '1REV', '3ORDER', '99BUYER', '4PRICE'], index=['W', 'PROMOSTAT'], aggfunc={'2QTY':np.sum, '1REV':np.sum, '3ORDER':pd.Series.nunique, '99BUYER':pd.Series.nunique, '4PRICE':np.mean})
# df4 = pd.DataFrame(piv4.to_records())
# df4 = df4.fillna(0)
# df4['MP'] = 'Tokopedia Martha Tilaar'

writer = pd.ExcelWriter(r'C:\Users\Syaichul Fadhil\Desktop\Research Plan\bhisma\GSK\SHOPEEGSK - JUL20.xlsx', engine='xlsxwriter')
dfsum.to_excel(writer, sheet_name='Sum', index=False)
df1.to_excel(writer, sheet_name='Product', index=False)
df3.to_excel(writer, sheet_name='Brand', index=False)
# df4.to_excel(writer, sheet_name='Weekly', index=False)
detail.to_excel(writer, sheet_name='raw', index=False)
writer.save()
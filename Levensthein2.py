from fuzzywuzzy import fuzz
from fuzzywuzzy import process


import pandas as pd
from pandas import DataFrame
path = 'C:/Users/v-greg.vavoules/Documents/Site Selection/'
filename = 'SF accounts.xlsx'

xl = pd.ExcelFile(path + filename)

dfSF = xl.parse("SF2")
dfSFAddy = dfSF[dfSF.columns[0]]
#print(dfSFAddy)

dfShopkick = xl.parse("SK")
dfShopkickAddy = dfShopkick[dfShopkick.columns[0]]
#print(dfWalmartAddy)

#guess2 = process.extractOne('5500 Grossmont, San Diego, CA 91942', dfAltAddy,scorer=fuzz.token_sort_ratio)
#print(guess2)

print('Shopkick Start')
list5 = []
list6 = []

for item in dfShopkickAddy:
    guess3 = process.extractOne(item, dfSFAddy, scorer=fuzz.token_sort_ratio)
    list5.append(guess3[0])
    list6.append(guess3[1])
    print(item)

print('Shopkick Data Frame')
dfexport2 = DataFrame({
    'Test Address': dfShopkickAddy, 'SalesForce Address': list5, 'Accuracy': list6
})

print('Export Shopkick File')
exportfilename2 = 'Address matching - Shopkick.xlsx'
dfexport2.to_excel(path + exportfilename2, sheet_name='Sheet1', index=False)

print('Done')

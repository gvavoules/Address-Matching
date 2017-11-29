from fuzzywuzzy import fuzz
from fuzzywuzzy import process


import pandas as pd
from pandas import DataFrame
path = 'C:/Users/v-greg.vavoules/Documents/'
filename = 'SF accounts.xlsx'

xl = pd.ExcelFile(path + filename)

dfSF = xl.parse("SF")
dfSFAddy = dfSF[dfSF.columns[1]]
#print(dfSFAddy)

dfWalmart = xl.parse("Walmart")
dfWalmartAddy = dfWalmart[dfWalmart.columns[1]]
#print(dfWalmartAddy)

dfAlt = xl.parse("Alt")
dfAltAddy = dfAlt[dfAlt.columns[0]]
#print(dfAltAddy)

guess2 = process.extractOne('5500 Grossmont, San Diego, CA 91942', dfAltAddy,scorer=fuzz.token_sort_ratio)
#print(guess2)

print('Start Alteryx')
list3 = []
list4 = []

for item in dfAltAddy:
    guess = process.extractOne(item, dfSFAddy, scorer=fuzz.token_sort_ratio)
    list3.append(guess[0])
    list4.append(guess[1])
    print(item)


print('Alt Data Frame')
dfexport = DataFrame({
    'Test Address': dfAltAddy, 'SalesForce Address': list3, 'Accuracy': list4
})

print('Walmart Start')
list5 = []
list6 = []

for item in dfWalmartAddy:
    guess3 = process.extractOne(item, dfSFAddy, scorer=fuzz.token_sort_ratio)
    list5.append(guess3[0])
    list6.append(guess3[1])
    print(item)

print('Walmart Data Frame')
dfexport2 = DataFrame({
    'Test Address': dfWalmartAddy, 'SalesForce Address': list5, 'Accuracy': list6
})

print('Export Alt File')
#print(dfexport)

exportfilename = 'Address matching - Alteryx.xlsx'
dfexport.to_excel(path + exportfilename, sheet_name='Sheet1', index=False)

print('Export Walmart File')
exportfilename2 = 'Address matching - Walmart.xlsx'
dfexport2.to_excel(path + exportfilename2, sheet_name='Sheet1', index=False)

print('Done')

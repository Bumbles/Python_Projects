# -*- coding: utf-8 -*-
"""
Created on Mon Jul 16 17:10:14 2018

@author: OPHIR - Josh Clark
"""

import os
import pandas as pd
import xlwt
import xlrd
import re
import datetime

# 0. Set variables
#Set size of basket
sstkcount = 50
#Set borrow cost threshold
bct = 0.05
#Set total short weight
sw = -0.50
#Define Fundamental Short stocks
fs = ['KGN']
#Set fundamental short basket weight
fsw = -0.01865
#Set basket short weight
bsw = sw - fsw

# 1. Get the weighted average market cap of the portfolio

os.chdir('S:\Ophir\Investment\Quant\Python\Portfolio Data\\')
# List all files and directories in current directory
filelist = [filename for filename in os.listdir('.') if filename.startswith("VAL-FTS_Ophir_")]
date_pattern = re.compile(r'(?P<date>\d{8})')

def get_date(filename):
    matched = date_pattern.search(filename)
    if not matched:
        return None
    return datetime.datetime.strptime(matched.groups('date')[0], "%Y%m%d")

dates = (get_date(fn) for fn in filelist)
dates = (d for d in dates if d is not None)
last_date = max(dates)
file = 'VAL-FTS_Ophir_'+last_date.strftime('%Y%m%d')+'.xls'

# Load spreadsheet
xl = pd.ExcelFile('S:\Ophir\Investment\Quant\Python\Portfolio Data\\' + file)
# Load a sheet into a DataFrame by name: df
df = xl.parse('Sheet1')
#Get Date
date=df.at[0,'As At Date']
#Drop rows with Apps accountin them
df = df[df['Asset Name'] != 'OAMARFAPPS']
dfcash = df[df[' Analysis Group 1'] == 'Cash']
df = df[df[' Analysis Group 1'] != 'Cash']
summary=dfcash.groupby('Portfolio').sum()
summary['Average Cost'] = 1
summary['Market Value'] = 1
summary['Asset'] = 'Cash'
summary['Asset Name'] = 'Cash'
summary['Portfolio'] = 'Cash'
summary['Portfolio']=summary.index
date=df.at[0,'As At Date']
summary['As At Date'] = date
dfi=df.append(summary)
df=dfi[dfi['Portfolio'] == 'OAMARF']
df['Port Weight'] = df['Market Value.1']/df['Market Value.1'].sum()
#Set portfolio FUM
OAMARFFUM = df['Market Value.1'].sum()
#Portfolio FUM minus Cash value = equities value
OAMARFFUM_EQ = OAMARFFUM - df.loc[df['Asset'] == 'Cash']['Market Value.1']
df['eqweight']=OAMARFFUM_EQ[0]
df['eqweight']=df['Market Value.1']/OAMARFFUM_EQ[0]
file = 'Bloomberg Consensus Database 3.0 (Values Only).xlsm'
# Load spreadsheet
# Open the file
xl = pd.ExcelFile('S:\Ophir\Investment\Quant\Python\External Data\\' + file)
# Get the first sheet as an object
dfbb = xl.parse(sheetname='dataimport',skiprows=3)
dfbb = dfbb.drop(dfbb.index[[0,1]])
dfbb = dfbb.loc[~dfbb.index.duplicated(keep='first')]
#df['Mcap'] = df['Asset'].map(dfbb.set_index('Unnamed: 0')['Mcap'])
file = 'mcaps.xlsx'
xl = pd.ExcelFile('S:\Ophir\Investment\Quant\Python\External Data\\' + file)
# Load a sheet into a DataFrame by name: df
dfmcaps = xl.parse('Sheet1')
df['Mcap'] = df['Asset'].map(dfmcaps.set_index('Code')['Mcap'])


wgtavmcap=df['Mcap']*df['eqweight']
wgtavmcap=wgtavmcap.sum()


# 2. import small ords

indexfile = 'Short Hedge.xlsm'
xl = pd.ExcelFile('S:\Ophir\Investment\Quant\Python\External Data\\' + indexfile)
df3 = xl.parse('XSO')
#Remove long holdings from universe
df3 = df3[-df3['Security'].isin(dfi['Asset'])]
#Remove super threshold borrow cost stocks from universe

##Import Borrow Cost
indexfile = 'Stock Borrow.xlsx'
xl = pd.ExcelFile('S:\Ophir\Investment\Quant\Python\External Data\\' + indexfile)
df5 = xl.parse('Sheet1')
df5['Security']=df5.Security.str.replace('.AX','')
df3['Indic Fee'] = df3['Security'].map(df5.set_index('Security')['Indic Fee'])
df3 = df3[df3['Indic Fee']<bct]
df3 = df3[['Security','Date','MarketCapitalisationEOD','ClosePrice','Indic Fee']]
#create new column
df3['wgtdmcap']= 'NaN'
# add nothing for first 50, then weighted av market cap thereafter
i = df3['wgtdmcap'].count()
df3 = df3.sort_values(['MarketCapitalisationEOD'], ascending=[0])
df3=df3.reset_index(drop=True)
m1=[]

for i in range(0,i-sstkcount):
    #Test - find weighted mcap of previous 50
    dfwgt = df3[0+i:i+sstkcount]
    dfwgt['Weight'] = dfwgt['MarketCapitalisationEOD']/dfwgt['MarketCapitalisationEOD'].sum()
    dfwgt['wgtdmcap'] = dfwgt.Weight*dfwgt.MarketCapitalisationEOD
    #bwm = basket weighted mcap
    m0 = dfwgt['wgtdmcap'].sum()
    m1.append(m0)

x=min(m1, key=lambda x:abs(x-wgtavmcap*1000000))
x=m1.index(x)
s1 = df3.iloc[x:x+sstkcount]

s1['Short Weight'] = bsw * s1['MarketCapitalisationEOD']/s1['MarketCapitalisationEOD'].sum()
s1['Target Shares'] = round(s1['Short Weight']*OAMARFFUM/s1['ClosePrice']*100)

col_list=['Security','Target Shares','Short Weight']
#Calculate borrow cost (bc)
bc = s1['Short Weight']*s1['Indic Fee']
borrow_cost=bc.sum()
s1=s1[col_list]
s1.to_csv('S:\Ophir\Investment\Quant\Python\Output\Short Weights.csv',index=False)

#Find row where market cap
#Create Exceptions List & remove from data frame e.g. no avail or high borrow cost

# 3. Remove Portfolio Holdings
# 4. List Weighted Average Market Cap for the 50 stocks above the adjacent stock
# 5. Pick out the stock number and generate the portfolio
#




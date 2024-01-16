import os
import pandas as pd
import openpyxl as pxl
folder = r'D:\2023\July\3G'
var = os.listdir(folder)
print(var)
daily = []
dates = []
for x in var:
    if x.endswith('.xlsx') or x.endswith('.xml') or x.endswith('.xlsb') or x.endswith('.csv') or x.endswith('.xls') or x.endswith('.XML') or x.endswith('.rar') or x.endswith('.csv') or x.endswith('.xlsm') or x.endswith('.MAP'):
        continue
    else:
        y = os.path.join(folder, x)
    for z in os.listdir(y):
        if 'RNC50_daily' in z:
            zfile = z
            dum = os.path.join(y,zfile)
            test = x
            print(test)
            dates.append(test)
            daily.append(dum)
#dum = os.path.join(y,zfile)
print(daily)
print(len(daily))
#excelfile = os.path.join(folder,file)
#wb = openpyxl.load_workbook(excelfile)
#ws = wb['V CSSR']
#print(ws.columns.values)
df1 = pd.DataFrame()
df2 = pd.DataFrame()
df3 = pd.DataFrame()
df4 = pd.DataFrame()
# count = 0
cols = [0,8]
i = 0
counter = 1
for site in daily:
    if site.endswith('.xlsx'):
        print(site)
        print(counter)
        xl = pd.read_excel(site, sheet_name='V CSSR')
        counter+=1
        #sheets = xl.sheet_names
    #if 'V CSSR' in sheets: #some workbooks do not contain the same sheet name
        df = pd.read_excel(site, sheet_name='V CSSR', usecols=[0,8])
        #print(df['Comment'].head())
        #if df[]
        df1 = pd.concat([df1, df])

        #df1 = df1[df['Comment'].str.len() > 0]
        #df1= df1.reindex(df1.columns, axis=1)
    #count = count +1
    #if count==3:
        #break

df1 = df1[df1['Comment'].str.len() > 0]
#df1.to_excel(f'{folder}/test.xlsx', sheet_name='V CSSR', index_label=False)
#tst = df['Comment'].str.contains('True')
#tst = df = df[df['Comment'].str.len() > 0]
#print(tst)
#DCR CONCAT
for site in daily:
    xl = pd.ExcelFile(site)
    sheets = xl.sheet_names
    if 'V DCR' in sheets: #some workbooks do not contain the same sheet name
        df = pd.read_excel(site, sheet_name='V DCR', usecols=[0,9])
        #print(df.head())
        #if df[]
        df2 = pd.concat([df2, df])
#df2['Comment'].replace("#N/A","",inplace=True)
df2 = df2[df2['Comment'].str.len() > 0]
#HS CONCAT
for site in daily:
    xl = pd.ExcelFile(site)
    sheets = xl.sheet_names
    if 'HS CSSR' in sheets: #some worbooks do not contain the same sheet name
        df = pd.read_excel(site, sheet_name='HS CSSR', usecols=[0,8])
        #print(df.head())
        #if df[]
        df3 = pd.concat([df3, df])

df3 = df3[df3['Comment'].str.len() > 0]
df5 = pd.DataFrame()
path = os.path.join(folder, 'July.xlsx')
#path2 = os.path.join(folder, 'Mar2.xlsx')
for site in daily:
    xl = pd.ExcelFile(site)
    sheets = xl.sheet_names
    if 'HS DCR' in sheets: #some workbooks do not contain the same sheet name
        df = pd.read_excel(site, sheet_name='HS DCR', usecols=[0,9])
        df4 = pd.concat([df4, df])


df4 = df4[df4['Comment'].str.len() > 0]
for site in daily:
    xl = pd.ExcelFile(site)
    sheets = xl.sheet_names
    if 'ul load' in sheets: #some workbooks do not contain the same sheet name
        df = pd.read_excel(site, sheet_name='ul load', usecols=[0,8])
        df5 = pd.concat([df5, df])
df5 = df5[df5['Comment'].str.len() > 0]
df1 = df1[df1['Comment'] != 'High AC Fails DL']
df1 = df1[df1['Comment'] != 'High BTS Fails']
df2 = df2[df2['Comment'] != 'High AC Fails DL']
df2 = df2[df2['Comment'] != 'High BTS Fails']
df3 = df3[df3['Comment'] != 'High AC Fails DL']
df3 = df3[df3['Comment'] != 'High BTS Fails']
df4 = df4[df4['Comment'] != 'High BTS Fails']
df4 = df4[df4['Comment'] != 'High AC Fails DL']
df4 = df4[df4['Comment'] != 'low fails']
with pd.ExcelWriter(path) as writer:
    df1.to_excel(writer, sheet_name='V CSSR')
    df2.to_excel(writer, sheet_name='V DCR')
    df3.to_excel(writer, sheet_name='HS CSSR')
    df4.to_excel(writer, sheet_name='HS DCR')
    df5.to_excel(writer, sheet_name='ul load')


#!/usr/bin/env python
# coding: utf-8

# <h1><center>POR</center></h1>

# #### Transpose POR

# In[ ]:


import os
import xlwings as xw
import pandas as pd
import glob
import datetime
import pyodbc
import sqlalchemy as db
from urllib.parse import quote_plus


# In[ ]:


path = r'C:\Users\KohMansf\Documents\STAMS\Waterfall\POR Transpose'
os.chdir(path)
files = glob.glob(path + '/*.xlsb')

latest_por = max(files, key=os.path.getctime)


# In[ ]:


app = xw.App()
book = xw.Book(latest_por)
sheet = book.sheets('POR Data')
latest_por = sheet.range('A1').options(
    pd.DataFrame, expand='table', index=False).value
book.close()
app.kill()


# In[ ]:


latest_por.head()


# In[ ]:


Dates = latest_por.columns[10:]
Col = latest_por.columns[0:10]


# In[ ]:


latest_por.iloc[:, 10:] = latest_por.iloc[:, 10:].fillna(0)
latest_por.tail(10)


# In[ ]:


latest_por = pd.melt(latest_por,
                     id_vars=Col,
                     value_vars=Dates,
                     value_name='Value',
                     var_name='Attribute')
latest_por.iloc[:, 10] = latest_por.iloc[:, 10].apply(pd.to_datetime)


# In[ ]:


latest_por.head(10)


# In[ ]:


latest_por = latest_por.iloc[:, [0, 2, 4, 5, 6, 7, 9, 10, 11]]


# In[ ]:


yyyy_ww = latest_por.iloc[0, 0]

file = 'POR_' + yyyy_ww + '.csv'


# In[ ]:


database_path = r'C:\Users\KohMansf\Documents\STAMS\Waterfall\Database\POR'
os.chdir(database_path)
latest_por.to_csv(file, index=False)


# #### Data Cleaning

# In[ ]:


now = datetime.datetime.now()
year = str(now.year)


# In[ ]:


por_files = os.listdir(database_path)
files_csv = [f for f in por_files if f[4:8] == year]


# In[ ]:


df_por = pd.DataFrame()

for f in files_csv:
    data = pd.read_csv(f, parse_dates=['Attribute'], dayfirst=True)
    df_por = df_por.append(data)

df_por.tail()


# In[ ]:


df_por['Attribute'] = df_por['Attribute'].apply(pd.to_datetime)
df_por['YYYYWW'] = df_por['Attribute'].apply(lambda x: str(
    x.isocalendar()[0]) + str(x.isocalendar()[1]).zfill(2))
df_por.tail()


# **We get Region base on target location (Data from original POR file)**

# In[ ]:


df_region = pd.read_excel(
    r'C:\Users\KohMansf\Documents\STAMS\Waterfall\Database\Country of Target Location\Country of Target Location.xlsx',
    sheet_name='Country of Target Location',
    na_filter=False,
    usecols='A:B')


# In[ ]:


df_merged = pd.merge(df_por, df_region, on='Target Location')
df_merged['Planning Part Group'] = df_merged['Planning Part Group'].str.rstrip(
    'GROUP')
df_merged.head()


# In[ ]:


mpa_dict = {
    'NKG-THAILAND': 'NKG Thailand',
    'NKG Thailand': 'NKG Thailand',
    'NKG-YUEYANG': 'NKG Yue Yang',
    'DSG-VIETNAM': 'DSG Vietnam',
    'DSG Vietnam': 'DSG Vietnam',
    'FLEX-PTP': 'Flex PTP Malasya',
    'Flex PTP Malasya': 'Flex PTP Malasya',
    'FLEX-ZHUHAI': 'Flex Zhuhai',
    'Flex Zhuhai': 'Flex Zhuhai',
    'FOXCONN': 'Foxconn ChongQing',
    'Foxconn ChongQing': 'Foxconn ChongQing',
    'Jabil Circuit De Chihuahua': 'Jabil Circuit De Chihuahua',
    'Jabil Circuit Netherlands BV': 'Jabil Circuit Netherlands BV'
}


# In[ ]:


df_merged['MPA'] = df_merged['MPA'].map(lambda x: mpa_dict.get(x, x))
df_merged.head()


# In[ ]:


df_final_por = df_merged[[
    'Current Cycle on Display', 'Product Line', 'Platform',
    'Planning Part Group', 'Target Location', 'Planning Part', 'MPA', 'region',
    'Value', 'YYYYWW', 'Attribute'
]]
df_final_por


# In[ ]:


df_final_por = df_final_por.rename(
    columns={
        "Current Cycle on Display": "Planning_Wk",
        "Product Line": "Product_Line",
        "Planning Part Group": "Program",
        "Target Location": "Target_Location",
        "Planning Part": "SKU",
        "Value": "Qty",
        "region": "Region",
        "Attribute": "DATES"
    })
df_final_por['QtyType'] = 'POR'
df_final_por.head()


# #### Output File

# In[ ]:


output = year + '_to_upload.csv'


# In[ ]:


df_final_por.to_csv(output, index=False)
os.chdir(r'C:\Users\KohMansf\Documents\STAMS\Exposure Simulator\Database\POR')
df_final_por.to_csv(output, index=False)


# #### Output to Database

# **HP Server**

# In[ ]:


table = 'POR' + year


# In[ ]:


conn = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER=15.46.110.222,1433;DATABASE=POR;UID=Admin;PWD=123789'

quoted = quote_plus(conn)
new_con = 'mssql+pyodbc:///?odbc_connect={}'.format(quoted)
engine = db.create_engine(new_con, fast_executemany=True)

connection = engine.connect()

table_name = table


# In[ ]:


df_final_por.to_sql(table_name,
                    engine,
                    if_exists='replace',
                    chunksize=None,
                    index=False,
                    dtype={
                        'Planning_Wk': db.types.VARCHAR(length=7),
                        'MPA': db.types.VARCHAR(length=50),
                        'SKU': db.types.VARCHAR(length=50),
                        'Program': db.types.VARCHAR(length=50),
                        'Platform': db.types.VARCHAR(length=50),
                        'Product_Line': db.types.VARCHAR(length=2),
                        'Target_Location': db.types.VARCHAR(length=20),
                        'YYYYWW': db.types.INTEGER(),
                        'Region': db.types.VARCHAR(length=8),
                        'Qty': db.types.INTEGER(),
                        'QtyType': db.types.VARCHAR(length=4),
                        'DATES': db.types.Date
                    })
print(df_final_por.head())
print(df_final_por.shape)


# In[ ]:





#!/usr/bin/env python
# coding: utf-8

# <h1><center> FULLSHIPVPOR </center></h1>

# In[ ]:


import pyodbc
import pandas as pd
import numpy as np
import sqlalchemy as db
from urllib.parse import quote_plus
import os


# In[ ]:


path = r'C:\Users\KohMansf\Documents\STAMS\Waterfall\Database\FULLSHIPVPOR'
os.chdir(path)


# #### Get data from database

# **HP Server**

# In[ ]:


conn = "DRIVER={ODBC Driver 17 for SQL Server};SERVER=15.46.110.222,1433;DATABASE=POR;UID=Admin;PWD=123789"
metadata = db.MetaData()
quoted = quote_plus(conn)
new_con = 'mssql+pyodbc:///?odbc_connect={}'.format(quoted)
engine = db.create_engine(new_con)
metadata.reflect(bind=engine)


# **POR Data**

# In[ ]:


df_por = pd.DataFrame()
for key in metadata.tables.keys():
    key = db.Table(key, metadata, autoload=True, autoload_with=engine)
    query = db.select([
        key.columns.Planning_Wk, key.columns.YYYYWW, key.columns.Region,
        key.columns.MPA, key.columns.Qty, key.columns.QtyType,
        key.columns.Platform, key.columns.DATES
    ])
    ResultPOR = engine.connect().execute(query)
    ResultSet = ResultPOR.fetchall()
    df_por = df_por.append(ResultSet)
df_por.columns = ResultSet[0].keys()


# In[ ]:


df_por.head()


# In[ ]:


por_plan = df_por['Planning_Wk'].unique()
por_plan = por_plan.astype(str)
por_plan = np.core.defchararray.replace(por_plan, 'W', '')
por_plan = por_plan.astype(int)
por_plan = np.sort(por_plan)
por_plan


# **HP Server**

# In[ ]:


conn = "DRIVER={ODBC Driver 17 for SQL Server};SERVER=15.46.110.222,1433;DATABASE=SHIPMENT;UID=Admin;PWD=123789"
metadata = db.MetaData()
quoted = quote_plus(conn)
new_con = 'mssql+pyodbc:///?odbc_connect={}'.format(quoted)
engine = db.create_engine(new_con)
metadata.reflect(bind=engine)
SHIPMENT = db.Table('SHIPMENT', metadata, autoload=True, autoload_with=engine)


# **SHIPMENT Data**

# In[ ]:


queryship = db.select([
    SHIPMENT.columns.YYYYWW, SHIPMENT.columns.Region, SHIPMENT.columns.MPA,
    SHIPMENT.columns.Qty, SHIPMENT.columns.QtyType, SHIPMENT.columns.Platform,
    SHIPMENT.columns.DATES
]).order_by(SHIPMENT.columns.YYYYWW, SHIPMENT.columns.Platform)
ResultSHIPMENT = engine.connect().execute(queryship)
ResultSetSHIP = ResultSHIPMENT.fetchall()


# In[ ]:


df_shipment = pd.DataFrame(ResultSetSHIP)
df_shipment.columns = ResultSetSHIP[0].keys()


# In[ ]:


df_shipment.head()


# In[ ]:


def porship_inter(por_plan, YYYYWW):
    df_shipment.loc[(YYYYWW <= por_plan), 'Planning_Wk'] = por_plan


# In[ ]:


df_shipment['Planning_Wk'] = 0
new_df = pd.DataFrame()
for plan in por_plan:
    porship_inter(plan, df_shipment['YYYYWW'])
    new_df = new_df.append(df_shipment)
new_df = new_df[(new_df[['Planning_Wk']] != 0).all(axis=1)]
new_df


# In[ ]:


new_df['Planning_Wk'] = new_df['Planning_Wk'].astype(str)
new_df['Planning_Wk'] = new_df.Planning_Wk.str.slice(
    stop=4) + "W" + new_df.Planning_Wk.str.slice(start=4)


# In[ ]:


new_df = new_df[[
    'Planning_Wk', 'YYYYWW', 'Region', 'MPA', 'Qty', 'QtyType', 'Platform',
    'DATES'
]]
new_df


# In[ ]:


FinalShipVPOR = pd.concat([df_por, new_df], ignore_index=True)
FinalShipVPOR


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


FinalShipVPOR['MPA'] = FinalShipVPOR['MPA'].map(lambda x: mpa_dict.get(x, x))
FinalShipVPOR.head()


# #### Output to file and database

# In[ ]:


FinalShipVPOR.to_csv('FULLSHIPVPOR.csv', index=False)


# In[ ]:


conn = "DRIVER={ODBC Driver 17 for SQL Server};SERVER=15.46.110.222,1433;DATABASE=FULLSHIPVPOR;UID=Admin;PWD=123789"

quoted = quote_plus(conn)
new_con = 'mssql+pyodbc:///?odbc_connect={}'.format(quoted)
engine = db.create_engine(new_con, fast_executemany=True)

connection = engine.connect()

table_name = 'FULLSHIPVPOR'


# In[ ]:


FinalShipVPOR.to_sql(table_name,
                       engine,
                       if_exists='replace',
                       chunksize=None,
                       index=False,
                       dtype={
                           'Planning_Wk': db.types.VARCHAR(length=7),
                           'YYYYWW': db.types.INTEGER(),
                           'Region': db.types.VARCHAR(length=8),
                           'MPA': db.types.VARCHAR(length=50),
                           'Qty': db.types.INTEGER(),
                           'QtyType': db.types.VARCHAR(length=4),
                           'Platform': db.types.VARCHAR(length=50),
                           'DATES': db.types.Date
                       })
print(FinalShipVPOR.head())
print(FinalShipVPOR.shape)


# In[ ]:





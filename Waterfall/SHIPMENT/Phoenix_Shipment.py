#!/usr/bin/env python
# coding: utf-8

# <h1><center> SHIPMENT FILE </center></h1>

# In[ ]:


import pandas as pd
import numpy as np
import os
import pyodbc
import sqlalchemy as db
from urllib.parse import quote_plus


# #### PHOENIX

# In[ ]:


path = r'C:\PHOENIX\CSV_Files'
os.chdir(path)


# In[ ]:


phoenix = pd.read_csv('v_Output_Lite.csv',
                      encoding="ISO-8859-1", index_col='DATA_ELEMENT', low_memory=False)


# In[ ]:


phoenix.head()


# **Filter Data**

# In[ ]:


phoenix_actuals = phoenix.loc['ACTUALS', [
    'PLTFRM_NM', 'PLNG_PART_NR', 'REGION_CD', 'MPA_NM', 'CAL_DAY_DT', '#BUILD_ACTUAL_QT']]


# In[ ]:


phoenix_actuals.head()


# In[ ]:


phoenix_actuals_drop = phoenix_actuals.dropna(
    axis=0, subset=['PLTFRM_NM', 'REGION_CD', 'MPA_NM'])


# In[ ]:


phoenix_actuals_drop = phoenix_actuals_drop.reset_index()


# In[ ]:


phoenix_actuals_drop


# In[ ]:


phoenix_filter = phoenix_actuals_drop.loc[(phoenix_actuals_drop['MPA_NM'] == 'DSG Korea') |
                                          (phoenix_actuals_drop['MPA_NM'] == 'DSG Vietnam') |
                                          (phoenix_actuals_drop['MPA_NM'] == 'NKG Yue Yang') |
                                          (phoenix_actuals_drop['MPA_NM'] == 'NKG Thailand') |
                                          (phoenix_actuals_drop['MPA_NM'] == 'Unknown MPA') |
                                          (phoenix_actuals_drop['MPA_NM'] == 'Foxconn ChongQing') |
                                          (phoenix_actuals_drop['MPA_NM'] == 'Flex Zhuhai') |
                                          (phoenix_actuals_drop['MPA_NM'] == 'Flex PTP Malasya')]


# In[ ]:


phoenix_filter = phoenix_filter.drop(columns={'DATA_ELEMENT'})


# In[ ]:


phoenix_filter['CAL_DAY_DT'] = phoenix_filter['CAL_DAY_DT'].apply(
    pd.to_datetime)


# In[ ]:


phoenix_filter = pd.pivot_table(phoenix_filter, values='#BUILD_ACTUAL_QT',
                                index=['PLTFRM_NM', 'PLNG_PART_NR',
                                       'REGION_CD', 'MPA_NM', 'CAL_DAY_DT'],
                                aggfunc=np.sum).reset_index()


# In[ ]:


phoenix_filter


# In[ ]:


phoenix_filter['YYYYWW'] = phoenix_filter['CAL_DAY_DT'].apply(lambda x: str(x.isocalendar()[0]) +
                                                              str(x.isocalendar()[1] - 1).zfill(2))
phoenix_filter['QtyType'] = 'SHIP'

phoenix_filter = phoenix_filter.loc[phoenix_filter['#BUILD_ACTUAL_QT'] != 0]


# In[ ]:


phoenix_filter.head()


# In[ ]:


latest_date = phoenix_filter['CAL_DAY_DT'].max()


# In[ ]:


shipment_path = r'C:\Users\KohMansf\Documents\STAMS\Waterfall\Database\SHIPMENT'
os.chdir(shipment_path)


# In[ ]:


shipment_read = pd.read_csv('Shipment Data.csv')

shipment_read['CAL_DAY_DT'] = shipment_read['CAL_DAY_DT'].apply(pd.to_datetime)


# In[ ]:


shipment_read


# **If phoenix shipment date is later than current shipment data, we update**

# In[ ]:


shipment_date = shipment_read['CAL_DAY_DT'].max()
if latest_date > shipment_date:
    ship_concat = phoenix_filter.loc[phoenix_filter['CAL_DAY_DT']
                                     == latest_date]
    ship_output = pd.concat([shipment_read, ship_concat])
    ship_output.to_csv('Shipment Data.csv', index=False)


# #### Shipment to database

# **HP Server**

# In[ ]:


conn = "DRIVER={ODBC Driver 17 for SQL Server};SERVER=15.46.110.222,1433;DATABASE=SHIPMENT;UID=Admin;PWD=123789"

quoted = quote_plus(conn)
new_con = 'mssql+pyodbc:///?odbc_connect={}'.format(quoted)
engine = db.create_engine(new_con, fast_executemany=True)

connection = engine.connect()

table_name = 'SHIPMENT'


# In[ ]:


shipment_df = pd.read_csv('Shipment Data.csv')
col = ['Platform', 'SKU', 'Region', 'MPA', 'DATES', 'Qty', 'YYYYWW', 'QtyType']
shipment_df.columns = col
shipment_df.to_sql(table_name,
                   engine,
                   if_exists='replace',
                   chunksize=None,
                   index=False,
                   dtype={
                       'SKU': db.types.VARCHAR(length=50),
                       'Platform': db.types.VARCHAR(length=50),
                       'Region': db.types.VARCHAR(length=50),
                       'MPA': db.types.VARCHAR(length=50),
                       'DATES': db.types.Date,
                       'Qty': db.types.INTEGER(),
                       'YYYYWW': db.types.INTEGER(),
                       'QtyType': db.types.VARCHAR(length=4)
                   })
print(shipment_df.head())
print(shipment_df.shape)


# In[ ]:





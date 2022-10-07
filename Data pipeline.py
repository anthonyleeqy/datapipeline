#!/usr/bin/env python
# coding: utf-8

# # A Data Pipeline to Ingest Data from SQL Server and ready for Model

# In[1]:


#Import packages
import os 
import pathlib
import pymssql
from sqlalchemy import create_engine, text
import bcpy
import pyodbc
import openpyxl
import clevercsv
from icecream import ic
import pyperclip


#import pyperclip


import pandas as pd
server = '#####'
print(server)
db = '#####'
###Server name and database name are redacted. Put SQL and Database name here to proceed.


#processDate ='2021-05-31'
processDate ='2021-12-31'
dataDate = '2021-12-31'  #last business date
lastbizDate = dataDate


pd.set_option("display.max_columns",None)
pd.set_option("display.max_rows",None)

constr = f"mssql+pyodbc://{server}/{db}?driver=SQL+Server"
print(constr)
engine = create_engine(constr,echo=False) 
dbconnection = engine.connect()



#test = dbconnection.execute('SELECT dt=getdate()').fetchall()
#print(test[0])

# to get a clean copy of text to clipboard for testing in SSMS:
def cleanSQL(sql):
    sqlt = sql
    sqlt = sqlt.replace('\n'," ")
    sqlt = sqlt.replace('\t'," ")
    df = pd.DataFrame([])
    pyperclip.copy(sqlt)
    return 


# # Define import file paths and names

# In[ ]:


AExcelFname = r"###.xls"
BExcelFname = r"###.xls"


# # Import A data from Excel file to Database Table 

# In[ ]:


dfAspireXl = pd.read_excel(AExcelFname) #, sheet_name='Loans',skiprows = range(0, 5) ,usecols = "C:BJ")

#number of rows
print(f"rows {len(dfAspireXl.index)}")
#number of cols
print(f"cols {dfAspireXl.shape[1]}")
dfAspireXl.head(2)


# In[ ]:


# upload data to staging table

table_name = 'A'
dfAspireXl.to_sql(table_name, con=dbconnection, if_exists='replace')


# # Import B data from Excel file to Database Table 
# 

# In[ ]:


dfAspireXl = pd.read_excel(BExcelFname) #, sheet_name='Loans',skiprows = range(0, 5) ,usecols = "C:BJ")

#number of rows
print(f"rows {len(dfAspireXl.index)}")
#number of cols
print(f"cols {dfAspireXl.shape[1]}")
dfAspireXl.head(2)


# In[ ]:


# upload data to staging table

table_name = 'B'
dfAspireXl.to_sql(table_name, con=dbconnection, if_exists='replace')


# In[ ]:





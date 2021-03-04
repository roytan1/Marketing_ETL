'''
The objective of this script is to insert and update records from the Source file with the data currently residing in the database.
Check the process flow chart for the business logic defined by the business.
'''

import openpyxl
import pyodbc
import pandas as pd
import os.path
import xlsxwriter
import numpy as np
from pathlib import Path
from configparser import ConfigParser
from fuzzywuzzy import fuzz
from datetime import datetime
import shutil

# Read Config file
config = ConfigParser()
config.read('config.ini')

# pd.show_versions()

# SQL Database Connection
host = config['database']['host']
db = config['database']['db']
pwd = config['database']['user']
uid = config['database']['pass']
dir1 = config['File1']['dir1']
dir2 = config['File2']['dir2']
dir3 = config['File3']['dir3']
dir4 = config['File4']['dir4']
dir5 = config['File5']['dir5']

dt = datetime.today().strftime('%Y%m%d')

# Initial First Upload of MasterList
# src = dir1
# dst = dir4 + "_" + dt + ".xlsx"

# For VendorUpdate_NewValue
src = dir2
dst = dir5 + "_" + dt + ".xlsx"

xpt = dir3 + "_Exception" + "_" + dt + ".xlsx"

conn = pyodbc.connect('driver={ODBC Driver 17 for SQL Server};SERVER='+host+';DATABASE='+db+';UID='+uid+';PWD='+pwd)

cursor = conn.cursor()
cursor.execute(
    'SELECT ID, LinkedInID, CompanyName, CompanyWebsite, EmployeeRange, City, RegionStateProvince, Country, '
    'BusinessClassification, BusinessSubclassification, Active, '
    'Source, LinkedInURL, [Description], [Type], CompanyAddress, Phone, EmployeesonLinkedIn, '
    'Founded, Growth6mth, Growth1yr, Growth2yr, VTSID FROM Marketing.dbo.Marketing_ETL ORDER BY ID'
    # ', UltimateParent, Parent, Subsidiaries '
)

result = cursor.fetchone()
# print(len(result))

fields = [i[0] for i in cursor.description]
db_result = [dict(zip(fields,row)) for row in cursor.fetchall()] 

print(db_result[0]["LinkedInID"])
print(db_result[0][2])
print(db_result[0]["CompanyWebsite"])
print(db_result[0]["EmployeeRange"])

# for row in result:
    # db_Id = row[0]
#     print(row)

conn.close()
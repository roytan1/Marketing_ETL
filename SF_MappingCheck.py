'''
The objective of this script is to update the Mapping Finding status of the Studio from external list against Salesforce client list.
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
pwd = config['database']['pass']
uid = config['database']['user']

host1 = config['database']['host1']
db1 = config['database']['db1']
pwd1 = config['database']['pass1']
uid1 = config['database']['user1']

dt = datetime.today().strftime('%Y%m%d')

conn = pyodbc.connect(
    'driver={ODBC Driver 17 for SQL Server};SERVER=' + host + ';DATABASE=' + db + ';UID=' + uid + ';PWD=' + pwd)

cursor = conn.cursor()
cursor.execute(
    'SELECT ID, CompanyName, Country, SF_Mapping_Status FROM Marketing_ETL ORDER BY ID'
    # ', UltimateParent, Parent, Subsidiaries '
)

result = cursor.fetchall()
# print(len(result))

cursor.close()
# conn.close()

conn1 = pyodbc.connect(
    'driver={ODBC Driver 17 for SQL Server};SERVER=' + host1 + ';DATABASE=' + db1 + ';UID=' + uid1 + ';PWD=' + pwd1)

cursor = conn1.cursor()
cursor.execute(
    "SELECT SFAccountID, AccountName, CASE WHEN BillingCountry='Korea, Republic of' THEN 'South Korea' WHEN BillingCountry='Russian Federation' THEN 'Russia' ELSE BillingCountry END AS BillingCountry FROM [dbo].[BI_PICAccountView] WHERE deleted=0 AND AccountName NOT LIKE '%CLOSE%'"
    # ', UltimateParent, Parent, Subsidiaries '
)

result1 = cursor.fetchall()
# print(len(result))

cursor.close()
# conn.close()


def updateRecord(SF_GUID, SF_Status, db_Id):
    cursor = conn.cursor()
    
    try:
    
        cursor.execute(
            "UPDATE Marketing.dbo.Marketing_ETL SET SF_GUID=?, SF_Mapping_Status=? WHERE ID=?"
            , SF_GUID, SF_Status, db_Id
        )

        conn.commit()

        cursor.close()

    except (pyodbc.Error, pyodbc.Warning) as err:
        print("Update Error on ID = " + str(db_Id))
        err = str(err)


for row in result:
    # print(row[0])
    # Cache database into variable list

    # ID, CompanyName, Country, SF_Mapping_Status

    itemFound = False

    db_Id = row[0]
    db_Name = row[1]
    db_Country = row[2]
    db_Status = row[3]

    if db_Status != 1:

        for row1 in result1:
        
        # SFAccountID, AccountNo, AccountName, BillingCountry
            db1_SFAId = row1[0]
            db1_AccName = row1[1]
            db1_Country = row1[2]

            Ratio = fuzz.ratio(db_Name.upper(), db1_AccName.upper())
            # print (db_Name, db1_AccName, Ratio) 

            if (Ratio >= 85):
                
                # if db_Id == 6:
                #    print (db_Name, db1_AccName, "Pass Name Check") 

                if db_Country == db1_Country:
                
                    # if db_Id == 6:
                    #    print (db_Name, db1_AccName, db_Country, db1_Country, "Pass Country Check") 
                
                    updateRecord(db1_SFAId, 1, db_Id)
                    itemFound = True
                    break
                else:
                    # if db_Id == 6:
                    #    print (db_Name, db1_AccName, db_Country, db1_Country, "Pass Name Check, Fail Country Check") 

                    updateRecord(db1_SFAId, 3, db_Id)
                    itemFound = True
                    break

    if not itemFound:
        #if db_Id == 6:
        #    print (db_Name, db1_AccName, db_Country, db1_Country, "Studio Not Found")
            
        updateRecord('', 2, db_Id)
                
conn.close()
conn1.close()
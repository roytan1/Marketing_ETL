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
    'SELECT SFAccountID FROM SF_Studio_Mapping_TEST'
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
    "SELECT ID, SFAccountID, AccountNo, AccountName, ParentName, ParentNo, AccountOwner, LegalEntityOnContract, AccountType, AccountCurrency, BillToCountryRegion, PaymentTerms, "
    "CASE WHEN BillingCountry='Korea, Republic of' THEN 'South Korea' WHEN BillingCountry='Russian Federation' THEN 'Russia' ELSE BillingCountry END AS BillingCountry, BillingState, "
    "BillingCity, BillingStreet, BillingPostalCode, Deleted, SFCreationDate, SFLastModifiedDate, SFServiceSyncTime, UltimateParent FROM [dbo].[BI_PICAccountView] "
    "WHERE deleted=0 AND AccountName NOT LIKE '%CLOSE%'"
    # ', UltimateParent, Parent, Subsidiaries '
)

result1 = cursor.fetchall()
# print(len(result))

cursor.close()
# conn.close()

def insertRecord(db_ID, db_SFAccountID, db_AccountNo, db_AccountName, db_ParentName, db_ParentNo, db_AccountOwner, db_LegalEntityOnContract, db_AccountType, db_AccountCurrency,
        db_BillToCountryRegion, db_PaymentTerms, db_BillingCountry, db_BillingState, db_BillingCity, db_BillingStreet, db_BillingPostalCode, db_Deleted, db_SFCreationDate, 
        db_SFLastModifiedDate, db_SFServiceSyncTime, db_UltimateParent):

    cursor = conn.cursor()
    
    try:
    
        cursor.execute(
            "INSERT INTO SF_Studio_Mapping_TEST(ID, SFAccountID, AccountNo, AccountName, ParentName, ParentNo, AccountOwner, LegalEntityOnContract, AccountType, AccountCurrency, "
            "BillToCountryRegion, PaymentTerms, BillingCountry, BillingState, BillingCity, BillingStreet, BillingPostalCode, Deleted, SFCreationDate, SFLastModifiedDate, "
            "SFServiceSyncTime, UltimateParent) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            , db_ID, db_SFAccountID, db_AccountNo, db_AccountName, db_ParentName, db_ParentNo, db_AccountOwner, db_LegalEntityOnContract, db_AccountType, db_AccountCurrency,
            db_BillToCountryRegion, db_PaymentTerms, db_BillingCountry, db_BillingState, db_BillingCity, db_BillingStreet, db_BillingPostalCode, db_Deleted, db_SFCreationDate, 
            db_SFLastModifiedDate, db_SFServiceSyncTime, db_UltimateParent
        )

        conn.commit()

        cursor.close()

    except (pyodbc.Error, pyodbc.Warning) as err:
        print("Insert Error on ID = " + str(db_ID))
        err = str(err)


def updateRecord(db_ID, db_ID1):
    cursor = conn.cursor()
    
    try:
    
        cursor.execute(
            "UPDATE Marketing.dbo.SF_Studio_Mapping_TEST SET DBID=? WHERE SFAccountID=?"
            , db_ID, db_ID1
        )

        conn.commit()

        cursor.close()

    except (pyodbc.Error, pyodbc.Warning) as err:
        print("Update Error on ID = " + str(db_ID))
        err = str(err)


for row in result1:
    # print(row[0])
    # Cache database into variable list

    # ID, CompanyName, Country, SF_Mapping_Status

    itemFound = False

    db_ID = row[0]
    db_SFAccountID = row[1]
    db_AccountNo = row[2]
    db_AccountName = row[3]
    db_ParentName = row[4]
    db_ParentNo = row[5]
    db_AccountOwner = row[6]
    db_LegalEntityOnContract = row[7]
    db_AccountType = row[8]
    db_AccountCurrency = row[9]
    db_BillToCountryRegion = row[10]
    db_PaymentTerms = row[11]
    db_BillingCountry = row[12]
    db_BillingState = row[13]
    db_BillingCity = row[14]
    db_BillingStreet = row[15]
    db_BillingPostalCode = row[16]
    db_Deleted = row[17]
    db_SFCreationDate = row[18]
    db_SFLastModifiedDate = row[19]
    db_SFServiceSyncTime = row[20]
    db_UltimateParent = row[21]

    if len(result) == 0:
        # print("1")
        insertRecord(db_ID, db_SFAccountID, db_AccountNo, db_AccountName, db_ParentName, db_ParentNo, db_AccountOwner, db_LegalEntityOnContract, db_AccountType, db_AccountCurrency,
        db_BillToCountryRegion, db_PaymentTerms, db_BillingCountry, db_BillingState, db_BillingCity, db_BillingStreet, db_BillingPostalCode, db_Deleted, db_SFCreationDate, 
        db_SFLastModifiedDate, db_SFServiceSyncTime, db_UltimateParent)

    else:

        for row1 in result:
            db_ID1 = row1[0]
            
            if db_SFAccountID == db_ID1:

                itemFound = True

                break

        if not itemFound:
            
            insertRecord(db_ID, db_SFAccountID, db_AccountNo, db_AccountName, db_ParentName, db_ParentNo, db_AccountOwner, db_LegalEntityOnContract, db_AccountType, db_AccountCurrency,
            db_BillToCountryRegion, db_PaymentTerms, db_BillingCountry, db_BillingState, db_BillingCity, db_BillingStreet, db_BillingPostalCode, db_Deleted, db_SFCreationDate, 
            db_SFLastModifiedDate, db_SFServiceSyncTime, db_UltimateParent)

    #if db_Id == 6:
    #    print (db_Name, db1_AccName, db_Country, db1_Country, "Studio Not Found")
        
#     break

cursor2 = conn.cursor()
cursor2.execute(
    "UPDATE SF_Studio_Mapping_TEST SET [DBID]=B.ID FROM SF_Studio_Mapping_TEST A, (SELECT SF_GUID, ID FROM [dbo].[Marketing_ETL] WHERE SF_Mapping_Status=1) B WHERE A.SFAccountID=B.SF_GUID AND A.DBID=''"
    # ', UltimateParent, Parent, Subsidiaries '
)

conn.commit()
cursor2.close()

conn.close()
conn1.close()
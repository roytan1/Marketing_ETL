'''
The objective of this script is to insert records from the SupportingDevelopers worksheet of GameTitles_MasterList file.
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
dir9 = config['File9']['dir9']
dir10 = config['File10']['dir10']
dir11 = config['File11']['dir11']

dt = datetime.today().strftime('%Y%m%d')
# Initial First Upload of MasterList
src = dir9
dst = dir10 + "_" + dt + ".xlsx"

# For GamesTitle_NewValue
# src = dir2
# dst = dir5 + "_" + dt + ".xlsx"

xptSupDev = dir11 + "_SupDev_Exception" + "_" + dt + ".xlsx"

# uft = dir11 + "_UnFound_Title" + "_" + dt + ".xlsx"

conn = pyodbc.connect(
    'driver={ODBC Driver 17 for SQL Server};SERVER=' + host + ';DATABASE=' + db + ';UID=' + uid + ';PWD=' + pwd)

cursor = conn.cursor()
cursor.execute(
    'SELECT ID, TitleID, TitleName, IGDB_Website, NewZoo_Website, SupportDev, StudioID, VTSID FROM Marketing.dbo.GamesTitles_SupDev ORDER BY ID'
)

result = cursor.fetchall()

cursor.close()

cursor1 = conn.cursor()
cursor1.execute(
    'SELECT ID, TitleID, TitleName FROM Marketing.dbo.GamesTitles ORDER BY ID'
)

result1 = cursor1.fetchall()

cursor1.close()

# Create Developer ExceptionCatch file and Remove file if exists
my_file = Path('%s' % (xptSupDev))
if my_file.is_file():
    os.remove('%s' % (xptSupDev))
else:
    # print("The file does not exist")
    print("No Support Developer Exception File found ....")

print("Creating Support Developer Exception File ....")

# For first upload of Developer worksheet
df = pd.read_excel('%s' % (dir9), sheet_name='SupportingDevelopers', engine='openpyxl')
# For VendorUpdate_NewValue
# df = pd.read_excel('%s' %(dir2), sheet_name='Titles', engine='openpyxl')

df = df.replace(np.nan, ' ', regex=True)
df_Header = df.columns.ravel()

# Create Developer Exception Worksheet
workbook = xlsxwriter.Workbook('%s' % (xptSupDev))
worksheet = workbook.add_worksheet()

for col_num, data in enumerate(df_Header):
    # print(data)
    if col_num <= 6:
        worksheet.write(0, col_num, data)

worksheet.write(0, 7, 'Error')

ExptRow = 1
# ufRow = 1

def InsertRecord(db_ID1, TId, Name, IGDB, NewZoo, SupportDev, StudioID, VTSID, ExptRow):
    cursor3 = conn.cursor()

    try:
        # print("First Upload..... ")
        cursor3.execute(
            "INSERT INTO GamesTitles_SupDev(ID, TitleID, TitleName, IGDB_Website, NewZoo_Website, SupportDev, StudioID, VTSID) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
            , db_ID1, TId, Name, IGDB, NewZoo, SupportDev, StudioID, VTSID
        )

        conn.commit()
        cursor3.close()

    except (pyodbc.Error, pyodbc.Warning) as err:
        print("Insert Error on TitleName_SupportDev = " + str(Name))
        err = str(err)

        worksheet.write(ExptRow, 0, TId)
        worksheet.write(ExptRow, 1, Name)
        worksheet.write(ExptRow, 2, IGDB)
        worksheet.write(ExptRow, 3, NewZoo)
        worksheet.write(ExptRow, 4, SupportDev)
        worksheet.write(ExptRow, 5, StudioID)
        worksheet.write(ExptRow, 6, VTSID)
        worksheet.write(ExptRow, 7, err)

        ExptRow = ExptRow + 1

        return ExptRow

# Retrieving of SupportingDevelopers worksheet data
for ir in range(0, len(df)):
    # for ir1 in range(8, 9):
    for ic in range(0, len(df.columns)):
        if ic == 1:

            itemFound = False
            itemFound1 = False

            if df.iat[ir, 0] not in ('0', '0.0', '#N/A', '', ' '):
                TId = int(df.iat[ir, 0])
            else:
                TId = str('')

            Name = str(df.iat[ir, 1])
            if Name in ('0', '0.0', '', ' ', '$N/A'):
                Name = ''

            IGDB = str(df.iat[ir, 2])
            if IGDB in ('0', '0.0', '', ' ', '$N/A'):
                IGDB = ''

            NewZoo = str(df.iat[ir, 3])
            if NewZoo in ('0', '0.0', '', ' ', '$N/A'):
                NewZoo = ''

            SupportDev = str(df.iat[ir, 4])
            if SupportDev in ('0', '0.0', '', ' ', '$N/A'):
                SupportDev = ''

            if df.iat[ir, 5] not in ('0', '0.0', '#N/A', '', ' '):
                StudioID = int(df.iat[ir, 5])
            else:
                StudioID = str('')

            VTSID = str(df.iat[ir, 6])
            if VTSID in ('0', '0.0', '', ' ', '$N/A'):
                VTSID = ''

            # A. Initial upload when table is empty in database
            if len(result) == 0:
                # if str(Name) != '':
                for row1 in result1:

                    db_ID1 = row1[0]
                    db_TID1 = row1[1]
                    db_Name1 = row1[2]

                    if str(Name) == str(db_Name1):

                        itemFound = True

                        InsertRecord(db_ID1, TId, Name, IGDB, NewZoo, SupportDev, StudioID, VTSID, ExptRow)

            # A. Database is not blank, NewValue file upload
            else:
                for row in result:

                    db_ID = row[0]
                    db_TId = row[1]
                    db_Name = row[2]
                    db_IGDB = row[3]
                    db_NewZoo = row[4]
                    db_SupDev = row[5]
                    db_StudioID = row[6]
                    db_VTSID = row[7]

                    # if Title + Supporting Developer found in SupportingDeveloper table in DB. Go to next Record from worksheet.
                    if str(Name) == str(db_Name) and str(SupportDev) == str(db_SupDev) and str(IGDB) == str(db_IGDB):
                        itemFound = True
                        break

                # if Title + Supporting Developer not found in SupportingDeveloper table in DB. Inserrt record into the SupportingDeveloper table.
                if not itemFound:
                    for row2 in result1:

                        db_ID1 = row2[0]
                        db_TID1 = row2[1]
                        db_Name1 = row2[2]

                        # if Title Found in title table
                        if str(Name) == str(db_Name1):
                            itemFound1 = True

                            InsertRecord(db_ID1, TId, Name, IGDB, NewZoo, SupportDev, StudioID, VTSID, ExptRow)

                    # If title not found after checking GamesTitle list, write exception to exception file.
                    if not itemFound1:
                        err = str("Title not found in Title table. Please kindly check the title list.")

                        worksheet.write(ExptRow, 0, TId)
                        worksheet.write(ExptRow, 1, Name)
                        worksheet.write(ExptRow, 2, IGDB)
                        worksheet.write(ExptRow, 3, NewZoo)
                        worksheet.write(ExptRow, 4, SupportDev)
                        worksheet.write(ExptRow, 5, StudioID)
                        worksheet.write(ExptRow, 6, VTSID)
                        worksheet.write(ExptRow, 7, err)

                        ExptRow = ExptRow + 1
        # End of Developer File Load

my_file = Path(dst)
if my_file.is_file():
    os.remove('%s' % (dst))
    shutil.copy(src, dst)
else:
    shutil.copy(src, dst)

workbook.close()

# if no exception, remove the file
if ExptRow == 1:
    os.remove('%s' % (xptSupDev))

conn.close()
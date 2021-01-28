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
dir9 = config['File9']['dir9']
dir10 = config['File10']['dir10']
dir11 = config['File11']['dir11']

dt = datetime.today().strftime('%Y%m%d')
# Initial First Upload of MasterList
src = dir9
dst = dir10 + "_" + dt + ".xlsx"

# For VendorUpdate_NewValue
# src = dir2
# dst = dir5 + "_" + dt + ".xlsx"

xptDev = dir11 + "_Dev_Exception" + "_" + dt + ".xlsx"
# xptPub = dir11 + "_Pub_Exception" + "_" + dt + ".xlsx"
# xptSupDev = dir11 + "_SupDev_Exception" + "_" + dt + ".xlsx"
# xptPortDev = dir11 + "_PortDev_Exception" + "_" + dt + ".xlsx"
# xptRelDate = dir11 + "_RelDate_Exception" + "_" + dt + ".xlsx"

# uft = dir11 + "_UnFound_Title" + "_" + dt + ".xlsx"

conn = pyodbc.connect(
    'driver={ODBC Driver 17 for SQL Server};SERVER=' + host + ';DATABASE=' + db + ';UID=' + uid + ';PWD=' + pwd)

cursor = conn.cursor()
cursor.execute(
    'SELECT ID, TitleID, TitleName, IGDB_Website, NewZoo_Website, Developer, StudioID, VTSID FROM Marketing.dbo.GamesTitles_Dev ORDER BY ID'
)

result = cursor.fetchall()

cursor.close()

'''
cursor1 = conn.cursor()
cursor1.execute(
    'SELECT ID, TitleID, TitleName, IGDB_Website, NewZoo_Website, Publisher, StudioID, VTSID FROM Marketing.dbo.GamesTitles_Pub ORDER BY ID'
)

result1 = cursor1.fetchall()

cursor1.close()

cursor2 = conn.cursor()
cursor2.execute(
    'SELECT ID, TitleID, TitleName, IGDB_Website, NewZoo_Website, SupportDev, StudioID, VTSID FROM Marketing.dbo.GamesTitles_SupDev ORDER BY ID'
)

result2 = cursor2.fetchall()

cursor2.close()

cursor3 = conn.cursor()
cursor3.execute(
    'SELECT ID, TitleID, TitleName, IGDB_Website, NewZoo_Website, PortDev, StudioID, VTSID FROM Marketing.dbo.GamesTitles_PortDev ORDER BY ID'
)

result3 = cursor3.fetchall()

cursor3.close()

cursor4 = conn.cursor()
cursor4.execute(
    'SELECT ID, TitleID, TitleName, ReleaseDate, IGDB_Website, NewZoo_Website FROM Marketing.dbo.GamesTitles_RelDate ORDER BY ID'
)

result4 = cursor4.fetchall()

cursor4.close()

'''

cursor1 = conn.cursor()
cursor1.execute(
    'SELECT ID, TitleID, TitleName FROM Marketing.dbo.GamesTitles ORDER BY ID'
)

result1 = cursor1.fetchall()

cursor1.close()


# Create Developer ExceptionCatch file and Remove file if exists
my_file = Path('%s' % (xptDev))
if my_file.is_file():
    os.remove('%s' % (xptDev))
else:
    # print("The file does not exist")
    print("No Developer Exception File found ....")

print("Creating Developer Exception File ....")


'''
# Create Unfound Title file and Remove file if exists
my_file1 = Path('%s' % (uft))
if my_file1.is_file():
    os.remove('%s' % (uft))
else:
    # print("The file does not exist")
    print("No UnFound Title File found ....")

print("Creating UnFound Title File ....")
'''

# For first upload of Developer worksheet
df = pd.read_excel('%s' % (dir9), sheet_name='Developers', engine='openpyxl')
# For VendorUpdate_NewValue
# df = pd.read_excel('%s' %(dir2), sheet_name='Titles', engine='openpyxl')

df = df.replace(np.nan, ' ', regex=True)
df_Header = df.columns.ravel()

# Create Developer Exception Worksheet
workbook = xlsxwriter.Workbook('%s' % (xptDev))
worksheet = workbook.add_worksheet()

for col_num, data in enumerate(df_Header):
    # print(data)
    if col_num <= 6:
        worksheet.write(0, col_num, data)

worksheet.write(0, 7, 'Error')

'''

# For first upload of Publisher worksheet
df1 = pd.read_excel('%s' % (dir9), sheet_name='Publishers', engine='openpyxl')
# For VendorUpdate_NewValue
# df = pd.read_excel('%s' %(dir2), sheet_name='Titles', engine='openpyxl')

df1 = df1.replace(np.nan, ' ', regex=True)
df_Header1 = df1.columns.ravel()

# Create Publisher Exception Worksheet
workbook1 = xlsxwriter.Workbook('%s' % (xptPub))
worksheet1 = workbook1.add_worksheet()

for col_num, data in enumerate(df_Header1):
    # print(data)
    if col_num <= 6:
        worksheet1.write(0, col_num, data)

worksheet1.write(0, 7, 'Error')


# For first upload of SupportingDevelopers worksheet
df2 = pd.read_excel('%s' % (dir9), sheet_name='SupportingDevelopers', engine='openpyxl')
# For VendorUpdate_NewValue
# df = pd.read_excel('%s' %(dir2), sheet_name='Titles', engine='openpyxl')

df2 = df2.replace(np.nan, ' ', regex=True)
df_Header2 = df2.columns.ravel()

# Create Exception Worksheet
workbook2 = xlsxwriter.Workbook('%s' % (xptSupDev))
worksheet2 = workbook2.add_worksheet()

for col_num, data in enumerate(df_Header2):
    # print(data)
    if col_num <= 6:
        worksheet.write(0, col_num, data)

worksheet2.write(0, 7, 'Error')


# For first upload of PortingDevelopers Worksheet
df3 = pd.read_excel('%s' % (dir9), sheet_name='PortingDevelopers', engine='openpyxl')
# For VendorUpdate_NewValue
# df = pd.read_excel('%s' %(dir2), sheet_name='Titles', engine='openpyxl')

df3 = df3.replace(np.nan, ' ', regex=True)
df_Header3 = df3.columns.ravel()

# Create Exception Worksheet
workbook3 = xlsxwriter.Workbook('%s' % (xptPortDev))
worksheet3 = workbook3.add_worksheet()

for col_num, data in enumerate(df_Header3):
    # print(data)
    if col_num <= 6:
        worksheet3.write(0, col_num, data)

worksheet3.write(0, 7, 'Error')


# For first upload of Release Date Worksheet
df4 = pd.read_excel('%s' % (dir9), sheet_name='ReleaseDates', engine='openpyxl')
# For VendorUpdate_NewValue
# df = pd.read_excel('%s' %(dir2), sheet_name='Titles', engine='openpyxl')

df4 = df4.replace(np.nan, ' ', regex=True)
df_Header4 = df4.columns.ravel()

# Create Exception Worksheet
workbook4 = xlsxwriter.Workbook('%s' % (xptRelDate))
worksheet4 = workbook4.add_worksheet()

for col_num, data in enumerate(df_Header4):
    # print(data)
    if col_num <= 5:
        worksheet.write(0, col_num, data)

worksheet.write(0, 6, 'Error')

# Create UnFound Worksheet
workbook1 = xlsxwriter.Workbook('%s' % (uft))
worksheet1 = workbook1.add_worksheet()

for col_num, data in enumerate(df_Header):
    # print(data)
    if col_num <= 12:
        worksheet1.write(0, col_num, data)

# workbook.close()
'''

ExptRow = 1
# ufRow = 1

# print(len(df))
# print(len(df.columns))

# Check for identical company

# value1 = 1
# value2 = 1

# Start of Developer File Load
for ir in range(0, len(df)):
    # for ir1 in range(8, 9):
    for ic in range(0, len(df.columns)):
        # print(ir1, ic1)
        if ic == 1:
            #print("First Loop", value1)

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

            Developers = str(df.iat[ir, 4])
            if Developers in ('0', '0.0', '', ' ', '$N/A'):
                Developers = ''

            if df.iat[ir, 5] not in ('0', '0.0', '#N/A', '', ' '):
                StudioID = int(df.iat[ir, 5])
            else:
                StudioID = str('')

            VTSID = str(df.iat[ir, 6])
            if VTSID in ('0', '0.0', '', ' ', '$N/A'):
                VTSID = ''

            # print(TId, Name, IGDB, NewZoo, Developers, StudioID, VTSID, "Start")

            # A. If database is blank, it is an initial upload
            if len(result) == 0:

                # if str(Name) != '':

                for row1 in result1:

                    db_ID1 = row1[0]
                    db_TID1 = row1[1]
                    db_Name1 = row1[2]

                    # print(db_ID1, TId, Name, db_Name1, 1)

                    if str(Name) == str(db_Name1):

                        # print("Second Loop", value2)

                        # print(db_ID1, db_TID1, db_Name1, TId, Name, Developers, StudioID, VTSID, 2)

                        itemFound = True

                        cursor2 = conn.cursor()

                        try:
                            # print("First Upload..... ")
                            cursor2.execute(
                                "INSERT INTO GamesTitles_Dev(ID, TitleID, TitleName, IGDB_Website, NewZoo_Website, Developer, StudioID, VTSID) "
                                "VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
                                , db_ID1, TId, Name, IGDB, NewZoo, Developers, StudioID, VTSID
                                )

                            conn.commit()
                            cursor2.close()

                        except (pyodbc.Error, pyodbc.Warning) as err:
                            print("Insert Error on TitleName_Dev = " + str(Name))
                            err = str(err)

                            worksheet.write(ExptRow, 0, TId)
                            worksheet.write(ExptRow, 1, Name)
                            worksheet.write(ExptRow, 2, IGDB)
                            worksheet.write(ExptRow, 3, NewZoo)
                            worksheet.write(ExptRow, 4, Developers)
                            worksheet.write(ExptRow, 5, StudioID)
                            worksheet.write(ExptRow, 6, VTSID)
                            worksheet.write(ExptRow, 7, err)

                            ExptRow = ExptRow + 1

                    # value2 + 1
            # A. Database is not blank, NewValue file upload
            else:
                # print("Test1")
                for row in result:

                    db_ID = row[0]
                    db_TId = row[1]
                    db_Name = row[2]
                    db_IGDB = row[3]
                    db_NewZoo = row[4]
                    db_Developer = row[5]
                    db_StudioID = row[6]
                    db_VTSID = row[7]

                    # if Title with different publisher Found in Publisher table
                    # if str(Valid) != str(db_Valid):
                    if str(Name) == str(db_Name) and str(Developers) == str(db_Developer) and str(IGDB) == str(db_IGDB):
                        itemFound = True
                        # print(Name, db_Name, Publishers, db_Publisher, 1)
                        break

                # New Record found
                if not itemFound:

                    for row2 in result1:

                        db_ID1 = row2[0]
                        db_TID1 = row2[1]
                        db_Name1 = row2[2]

                        # if Title Found in title table
                        if str(Name) == str(db_Name1):
                            # print(Name, db_Name, Publishers, db_Publisher, db_Name1, 2)
                            itemFound1 = True

                            # print(db_ID1, TId, Name, db_Name, Developers, db_Name1, 3)

                            cursor3 = conn.cursor()

                            try:
                                # print("First Upload..... ")
                                cursor3.execute(
                                    "INSERT INTO GamesTitles_Pub(ID, TitleID, TitleName, IGDB_Website, NewZoo_Website, Developer, StudioID, VTSID) "
                                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
                                    , db_ID1, TId, Name, IGDB, NewZoo, Developers, StudioID, VTSID
                                )

                                conn.commit()
                                cursor3.close()

                            except (pyodbc.Error, pyodbc.Warning) as err:
                                print("Insert Error on TitleName_Developer = " + str(Name))
                                err = str(err)

                                worksheet.write(ExptRow, 0, TId)
                                worksheet.write(ExptRow, 1, Name)
                                worksheet.write(ExptRow, 2, IGDB)
                                worksheet.write(ExptRow, 3, NewZoo)
                                worksheet.write(ExptRow, 4, Developers)
                                worksheet.write(ExptRow, 5, StudioID)
                                worksheet.write(ExptRow, 6, VTSID)
                                worksheet.write(ExptRow, 7, err)

                                ExptRow = ExptRow + 1

                    if not itemFound1:
                        # print("Insert Error on TitleName_Publisher = " + str(Name))
                        err = str("Title not found in Title table. Please kindly check the title list.")

                        worksheet.write(ExptRow, 0, TId)
                        worksheet.write(ExptRow, 1, Name)
                        worksheet.write(ExptRow, 2, IGDB)
                        worksheet.write(ExptRow, 3, NewZoo)
                        worksheet.write(ExptRow, 4, Developers)
                        worksheet.write(ExptRow, 5, StudioID)
                        worksheet.write(ExptRow, 6, VTSID)
                        worksheet.write(ExptRow, 7, err)

                        ExptRow = ExptRow + 1

        # End of Developer File Load

        # value1 + 1

my_file = Path(dst)
if my_file.is_file():
    os.remove('%s' % (dst))
    shutil.copy(src, dst)
else:
    shutil.copy(src, dst)

workbook.close()
# workbook1.close()
# workbook2.close()
# workbook3.close()
# workbook4.close()

# if no exception, remove the file
if ExptRow == 1:
    os.remove('%s' % (xptDev))

conn.close()
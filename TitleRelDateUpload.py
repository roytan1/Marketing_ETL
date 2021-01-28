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

xptRelDate = dir11 + "_RelDate_Exception" + "_" + dt + ".xlsx"

# uft = dir11 + "_UnFound_Title" + "_" + dt + ".xlsx"

conn = pyodbc.connect(
    'driver={ODBC Driver 17 for SQL Server};SERVER=' + host + ';DATABASE=' + db + ';UID=' + uid + ';PWD=' + pwd)

cursor = conn.cursor()
cursor.execute(
    'SELECT ID, TitleID, TitleName, ReleasePlatform, ReleaseDate, IGDB_Website, NewZoo_Website FROM Marketing.dbo.GamesTitles_RelDate ORDER BY ID'
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
my_file = Path('%s' % (xptRelDate))
if my_file.is_file():
    os.remove('%s' % (xptRelDate))
else:
    # print("The file does not exist")
    print("No Release Date Exception File found ....")

print("Creating Release Date Exception File ....")


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
df = pd.read_excel('%s' % (dir9), sheet_name='ReleaseDates', engine='openpyxl')
# For VendorUpdate_NewValue
# df = pd.read_excel('%s' %(dir2), sheet_name='Titles', engine='openpyxl')

df = df.replace(np.nan, ' ', regex=True)
df_Header = df.columns.ravel()

# Create Developer Exception Worksheet
workbook = xlsxwriter.Workbook('%s' % (xptRelDate))
worksheet = workbook.add_worksheet()

for col_num, data in enumerate(df_Header):
    # print(data)
    if col_num <= 5:
        worksheet.write(0, col_num, data)

worksheet.write(0, 6, 'Error')

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

            RelPlatform = str(df.iat[ir, 4])
            if RelPlatform in ('0', '0.0', '', ' ', '$N/A'):
                RelPlatform = ''

            RelDate = str(df.iat[ir, 5])
            if RelDate in ('0', '0.0', '', ' ', '$N/A'):
                RelDate = ''

            IGDB = str(df.iat[ir, 10])
            if IGDB in ('0', '0.0', '', ' ', '$N/A'):
                IGDB = ''

            NewZoo = str(df.iat[ir, 11])
            if NewZoo in ('0', '0.0', '', ' ', '$N/A'):
                NewZoo = ''

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
                                "INSERT INTO GamesTitles_RelDate(ID, TitleID, TitleName, ReleasePlatform, ReleaseDate, IGDB_Website, NewZoo_Website) "
                                "VALUES (?, ?, ?, ?, ?, ?, ?)"
                                , db_ID1, TId, Name, RelPlatform, RelDate, IGDB, NewZoo
                                )

                            conn.commit()
                            cursor2.close()

                        except (pyodbc.Error, pyodbc.Warning) as err:
                            print("Insert Error on TitleName_RelDate = " + str(Name))
                            err = str(err)

                            worksheet.write(ExptRow, 0, TId)
                            worksheet.write(ExptRow, 1, Name)
                            worksheet.write(ExptRow, 2, RelPlatform)
                            worksheet.write(ExptRow, 3, RelDate)
                            worksheet.write(ExptRow, 4, IGDB)
                            worksheet.write(ExptRow, 5, NewZoo)
                            worksheet.write(ExptRow, 6, err)

                            ExptRow = ExptRow + 1

                    # value2 + 1
            # A. Database is not blank, NewValue file upload
            else:
                # print("Test1")
                for row in result:

                    db_ID = row[0]
                    db_TId = row[1]
                    db_Name = row[2]
                    db_RelPlatform = row[3]
                    db_RelDate = row[4]
                    db_IGDB = row[5]
                    db_NewZoo = row[6]

                    # if Title with different publisher Found in Publisher table
                    # if str(Valid) != str(db_Valid):
                    if str(Name) == str(db_Name) and str(RelPlatform) == str(db_RelPlatform) and str(IGDB) == str(db_IGDB):
                        itemFound = True
                        print(Name, db_Name, RelPlatform, db_RelPlatform, RelDate, db_RelDate, 1)

                        if str(RelDate) != '' and str(RelDate) != str(db_RelDate):
                            print(Name, db_Name, RelPlatform, db_RelPlatform, RelDate, db_RelDate, 2)

                            cursor4 = conn.cursor()

                            try:
                                # print("First Upload..... ")
                                cursor4.execute(
                                    "UPDATE Marketing.dbo.GamesTitles_RelDate SET ReleaseDate=? WHERE TitleName=? AND ReleasePlatform=? AND IGDB_Website=? "
                                    , RelDate, Name, RelPlatform, IGDB
                                )

                                conn.commit()
                                cursor4.close()

                            except (pyodbc.Error, pyodbc.Warning) as err:
                                print("Update Error on TitleName_RelDate = " + str(Name))
                                err = str(err)

                                worksheet.write(ExptRow, 0, TId)
                                worksheet.write(ExptRow, 1, Name)
                                worksheet.write(ExptRow, 2, RelPlatform)
                                worksheet.write(ExptRow, 3, RelDate)
                                worksheet.write(ExptRow, 4, IGDB)
                                worksheet.write(ExptRow, 5, NewZoo)
                                worksheet.write(ExptRow, 6, err)

                                ExptRow = ExptRow + 1

                            break
                        else:
                            print(Name, db_Name, RelPlatform, db_RelPlatform, RelDate, db_RelDate, 3)
                            break

                # New Record found
                if not itemFound:
                    print(Name, db_Name, RelPlatform, db_RelPlatform, RelDate, db_RelDate, 4)

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
                                    "INSERT INTO GamesTitles_RelDate(ID, TitleID, TitleName, ReleasePlatform, ReleaseDate, IGDB_Website, NewZoo_Website) "
                                    "VALUES (?, ?, ?, ?, ?, ?, ?)"
                                    , db_ID1, TId, Name, RelPlatform, RelDate, IGDB, NewZoo
                                )

                                conn.commit()
                                cursor3.close()

                            except (pyodbc.Error, pyodbc.Warning) as err:
                                print("Insert Error on TitleName_RelDate = " + str(Name))
                                err = str(err)

                                worksheet.write(ExptRow, 0, TId)
                                worksheet.write(ExptRow, 1, Name)
                                worksheet.write(ExptRow, 2, RelPlatform)
                                worksheet.write(ExptRow, 3, RelDate)
                                worksheet.write(ExptRow, 4, IGDB)
                                worksheet.write(ExptRow, 5, NewZoo)
                                worksheet.write(ExptRow, 6, err)

                                ExptRow = ExptRow + 1

                    if not itemFound1:
                        # print("Insert Error on TitleName_Publisher = " + str(Name))
                        err = str("Title not found in Title table. Please kindly check the title list.")

                        worksheet.write(ExptRow, 0, TId)
                        worksheet.write(ExptRow, 1, Name)
                        worksheet.write(ExptRow, 2, RelPlatform)
                        worksheet.write(ExptRow, 3, RelDate)
                        worksheet.write(ExptRow, 4, IGDB)
                        worksheet.write(ExptRow, 5, NewZoo)
                        worksheet.write(ExptRow, 6, err)

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
    os.remove('%s' % (xptRelDate))

conn.close()
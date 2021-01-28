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

xpt = dir11 + "_Exception" + "_" + dt + ".xlsx"
uft = dir11 + "_UnFound_Title" + "_" + dt + ".xlsx"

conn = pyodbc.connect(
    'driver={ODBC Driver 17 for SQL Server};SERVER=' + host + ';DATABASE=' + db + ';UID=' + uid + ';PWD=' + pwd)

cursor = conn.cursor()
cursor.execute(
    'SELECT TitleID, TitleName, Metascore, GameModes, Genre, Themes, Series, PlayerPerspectives, Franchises, '
    'GameEngine, AlternativeNames, IGDB_Website, NewZoo_Website, ID FROM Marketing.dbo.GamesTitles ORDER BY ID'
    # ', UltimateParent, Parent, Subsidiaries '
)

result = cursor.fetchall()
# print(len(result))

cursor.close()
# conn.close()

# Create ExceptionCatch file and Remove file if exists
my_file = Path('%s' %(xpt))
if my_file.is_file():
    os.remove('%s' %(xpt))
else:
    # print("The file does not exist")
    print("No Exception File found ....")

print("Creating Exception File ....")

# Create Unfound Title file and Remove file if exists
my_file1 = Path('%s' %(uft))
if my_file1.is_file():
    os.remove('%s' %(uft))
else:
    # print("The file does not exist")
    print("No UnFound Title File found ....")

print("Creating UnFound Title File ....")

# For first upload of MasterList
df = pd.read_excel('%s' %(dir9), sheet_name='Titles', engine='openpyxl')
# For VendorUpdate_NewValue
# df = pd.read_excel('%s' %(dir2), sheet_name='Titles', engine='openpyxl')

df = df.replace(np.nan, ' ', regex=True)
df_Header = df.columns.ravel()

# Create Exception Worksheet
workbook = xlsxwriter.Workbook('%s' % (xpt))
worksheet = workbook.add_worksheet()

for col_num, data in enumerate(df_Header):
    # print(data)
    if col_num <= 12:
        worksheet.write(0, col_num, data)

worksheet.write(0, 13, 'Error')

# Create UnFound Worksheet
workbook1 = xlsxwriter.Workbook('%s' % (uft))
worksheet1 = workbook1.add_worksheet()

for col_num, data in enumerate(df_Header):
    # print(data)
    if col_num <= 12:
        worksheet1.write(0, col_num, data)

# workbook.close()

ExptRow = 1
ufRow = 1

# print(len(df))
# print(len(df.columns))

# Check for identical company

for ir1 in range(0, len(df)):
    # for ir1 in range(8, 9):
    for ic1 in range(0, len(df.columns)):
        # print(ir1, ic1)

        # IGDB Column
        if (ic1 == 11):
            itemFound = False

            # Id = int(df.iat[ir1, 0])

            if df.iat[ir1, 0] not in ('0', '0.0', '#N/A', '', ' '):
                TId = int(df.iat[ir1, 0])
            else:
                TId = str('')

            Name = str(df.iat[ir1, 1])
            if Name in ('0', '0.0', '', ' ', '$N/A'):
                Name = ''

            Meta = str(df.iat[ir1, 2])
            if Meta in ('0', '0.0', '', ' ', '$N/A'):
                Meta = ''

            Mode = str(df.iat[ir1, 3])
            if Mode in ('0', '0.0', '', ' ', '$N/A'):
                Mode = ''

            Genre = str(df.iat[ir1, 4])
            if Genre in ('0', '0.0', '', ' ', '$N/A'):
                Genre = ''

            Themes = str(df.iat[ir1, 5])
            if Themes in ('0', '0.0', '', ' ', '$N/A'):
                Themes = ''

            Series = str(df.iat[ir1, 6])
            if Series in ('0', '0.0', '', ' ', '$N/A'):
                Series = ''

            PP = str(df.iat[ir1, 7])
            if PP in ('0', '0.0', '', ' ', '$N/A'):
                PP = ''

            Franchises = str(df.iat[ir1, 8])
            if Franchises in ('0', '0.0', '', ' ', '$N/A'):
                Franchises = ''

            Engine = str(df.iat[ir1, 9])
            if Engine in ('0', '0.0', '', ' ', '$N/A'):
                Engine = ''

            AltName = str(df.iat[ir1, 10])
            if AltName in ('0', '0.0', '', ' ', '$N/A'):
                AltName = ''

            IGDB = str(df.iat[ir1, 11])
            if IGDB in ('0', '0.0', '', ' ', '$N/A'):
                IGDB = ''

            NewZoo = str(df.iat[ir1, 12])
            if NewZoo in ('0', '0.0', '', ' ', '$N/A'):
                NewZoo = ''

            # A. If database is blank, it is an initial upload
            if len(result) == 0:

                cursor2 = conn.cursor()

                try:
                    # print("First Upload..... ")
                    cursor2.execute(
                        "INSERT INTO GamesTitles(TitleID, TitleName, Metascore, GameModes, Genre, Themes, Series, PlayerPerspectives, "
                        "Franchises, GameEngine, AlternativeNames, IGDB_Website, NewZoo_Website) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                        , TId, Name, Meta, Mode, Genre, Themes, Series, PP, Franchises, Engine, AltName, IGDB, NewZoo
                    )

                    conn.commit()
                    cursor2.close()

                except (pyodbc.Error, pyodbc.Warning) as err:
                    print("Insert Error on TitleName = " + str(Name))
                    err = str(err)

                    worksheet.write(ExptRow, 0, TId)
                    worksheet.write(ExptRow, 1, Name)
                    worksheet.write(ExptRow, 2, Meta)
                    worksheet.write(ExptRow, 3, Mode)
                    worksheet.write(ExptRow, 4, Genre)
                    worksheet.write(ExptRow, 5, Themes)
                    worksheet.write(ExptRow, 6, Series)
                    worksheet.write(ExptRow, 7, PP)
                    worksheet.write(ExptRow, 8, Franchises)
                    worksheet.write(ExptRow, 9, Engine)
                    worksheet.write(ExptRow, 10, AltName)
                    worksheet.write(ExptRow, 11, IGDB)
                    worksheet.write(ExptRow, 12, NewZoo)
                    worksheet.write(ExptRow, 13, err)

                    ExptRow = ExptRow + 1

            # A. Database is not blank, NewValue file upload
            else:

                # B. If IGDB is not blank
                if IGDB != '':

                    for row in result:
                        # print(row[0])
                        # Cache database into variable list
                        db_TId = row[0]
                        db_Name = row[1]
                        db_Meta = row[2]
                        db_Mode = row[3]
                        db_Genre = row[4]
                        db_Themes = row[5]
                        db_Series = row[6]
                        db_PP = row[7]
                        db_Franchises = row[8]
                        db_Engine = row[9]
                        db_AltName = row[10]
                        db_IGDB = row[11]
                        db_NewZoo = row[12]
                        db_Id = row[13]

                        # print(TId, Name, db_IGDB, IGDB, "2A")

                        # if NewZoo is not blank and IGDB or NewZoo match
                        if str(NewZoo) != '':
                            # print(Name, db_Name, LId, db_LId, "1")
                            # C. If IGDB is not blank and record found in database base onIGDB / If IGDB is not blank and record found based on IGDB and NewZoo - Update
                            if str(IGDB) == str(db_IGDB) or str(NewZoo) == str(db_NewZoo):
                                print(TId, Name, IGDB, db_IGDB,  3)
                                # print(Name, db_Name, LId, db_LId, "2")

                                IGDB = str(IGDB)

                                if TId in ('0', '', ' ', '$N/A'):
                                    TId = db_TId

                                if Name in ('0', '', ' ', '$N/A'):
                                    Name = db_Name

                                if Meta in ('0', '', ' ', '$N/A'):
                                    Meta = db_Meta

                                if Mode in ('0', '', ' ', '$N/A'):
                                    Mode = db_Mode

                                if Genre in ('0', '', ' ', '$N/A'):
                                    Genre = db_Genre

                                if Themes in ('0', '', ' ', '$N/A'):
                                    Themes = db_Themes

                                if Series in ('0', '', ' ', '$N/A'):
                                    Series = db_Series

                                if PP in ('0', '', ' ', '$N/A'):
                                    PP = db_PP

                                if Franchises in ('0', '', ' ', '$N/A'):
                                    Franchises = db_Franchises

                                if Engine in ('0', '', ' ', '$N/A'):
                                    Engine = db_Engine

                                if AltName in ('0', '', ' ', '$N/A'):
                                    AltName = db_AltName

                                if NewZoo in ('0', '', ' ', '$N/A'):
                                    NewZoo = db_NewZoo

                                itemFound = True

                                try:
                                    # print("Update " + Name + " 1")
                                    # print("3")
                                    cursor1 = conn.cursor()

                                    cursor1.execute(
                                        "UPDATE Marketing.dbo.GamesTitles SET TitleID=?, TitleName=?, Metascore=?, GameModes=?, "
                                        "Genre=?, Themes=?, Series=?, PlayerPerspectives=?, Franchises=?, GameEngine=?, AlternativeNames=?, "
                                        "NewZoo_Website=? WHERE ID=?"
                                        , TId, Name, Meta, Mode, Genre, Themes, Series, PP, Franchises, Engine, AltName, NewZoo, db_Id
                                    )

                                    conn.commit()

                                    cursor1.close()

                                except (pyodbc.Error, pyodbc.Warning) as err:
                                    print("Update Error on TitleName = " + str(Name))
                                    err = str(err)

                                    worksheet.write(ExptRow, 0, TId)
                                    worksheet.write(ExptRow, 1, Name)
                                    worksheet.write(ExptRow, 2, Meta)
                                    worksheet.write(ExptRow, 3, Mode)
                                    worksheet.write(ExptRow, 4, Genre)
                                    worksheet.write(ExptRow, 5, Themes)
                                    worksheet.write(ExptRow, 6, Series)
                                    worksheet.write(ExptRow, 7, PP)
                                    worksheet.write(ExptRow, 8, Franchises)
                                    worksheet.write(ExptRow, 9, Engine)
                                    worksheet.write(ExptRow, 10, AltName)
                                    worksheet.write(ExptRow, 11, IGDB)
                                    worksheet.write(ExptRow, 12, NewZoo)
                                    worksheet.write(ExptRow, 13, err)

                            else:
                                Ratio = fuzz.ratio(Name, db_Name)

                                if (Ratio >= 90):
                                    itemFound = True

                                    print(TId, Name, db_Name, IGDB, db_IGDB, Ratio, 4)

                                    worksheet1.write(ufRow, 0, TId)
                                    worksheet1.write(ufRow, 1, Name)
                                    worksheet1.write(ufRow, 2, Meta)
                                    worksheet1.write(ufRow, 3, Mode)
                                    worksheet1.write(ufRow, 4, Genre)
                                    worksheet1.write(ufRow, 5, Themes)
                                    worksheet1.write(ufRow, 6, Series)
                                    worksheet1.write(ufRow, 7, PP)
                                    worksheet1.write(ufRow, 8, Franchises)
                                    worksheet1.write(ufRow, 9, Engine)
                                    worksheet1.write(ufRow, 10, AltName)
                                    worksheet1.write(ufRow, 11, IGDB)
                                    worksheet1.write(ufRow, 12, NewZoo)
                                    # worksheet.write(ExptRow, 13, err)

                                    ufRow = ufRow + 1


                        # if NewZoo is blank and IGDB match
                        else:
                            if str(IGDB) == str(db_IGDB):
                                print(TId, Name, IGDB, db_IGDB,  5)
                                # print(Name, db_Name, LId, db_LId, "2")

                                IGDB = str(IGDB)

                                if TId in ('0', '', ' ', '$N/A'):
                                    TId = db_TId

                                if Name in ('0', '', ' ', '$N/A'):
                                    Name = db_Name

                                if Meta in ('0', '', ' ', '$N/A'):
                                    Meta = db_Meta

                                if Mode in ('0', '', ' ', '$N/A'):
                                    Mode = db_Mode

                                if Genre in ('0', '', ' ', '$N/A'):
                                    Genre = db_Genre

                                if Themes in ('0', '', ' ', '$N/A'):
                                    Themes = db_Themes

                                if Series in ('0', '', ' ', '$N/A'):
                                    Series = db_Series

                                if PP in ('0', '', ' ', '$N/A'):
                                    PP = db_PP

                                if Franchises in ('0', '', ' ', '$N/A'):
                                    Franchises = db_Franchises

                                if Engine in ('0', '', ' ', '$N/A'):
                                    Engine = db_Engine

                                if AltName in ('0', '', ' ', '$N/A'):
                                    AltName = db_AltName

                                if NewZoo in ('0', '', ' ', '$N/A'):
                                    NewZoo = db_NewZoo

                                itemFound = True

                                try:
                                    # print("Update " + Name + " 1")
                                    cursor1 = conn.cursor()

                                    cursor1.execute(
                                        "UPDATE Marketing.dbo.GamesTitles SET TitleID=?, TitleName=?, Metascore=?, GameModes=?, "
                                        "Genre=?, Themes=?, Series=?, PlayerPerspectives=?, Franchises=?, GameEngine=?, AlternativeNames=?, "
                                        "NewZoo_Website=? WHERE ID=?"
                                        , TId, Name, Meta, Mode, Genre, Themes, Series, PP, Franchises, Engine, AltName, NewZoo, db_Id
                                    )

                                    conn.commit()

                                    cursor1.close()

                                except (pyodbc.Error, pyodbc.Warning) as err:
                                    print("Update Error on TitleName = " + str(Name))
                                    err = str(err)

                                    worksheet.write(ExptRow, 0, TId)
                                    worksheet.write(ExptRow, 1, Name)
                                    worksheet.write(ExptRow, 2, Meta)
                                    worksheet.write(ExptRow, 3, Mode)
                                    worksheet.write(ExptRow, 4, Genre)
                                    worksheet.write(ExptRow, 5, Themes)
                                    worksheet.write(ExptRow, 6, Series)
                                    worksheet.write(ExptRow, 7, PP)
                                    worksheet.write(ExptRow, 8, Franchises)
                                    worksheet.write(ExptRow, 9, Engine)
                                    worksheet.write(ExptRow, 10, AltName)
                                    worksheet.write(ExptRow, 11, IGDB)
                                    worksheet.write(ExptRow, 12, NewZoo)
                                    worksheet.write(ExptRow, 13, err)

                                    '''
                                    if not itemFound:
                                        print("6")
                                if not itemFound:
                                    print("7")
                            if not itemFound:
                                print("8")
                        if not itemFound:
                            print("9")
                        '''

                            else:
                                Ratio = fuzz.ratio(Name, db_Name)

                                if (Ratio >= 90):
                                    itemFound = True

                                    print(TId, Name, db_Name, IGDB, db_IGDB, Ratio, 6)

                                    worksheet1.write(ufRow, 0, TId)
                                    worksheet1.write(ufRow, 1, Name)
                                    worksheet1.write(ufRow, 2, Meta)
                                    worksheet1.write(ufRow, 3, Mode)
                                    worksheet1.write(ufRow, 4, Genre)
                                    worksheet1.write(ufRow, 5, Themes)
                                    worksheet1.write(ufRow, 6, Series)
                                    worksheet1.write(ufRow, 7, PP)
                                    worksheet1.write(ufRow, 8, Franchises)
                                    worksheet1.write(ufRow, 9, Engine)
                                    worksheet1.write(ufRow, 10, AltName)
                                    worksheet1.write(ufRow, 11, IGDB)
                                    worksheet1.write(ufRow, 12, NewZoo)
                                    # worksheet.write(ExptRow, 13, err)

                                    ufRow = ufRow + 1

                    # If the record from excel is not found in database (Based on LinkedInId, CompanyName, Country and State)
                    if not itemFound:
                        # print(TId, Name, IGDB, db_IGDB, 4)
                        # print(TId, Name, db_Name)
                        # Ratio = fuzz.ratio(Name, db_Name)

                        # if (Ratio <= 90):
                            # print(LId, db_LId, Name, db_Name, "3")
                            print(TId, Name, db_Name, IGDB, db_IGDB,  Ratio, 7)
                            # print(Name, db_Name, "Fuzzy check")
                            # print(LId, Web, EmpRange, City, Region, Country, BizClass, BizSub, Active, Source, URL,
                            # Desc, Typ, Add, Ph, EmpLk, Found, SixMth, OneYr, TwoYr, db_LId)

                            TId = str(TId)

                            '''
                            if TId in ('0', '', ' ', '$N/A'):
                                TId = db_TId

                            if Name in ('0', '', ' ', '$N/A'):
                                Name = db_Name

                            if Meta in ('0', '', ' ', '$N/A'):
                                Meta = db_Meta

                            if Mode in ('0', '', ' ', '$N/A'):
                                Mode = db_Mode

                            if Genre in ('0', '', ' ', '$N/A'):
                                Genre = db_Genre

                            if Themes in ('0', '', ' ', '$N/A'):
                                Themes = db_Themes

                            if Series in ('0', '', ' ', '$N/A'):
                                Series = db_Series

                            if PP in ('0', '', ' ', '$N/A'):
                                PP = db_PP

                            if Franchises in ('0', '', ' ', '$N/A'):
                                Franchises = db_Franchises

                            if Engine in ('0', '', ' ', '$N/A'):
                                Engine = db_Engine

                            if AltName in ('0', '', ' ', '$N/A'):
                                AltName = db_AltName

                            if NewZoo in ('0', '', ' ', '$N/A'):
                                NewZoo = db_NewZoo
                            '''
                            # try:

                            # print(Name, LId, "3")

                            # print('2', VId, Name, Region, Country, EmpLk, 'Insert Record')
                            cursor4 = conn.cursor()

                            try:
                                # print("Insert " + Name + " 4 ")
                                cursor4.execute(
                                    "INSERT INTO GamesTitles(TitleID, TitleName, Metascore, GameModes, Genre, Themes, Series, PlayerPerspectives, "
                                    "Franchises, GameEngine, AlternativeNames, IGDB_Website, NewZoo_Website) "
                                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                                    , TId, Name, Meta, Mode, Genre, Themes, Series, PP, Franchises, Engine, AltName, IGDB, NewZoo
                                )

                                conn.commit()
                                cursor4.close()

                            except (pyodbc.Error, pyodbc.Warning) as err:
                                print("Insert Error on TitleName = " + str(Name))
                                err = str(err)

                                worksheet.write(ExptRow, 0, TId)
                                worksheet.write(ExptRow, 1, Name)
                                worksheet.write(ExptRow, 2, Meta)
                                worksheet.write(ExptRow, 3, Mode)
                                worksheet.write(ExptRow, 4, Genre)
                                worksheet.write(ExptRow, 5, Themes)
                                worksheet.write(ExptRow, 6, Series)
                                worksheet.write(ExptRow, 7, PP)
                                worksheet.write(ExptRow, 8, Franchises)
                                worksheet.write(ExptRow, 9, Engine)
                                worksheet.write(ExptRow, 10, AltName)
                                worksheet.write(ExptRow, 11, IGDB)
                                worksheet.write(ExptRow, 12, NewZoo)
                                worksheet.write(ExptRow, 13, err)

                                ExptRow = ExptRow + 1

                                # print(Web, LId, 'Record Updated', Stmt)
                                # print('5', LId, Name, Region, Country, 'Create Record')
                                # print(VId, now, now1, now2)

                                '''
                                else:
                                    # print("Insert Error on TitleName = " + str(Name))
                                    # err = str(err)
                                    print(TId, Name, db_Name, IGDB, db_IGDB, 6)
        
                                    worksheet1.write(ufRow, 0, TId)
                                    worksheet1.write(ufRow, 1, Name)
                                    worksheet1.write(ufRow, 2, Meta)
                                    worksheet1.write(ufRow, 3, Mode)
                                    worksheet1.write(ufRow, 4, Genre)
                                    worksheet1.write(ufRow, 5, Themes)
                                    worksheet1.write(ufRow, 6, Series)
                                    worksheet1.write(ufRow, 7, PP)
                                    worksheet1.write(ufRow, 8, Franchises)
                                    worksheet1.write(ufRow, 9, Engine)
                                    worksheet1.write(ufRow, 10, AltName)
                                    worksheet1.write(ufRow, 11, IGDB)
                                    worksheet1.write(ufRow, 12, NewZoo)
                                    # worksheet.write(ExptRow, 13, err)
        
                                    ufRow = ufRow + 1
                                '''

my_file = Path(dst)
if my_file.is_file():
    os.remove('%s' %(dst))
    shutil.copy(src, dst)
else:
    shutil.copy(src, dst)

workbook.close()
workbook1.close()

# if no exception, remove the file
if ExptRow == 1:
    os.remove('%s' %(xpt))

# if no unFound, remove the file
if ufRow == 1:
    os.remove('%s' %(uft))

conn.close()
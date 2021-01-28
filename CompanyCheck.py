'''
The objective of this script is to identify all the unidentify companies based on the following 3 columns Ultimate Parent, Parent and Subsidiaries.
Update the information if VTSID is found and insert as new record if the VTSID is not found.
'''

import openpyxl
import re
import pyodbc
import pandas as pd
import os.path
import numpy as np
from pathlib import Path
from configparser import ConfigParser
import shutil
from datetime import datetime
import xlsxwriter
from fuzzywuzzy import fuzz

# Read Config file
config = ConfigParser()
config.read('config.ini')

# pd.show_versions()

# SQL Database Connection
host = config['database']['host']
db = config['database']['db']
pwd = config['database']['user']
uid = config['database']['pass']
dir6 = config['File6']['dir6']
dir7 = config['File7']['dir7']
dir8 = config['File8']['dir8']

dt = datetime.today().strftime('%Y%m%d')

src = dir6
dst = dir7 + "_" + dt + ".xlsx"

xpt = dir8 + "_Exception" + "_" + dt + ".xlsx"

# print(src, dst)

conn = pyodbc.connect('driver={ODBC Driver 17 for SQL Server};SERVER='+host+';DATABASE='+db+';UID='+uid+';PWD='+pwd)

cursor = conn.cursor()
cursor.execute(
    'SELECT ID, LinkedInID, CompanyName, CompanyWebsite, EmployeeRange, City, RegionStateProvince, Country, BusinessClassification, BusinessSubclassification, Active, '
    'Source, LinkedInURL, [Description], [Type], CompanyAddress, Phone, EmployeesonLinkedIn, Founded, Growth6mth, Growth1yr, Growth2yr FROM Marketing.dbo.Marketing_ETL'
)

result = cursor.fetchall()

cursor.close()
# conn.close()

# Create ExceptionCatch file and Remove file if exists
my_file = Path('%s' %(xpt))
if my_file.is_file():
    os.remove('%s' %(xpt))
else:
    # print("The file does not exist")
    print("No Exception File found ....")

print("Creating M&A Exception File ....")
df = pd.read_excel('%s' %(dir6), sheet_name='New', engine='openpyxl')
df = df.replace(np.nan, ' ', regex=True)
df_Header = df.columns.ravel()

workbook = xlsxwriter.Workbook('%s' %(xpt))
worksheet = workbook.add_worksheet()

for col_num, data in enumerate(df_Header):
    # print(data)
    if col_num <= 28:
        worksheet.write(0, col_num, data)

worksheet.write(0, 29, 'Error')

ExptRow = 1

# Check for the Unidentify Company based on Ultimate Parent, Parent and Subsidiaries

# row Initialisation for the Unidentify Company based on Ultimate Parent, Parent and Subsidiaries
xlsRow = 1

# for ir in range(10, 13):
for ir in range(0, len(df)):
    for ic in range(0, len(df.columns)):
        ItemFound = False

        if ic == 0:

            if df.iat[ir, 0] not in ('0', '0.0', '#N/A', '', ' '):
                Id = int(df.iat[ir, 0])
            else:
                Id = 0

            if df.iat[ir, 2] not in ('0', '0.0', '#N/A', '', ' '):
                LId = int(df.iat[ir, 2])
            else:
                LId = str('')

            Name = str(df.iat[ir, 3])
            if Name in ('0', '0.0', '', ' ', '$N/A'):
                Name = ''

            Web = str(df.iat[ir, 7])
            if Web in ('0', '0.0', '', ' ', '$N/A'):
                Web = ''

            EmpRange = str(df.iat[ir, 8])
            if EmpRange in ('0', '0.0', '', ' ', '$N/A'):
                EmpRange = ''

            UParent = str(df.iat[ir, 9])
            if UParent in ('0', '0.0', '', ' ', '$N/A'):
                UParent = ''

            Parent = str(df.iat[ir, 10])
            if Parent in ('0', '0.0', '', ' ', '$N/A'):
                Parent = ''

            Sub = str(df.iat[ir, 11])
            if Sub in ('0', '0.0', '', ' ', '$N/A'):
                Sub = ''

            City = str(df.iat[ir, 12])
            if City in ('0', '0.0', '', ' ', '$N/A'):
                City = ''

            Region = str(df.iat[ir, 13])
            if Region in ('0', '0.0', '', ' ', '$N/A'):
                Region = ''

            Country = str(df.iat[ir, 14])
            if Country in ('0', '0.0', '', ' ', '$N/A'):
                Country = ''

            BizClass = str(df.iat[ir, 15])
            if BizClass in ('0', '0.0', '', ' ', '$N/A'):
                BizClass = ''

            BizSub = str(df.iat[ir, 16])
            if BizSub in ('0', '0.0', '', ' ', '$N/A'):
                BizSub = ''

            Active = str(df.iat[ir, 17])
            if Active in ('0', '0.0', '', ' ', '$N/A'):
                Active = ''

            Source = str(df.iat[ir, 18])
            if Source in ('0', '0.0', '', ' ', '$N/A'):
                Source = ''

            URL = str(df.iat[ir, 19])
            if URL in ('0', '0.0', '', ' ', '$N/A'):
                URL = ''

            Desc = str(df.iat[ir, 20])
            if Desc in ('0', '0.0', '', ' ', '$N/A'):
                Desc = ''

            Typ = str(df.iat[ir, 21])
            if Typ in ('0', '0.0', '', ' ', '$N/A'):
                Typ = ''

            Add = str(df.iat[ir, 22])
            if Add in ('0', '0.0', '', ' ', '$N/A'):
                Add = ''

            Ph = str(df.iat[ir, 23])
            if Ph in (0, '0', '0.0', '#N/A', '', ' '):
                Ph = ''

            EmpLk = str(df.iat[ir, 24])
            if EmpLk in (0, '0', '0.0', '#N/A', '', ' '):
                EmpLk = ''
            # else:
            #   EmpLk = (EmpLk[:250] + '..') if len(EmpLk) > 250 else EmpLk

            Found = str(df.iat[ir, 25])
            if Found in (0, '0', '0.0', '#N/A', '', ' '):
                Found = ''

            # if df.iat[ir1, 25] not in (0, '0', '0.0', '#N/A', '', ' '):
            #    Found = int(df.iat[ir1, 25])
            # else:
            #     Found = str('')

            SixMth = df.iat[ir, 26]
            if str(SixMth) in ('0', '0.0', '', ' ', '$N/A'):
                SixMth = 0

            OneYr = df.iat[ir, 27]
            if str(OneYr) in ('0', '0.0', '', ' ', '$N/A'):
                OneYr = 0

            TwoYr = df.iat[ir, 28]
            if str(TwoYr) in ('0', '0.0', '', ' ', '$N/A'):
                TwoYr = 0

            for row in result:
                # print(row[0])
                # Cache database into variable list
                db_Id = row[0]
                db_LId = row[1]
                db_Name = row[2]
                db_Web = row[3]
                db_EmpRange = row[4]
                db_City = row[5]
                db_Region = row[6]
                db_Country = row[7]
                db_BizClass = row[8]
                db_BizSub = row[9]
                db_Active = row[10]
                db_Source = row[11]
                db_URL = row[12]
                db_Desc = row[13]
                db_Typ = row[14]
                db_Add = row[15]
                db_Ph = row[16]
                db_EmpLk = row[17]
                db_Found = row[18]
                db_SixMth = row[19]
                db_OneYr = row[20]
                db_TwoYr = row[21]
                # db_UParent = row[18]
                # db_Parent = row[19]
                # db_Sub = row[20]

                if str(Id) not in ('', ' ', '0'):
                    # First check - Id

                    if Id == db_Id:
                        # print('3', LId, db_LId)

                        LId = str(LId)
                        if LId in ('0', '', ' ', '$N/A'):
                            LId = db_LId

                        if Name in ('0', '', ' ', '$N/A'):
                            Name = db_Name

                        if Web in ('0', '', ' ', '$N/A'):
                            Web = db_Web

                        if EmpRange in ('0', '', ' ', '$N/A'):
                            EmpRange = db_EmpRange

                        if City in ('0', '', ' ', '$N/A'):
                            City = db_City

                        if Region in ('0', '', ' ', '$N/A'):
                            Region = db_Region

                        if Country in ('0', '', ' ', '$N/A'):
                            Country = db_Country

                        if BizClass in ('0', '', ' ', '$N/A'):
                            BizClass = db_BizClass

                        if BizSub in ('0', '', ' ', '$N/A'):
                            BizSub = db_BizSub

                        if Active in ('0', '', ' ', '$N/A'):
                            Active = db_Active

                        if Source in ('0', '', ' ', '$N/A'):
                            Source = db_Source

                        if URL in ('0', '', ' ', '$N/A'):
                            URL = db_URL

                        if Desc in ('0', '', ' ', '$N/A'):
                            Desc = db_Desc

                        if Typ in ('0', '', ' ', '$N/A'):
                            Typ = db_Typ

                        if Add in ('0', '', ' ', '$N/A'):
                            Add = db_Add

                        if Ph in ('0', '', ' ', '$N/A'):
                            Ph = db_Ph

                        if EmpLk in ('0', '', ' ', '$N/A'):
                            EmpLk = db_EmpLk

                        if Found in ('0', '', ' ', '$N/A'):
                            Found = db_EmpLk

                        if SixMth in (0, '0', '', ' ', '$N/A'):
                            SixMth = db_SixMth

                        if OneYr in (0, '0', '', ' ', '$N/A'):
                            OneYr = db_OneYr

                        if TwoYr in (0, '0', '', ' ', '$N/A'):
                            TwoYr = db_TwoYr

                        ItemFound = True

                        cursor1 = conn.cursor()

                        try:

                            # print(db_Id, "Updated", "1")
                            cursor1.execute(
                                "UPDATE Marketing.dbo.Marketing_ETL SET LinkedInId=?, CompanyName=?, CompanyWebsite=?, EmployeeRange=?, UltimateParent=?, Parent=?, "
                                "Subsidiaries=?, City=?, RegionStateProvince=?, Country=?, BusinessClassification=?, BusinessSubclassification=?, Active=?, Source=?, "
                                "LinkedInURL=?, [Description]=?, [Type]=?, CompanyAddress=?, Phone=?, EmployeesonLinkedIn=?, Founded=?, Growth6mth=?, Growth1yr=?, "
                                "Growth2yr=? WHERE ID= ?"
                                , LId, Name, Web, EmpRange, UParent, Parent, Sub, City, Region, Country, BizClass, BizSub, Active, Source, URL, Desc, Typ, Add, Ph, EmpLk,
                                Found, SixMth, OneYr, TwoYr, Id
                            )

                            conn.commit()
                            # print("Update Complete")

                            cursor1.close

                        except (pyodbc.Error, pyodbc.Warning) as err:
                            # print("Update Error on ID = " + str(Id))
                            err = str(err)
                            print(err)

            if not ItemFound:
                if str(Id) == '':
                    # not in ('0', '', ' ', '$N/A'):
                    print("ID cannot be empty")
                    err = "ID cannot be empty"

                else:
                    print(str(Id) + " Not Found in database")
                    err = str(Id) + " Not Found in database"

                worksheet.write(ExptRow, 0, Id)
                worksheet.write(ExptRow, 2, LId)
                worksheet.write(ExptRow, 3, Name)
                worksheet.write(ExptRow, 7, Web)
                worksheet.write(ExptRow, 8, EmpRange)
                worksheet.write(ExptRow, 9, UParent)
                worksheet.write(ExptRow, 10, Parent)
                worksheet.write(ExptRow, 11, Sub)
                worksheet.write(ExptRow, 12, City)
                worksheet.write(ExptRow, 13, Region)
                worksheet.write(ExptRow, 14, Country)
                worksheet.write(ExptRow, 15, BizClass)
                worksheet.write(ExptRow, 16, BizSub)
                worksheet.write(ExptRow, 17, Active)
                worksheet.write(ExptRow, 18, Source)
                worksheet.write(ExptRow, 19, URL)
                worksheet.write(ExptRow, 20, Desc)
                worksheet.write(ExptRow, 21, Typ)
                worksheet.write(ExptRow, 22, Add)
                worksheet.write(ExptRow, 23, Ph)
                worksheet.write(ExptRow, 24, EmpLk)
                worksheet.write(ExptRow, 25, Found)
                worksheet.write(ExptRow, 26, SixMth)
                worksheet.write(ExptRow, 27, OneYr)
                worksheet.write(ExptRow, 28, TwoYr)
                worksheet.write(ExptRow, 29, err)

                ExptRow = ExptRow + 1

my_file = Path(dst)
if my_file.is_file():
    os.remove('%s' %(dst))
    shutil.copy(src, dst)
else:
    shutil.copy(src, dst)

workbook.close()

if ExptRow == 1:
    os.remove('%s' % (xpt))

conn.close()

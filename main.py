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
src = dir1
dst = dir4 + "_" + dt + ".xlsx"

# For VendorUpdate_NewValue
# src = dir2
# dst = dir5 + "_" + dt + ".xlsx"

xpt = dir3 + "_Exception" + "_" + dt + ".xlsx"

conn = pyodbc.connect('driver={ODBC Driver 17 for SQL Server};SERVER='+host+';DATABASE='+db+';UID='+uid+';PWD='+pwd)

cursor = conn.cursor()
cursor.execute(
    'SELECT ID, LinkedInID, CompanyName, CompanyWebsite, EmployeeRange, City, RegionStateProvince, Country, BusinessClassification, BusinessSubclassification, Active, Source, '
    'LinkedInURL, [Description], [Type], CompanyAddress, Phone, EmployeesonLinkedIn, Founded, Growth6mth, Growth1yr, Growth2yr, VTSID FROM Marketing.dbo.Marketing_ETL_TEST ORDER BY ID'
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

# For first upload of MasterList
df = pd.read_excel('%s' %(dir1), sheet_name='New', engine='openpyxl')
# For VendorUpdate_NewValue
# df = pd.read_excel('%s' %(dir2), sheet_name='New', engine='openpyxl')

df = df.replace(np.nan, ' ', regex=True)
df_Header = df.columns.ravel()

workbook = xlsxwriter.Workbook('%s' %(xpt))
worksheet = workbook.add_worksheet()

for col_num, data in enumerate(df_Header):
    # print(data)
    if col_num <= 28:
        worksheet.write(0, col_num, data)

worksheet.write(0, 29, 'Error')

# workbook.close()

ExptRow = 1

# print(len(df))
# print(len(df.columns))

def InsertRecord(VId, LId, Name, Dev, SDev, PDev, Web, EmpRange, City, Region, Country, BizClass, BizSub, Active, Source, URL, Desc, Typ, Add, Ph, EmpLk, Found, SixMth, OneYr, TwoYr, ExptRow):

    cursor = conn.cursor()

    try:
        # print("First Upload..... ")
        cursor.execute(
            "INSERT INTO Marketing.dbo.Marketing_ETL_TEST(VTSID, LinkedInID, CompanyName, Developer, SupportingDeveloper, PortingDeveloper, CompanyWebsite, EmployeeRange, "
            "UltimateParent, Parent, Subsidiaries, City, RegionStateProvince, Country, BusinessClassification, BusinessSubclassification, Active, [Source], LinkedInURL, "
            "[Description], [Type], CompanyAddress, Phone, EmployeesonLinkedIn, Founded, Growth6mth, Growth1yr, Growth2yr) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            , VId, LId, Name, Dev, SDev, PDev, Web, EmpRange, '', '', '', City, Region, Country, BizClass, BizSub, Active, Source, URL, Desc, Typ, Add, Ph, EmpLk, Found, SixMth, 
            OneYr, TwoYr
        )

        conn.commit()
        cursor.close()

    except (pyodbc.Error, pyodbc.Warning) as err:
        print("Insert Error on Company = " + str(Name))
        err = str(err)

        worksheet.write(ExptRow, 0, VId)
        worksheet.write(ExptRow, 2, LId)
        worksheet.write(ExptRow, 3, Name)
        worksheet.write(ExptRow, 4, Dev)
        worksheet.write(ExptRow, 5, SDev)
        worksheet.write(ExptRow, 6, PDev)
        worksheet.write(ExptRow, 7, Web)
        worksheet.write(ExptRow, 8, EmpRange)
        # worksheet.write(ExptRow, 9, UParent)
        # worksheet.write(ExptRow, 10, Parent)
        # worksheet.write(ExptRow, 11, Sub)
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

        return ExptRow


def UpdateRecord(VId, LId, Name, Dev, SDev, PDev, Web, EmpRange, City, Region, Country, BizClass, BizSub, Active, Source, URL, Desc, Typ, Add, Ph, EmpLk, Found, SixMth, OneYr, TwoYr, db_Id):

    cursor = conn.cursor()

    try:
        cursor.execute(
            "UPDATE Marketing.dbo.Marketing_ETL_TEST SET LinkedInID=?, CompanyWebsite=?, EmployeeRange=?, City=?, RegionStateProvince=?, Country=?, BusinessClassification=?, "
            "BusinessSubclassification=?, Active=?, Source=?, LinkedInURL=?, [Description]=?, [Type]=?, CompanyAddress=?, Phone=?, EmployeesonLinkedIn=?, Founded=?, Growth6mth=?, "
            "Growth1yr=?, Growth2yr= ? WHERE ID=?"
            , LId, Web, EmpRange, City, Region, Country, BizClass, BizSub, Active, Source, URL, Desc, Typ, Add, Ph, EmpLk, Found, SixMth, OneYr, TwoYr, db_Id
        )

        conn.commit()

        cursor.close()

    except (pyodbc.Error, pyodbc.Warning) as err:
        print("Update Error on Company = " + str(Name))
        err = str(err)

        worksheet.write(ExptRow, 0, VId)
        worksheet.write(ExptRow, 2, LId)
        worksheet.write(ExptRow, 3, Name)
        worksheet.write(ExptRow, 4, Dev)
        worksheet.write(ExptRow, 5, SDev)
        worksheet.write(ExptRow, 6, PDev)
        worksheet.write(ExptRow, 7, Web)
        worksheet.write(ExptRow, 8, EmpRange)
        # worksheet.write(ExptRow, 9, UParent)
        # worksheet.write(ExptRow, 10, Parent)
        # worksheet.write(ExptRow, 11, Sub)
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

# Check for identical company
for ir1 in range(0, len(df)):
# for ir1 in range(8, 9):
    for ic1 in range(0, len(df.columns)):
        # print(ir1, ic1)

        if (ic1 == 2):
            itemFound = False
            # itemFound1 = False

            # Id = int(df.iat[ir1, 0])

            if str(df.iat[ir1, 0]) not in (0, '0', '0.0', '#N/A', '', ' '):
                VId = str(df.iat[ir1, 0])
            else:
                VId = str('')

            if str(df.iat[ir1, 2]) not in ('0', '0.0', '#N/A', '', ' '):
                LId = int(df.iat[ir1, 2])
            else:
                LId = str('')

            Name = str(df.iat[ir1, 3])
            if Name in ('0', '0.0', '', ' ', '$N/A'):
                Name = ''

            Dev = str(df.iat[ir1, 4])
            if Dev in ('0', '0.0', '', ' ', '$N/A'):
                Dev = ''

            SDev = str(df.iat[ir1, 5])
            if SDev in ('0', '0.0', '', ' ', '$N/A'):
                SDev = ''

            PDev = str(df.iat[ir1, 6])
            if PDev in ('0', '0.0', '', ' ', '$N/A'):
                PDev = ''

            Web = str(df.iat[ir1, 7])
            if Web in ('0', '0.0', '', ' ', '$N/A'):
                Web = ''

            EmpRange = str(df.iat[ir1, 8])
            if EmpRange in ('0', '0.0', '', ' ', '$N/A'):
                EmpRange = ''

            '''    
            UParent = str(df.iat[ir1, 9])
            if UParent in ('0', '0.0', '', ' ', '$N/A'):
                UParent = ''
    
            Parent = str(df.iat[ir1, 10])
            if Parent in ('0', '0.0', '', ' ', '$N/A'):
                Parent = ''
    
            Sub = str(df.iat[ir1, 11])
            if Sub in ('0', '0.0', '', ' ', '$N/A'):
                Sub = ''
            '''

            City = str(df.iat[ir1, 12])
            if City in ('0', '0.0', '', ' ', '$N/A'):
                City = ''

            Region = str(df.iat[ir1, 13])
            if Region in ('0', '0.0', '', ' ', '$N/A'):
                Region = ''

            Country = str(df.iat[ir1, 14])
            if Country in ('0', '0.0', '', ' ', '$N/A'):
                Country = ''

            BizClass = str(df.iat[ir1, 15])
            if BizClass in ('0', '0.0', '', ' ', '$N/A'):
                BizClass = ''

            BizSub = str(df.iat[ir1, 16])
            if BizSub in ('0', '0.0', '', ' ', '$N/A'):
                BizSub = ''

            Active = str(df.iat[ir1, 17])
            if Active in ('0', '0.0', '', ' ', '$N/A'):
                Active = ''

            Source = str(df.iat[ir1, 18])
            if Source in ('0', '0.0', '', ' ', '$N/A'):
                Source = ''

            URL = str(df.iat[ir1, 19])
            if URL in ('0', '0.0', '', ' ', '$N/A'):
                URL = ''

            Desc = str(df.iat[ir1, 20])
            if Desc in ('0', '0.0', '', ' ', '$N/A'):
                Desc = ''

            Typ = str(df.iat[ir1, 21])
            if Typ in ('0', '0.0', '', ' ', '$N/A'):
                Typ = ''

            Add = str(df.iat[ir1, 22])
            if Add in ('0', '0.0', '', ' ', '$N/A'):
                Add = ''

            Ph = str(df.iat[ir1, 23])
            if Ph in (0, '0', '0.0', '#N/A', '', ' '):
                Ph = ''

            '''
            EmpLk = str(df.iat[ir1, 24])
            if EmpLk in (0, '0', '0.0', '#N/A', '', ' '):
                EmpLk = ''
            '''

            if str(df.iat[ir1, 24]) not in ('0', '0.0', '#N/A', '', ' '):
                EmpLk = str(df.iat[ir1, 24])
            else:
                EmpLk = str('')

            '''
            Found = str(df.iat[ir1, 25])
            if Found in (0, '0', '0.0', '#N/A', '', ' '):
                Found = ''
            '''

            if str(df.iat[ir1, 25]) not in (0, '0', '0.0', '#N/A', '', ' '):
                Found = str(df.iat[ir1, 25])
            else:
                Found = str('')

            SixMth = df.iat[ir1, 26]
            if str(SixMth) in ('0', '0.0', '', ' ', '$N/A'):
                SixMth = 0

            OneYr = df.iat[ir1, 27]
            if str(OneYr) in ('0', '0.0', '', ' ', '$N/A'):
                OneYr = 0

            TwoYr = df.iat[ir1, 28]
            if str(TwoYr) in ('0', '0.0', '', ' ', '$N/A'):
                TwoYr = 0

            '''
            VId = df.iat[ir1, 0]
            if str(VId) in (0, '0', '0.0', '#N/A', '', ' '):
                VId = str('')
            '''

            # A. If database is blank, it is an initial upload
            if len(result) == 0:

                InsertRecord(VId, LId, Name, Dev, SDev, PDev, Web, EmpRange, City, Region, Country, BizClass, BizSub, Active, Source, URL, Desc, Typ, Add, Ph, EmpLk, Found, SixMth, OneYr, TwoYr, ExptRow)

            # A. Database is not blank, NewValue file upload
            else:

                # print(LId, Name)
                # B. If LinkedInId is not blank
                if LId != '':
                    for row in result:
                        # print(row[0])
                        # Cache database into variable list
                        db_Id = row[0]
                        db_LId = row[1]
                        db_Name = row[2]
                        db_Web = row[3]
                        db_EmpRange = row[4]
                        # db_UParent = row[4]
                        # db_Parent = row[5]
                        # db_Sub = row[6]
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
                        db_EmpLk = (row[17])
                        db_Found = (row[18])
                        db_SixMth = row[19]
                        db_OneYr = row[20]
                        db_TwoYr = row[21]
                        db_VId = row[22]

                        # print(Name, db_Name, LId, db_LId, "1")
                        # C. If LinkedInId is not blank and record found in database - Update
                        if str(LId) == str(db_LId):

                            # print(Name, db_Name, LId, db_LId, "2")

                            LId = str(LId)

                            if Name in ('0', '', ' ', '$N/A'):
                                Name = db_Name

                            if Web in ('0', '', ' ', '$N/A'):
                                Web = db_Web

                            if EmpRange in ('0', '', ' ', '$N/A'):
                                EmpRange = db_EmpRange

                            '''
                            if UParent in ('0', '', ' ', '$N/A'):
                                UParent = db_UParent

                            if Parent in ('0', '', ' ', '$N/A'):
                                Parent = db_Parent

                            if Sub in ('0', '', ' ', '$N/A'):
                                Sub = db_Sub
                            '''

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

                            if VId in (0, '0', '', ' ', '$N/A'):
                                VId = db_VId

                            itemFound = True

                            UpdateRecord(VId, LId, Name, Dev, SDev, PDev, Web, EmpRange, City, Region, Country, BizClass, BizSub, Active, Source, URL, Desc, Typ, Add, Ph, EmpLk, Found, SixMth, OneYr, TwoYr, db_Id)

                        # C. If LinkedInId is not blank and fuzzy check found in database - Update
                        else:

                            Ratio = fuzz.ratio(Name, db_Name)

                            if (Ratio >= 80) and (Region == db_Region) and (Country == db_Country):
                                # print(LId, db_LId, Name, db_Name, "3")

                                LId = str(LId)

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

                                if VId in (0, '0', '', ' ', '$N/A'):
                                    VId = db_VId

                                itemFound = True

                                UpdateRecord(VId, LId, Name, Dev, SDev, PDev, Web, EmpRange, City, Region, Country, BizClass, BizSub, Active, Source, URL, Desc, Typ, Add, Ph, EmpLk, Found, SixMth, OneYr, TwoYr, db_Id)

                # B. If LinkedId is blank and fuzzy check found in database - Update
                else:

                    for row in result:
                        # print(row[0])
                        # Cache database into variable list
                        db_Id = row[0]
                        db_LId = row[1]
                        db_Name = row[2]
                        db_Web = row[3]
                        db_EmpRange = row[4]
                        # db_UParent = row[4]
                        # db_Parent = row[5]
                        # db_Sub = row[6]
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
                        db_EmpLk = (row[17])
                        db_Found = (row[18])
                        db_SixMth = row[19]
                        db_OneYr = row[20]
                        db_TwoYr = row[21]
                        db_VId = row[22]

                        Ratio = fuzz.ratio(Name, db_Name)

                        if (Ratio >= 80) and (Region == db_Region) and (Country == db_Country):

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

                            if VId in (0, '0', '', ' ', '$N/A'):
                                VId = db_VId

                            itemFound = True

                            UpdateRecord(VId, LId, Name, Dev, SDev, PDev, Web, EmpRange, City, Region, Country, BizClass, BizSub, Active, Source, URL, Desc, Typ, Add, Ph, EmpLk, Found, SixMth, OneYr, TwoYr, db_Id)

                        # if not itemFound1:
                            # print(Name, "")
                    # if not itemFound1:
                        # print(Name, "ItemFound1")
                # if not itemFound:
                    # print(Name, "ItemFound")

                # If the record from excel is not found in database (Based on LinkedInId, CompanyName, Country and State
                if not itemFound:

                    InsertRecord(VId, LId, Name, Dev, SDev, PDev, Web, EmpRange, City, Region, Country, BizClass, BizSub, Active, Source, URL, Desc, Typ, Add, Ph, EmpLk, Found, SixMth, OneYr, TwoYr, ExptRow)

my_file = Path(dst)
if my_file.is_file():
    os.remove('%s' %(dst))
    shutil.copy(src, dst)
else:
    shutil.copy(src, dst)

workbook.close()

# if no exception, remove the file
if ExptRow == 1:
    os.remove('%s' %(xpt))

conn.close()
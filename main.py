import os

import numpy as np
import openpyxl
import pandas as pd
import os.path
import shutil
import logging
import dateutil.parser

from datetime import datetime
from pathlib import Path

from openpyxl.styles import Border, Side
# from Utils import Utils
from os import path
import shutil
from shutil import make_archive
from shutil import copyfile


# Press the green button in the gutter to run the script.
def CreateFolders(folderName, __rootPath):
    folderList = ['in', 'out', 'log', 'company', 'toSend_Booster_Signup_Conso']

    folderName = os.path.join(__rootPath, folderName)

    for folder in folderList:

        if not path.exists(folderName):
            os.makedirs(folderName)

        folderPath = os.path.join(folderName, folder)

        # Check if folder exist
        if not path.exists(folderPath):
            os.makedirs(folderPath)

    return folderName


def mergingColumn(df, range, columnName):
    df_result = df.iloc[:, range].copy()
    # print(df_result.columns.values.tolist())

    df[columnName] = df_result[
        df_result.columns.values.tolist()].apply(
        lambda x: get_value_from_listRow(x), axis=1)
    pass


def getData(inFile):
    filePath = os.path.join(inPath, inFile)
    df = pd.read_excel(filePath, dtype=str, na_filter=False)

    company = inFile.split('_COVID')[0]

    withSubsColumn_arr = ['Philippine Airlines, Inc. and PAL Express Employees',
                          'Philippine National Bank and Subsidiaries Employees',
                          'PMFTC Inc. Employees',
                          'Tanduay Distillers, Inc. and Subsidiaries Employees',
                          'MacroAsia Corp., Subsidiaries & Affiliates Employees']

    scenarioList = ['Where did you get your initial doses of COVID-19 vaccines?',
                    'Do you have prior registration with eZConsult via LTGC?',
                    'Enter your eZConsult Patient ID number.'
                    ]

    # Insert New Column if files have no addition subs
    if not company in withSubsColumn_arr:
        df.insert(7, 'What is the name of your employer?', np.nan, True)

    mergingColumn(df, [16, 17], 'new_eZConsult_Patient_ID')
    mergingColumn(df, [18, 28], 'new_vaccine_brand')
    mergingColumn(df, [19, 29, 25, 35], 'new_FirstDose_Date')
    mergingColumn(df, [20, 30, 26, 36], 'new_FirstDose_BatchNo')
    mergingColumn(df, [21, 31, 27, 37], 'new_FirstDose_VacSite')

    mergingColumn(df, [22, 32], 'new_SecondDose_Date')
    mergingColumn(df, [23, 33], 'new_SecondDose_BatchNo')
    mergingColumn(df, [24, 34], 'new_SecondDose_VacSite')

    df_result_province = df.iloc[:, 45:62].copy()
    df['new_province'] = df_result_province[df_result_province.columns.values.tolist()].apply(
        lambda x: get_value_from_listRow(x), axis=1)

    df_result_city = df.iloc[:, 62:149].copy()
    df['new_city'] = df_result_city[df_result_city.columns.values.tolist()].apply(lambda x: get_value_from_listRow(x),
                                                                                  axis=1)
    df.drop(df.iloc[:, 45:149], axis=1, inplace=True)
    df.drop(df.iloc[:, 16:38], axis=1, inplace=True)

    df['Company'] = company

    column_names = ['ID', 'Start time', 'Completion time', 'Email', 'Name',
                    'By filling out this form, I agree to participate in the booster vaccination program sponsored by LT Group, Inc., its subsidiaries and affiliates',
                    'Consent for Personal Data Collection', 'What is the name of your employer?', 'Employee Number',
                    'Last Name',
                    'First Name', 'Middle Name', 'Suffix', 'Birthdate',
                    'Where did you get your initial doses of COVID-19 vaccines?',
                    'Do you have prior registration with eZConsult via LTGC?', 'new_eZConsult_Patient_ID',
                    'new_vaccine_brand',
                    'new_FirstDose_Date', 'new_FirstDose_BatchNo', 'new_FirstDose_VacSite', 'new_SecondDose_Date',
                    'new_SecondDose_BatchNo', 'new_SecondDose_VacSite', 'Select your appropriate Priority Group',
                    'Are you a member of an indigenous group?', 'Are you a PWD (Person with Disability)?', 'Occupation',
                    'Mobile Number', 'Email Address', 'Region', 'new_province', 'new_city', 'Barangay', 'Sex',
                    'Allergy to vaccines or components of vaccines?',
                    'With Comorbidity/ies?', 'Please specify comorbidity/ies',
                    'Select preferred vaccination site for your booster shot.',
                    'Company']

    df = df.reindex(columns=column_names)
    df.rename(columns={
        "new_eZConsult_Patient_ID": "Enter your eZConsult Patient ID number.",
        "new_vaccine_brand": "Select the vaccine brand.",
        "new_FirstDose_Date": "Date of first dose",
        "new_FirstDose_BatchNo": "Batch No. / Lot No. of first dose",
        "new_FirstDose_VacSite": "Place of vaccination of first dose",
        "new_SecondDose_Date": "Date of second dose",
        "new_SecondDose_BatchNo": "Batch No. / Lot No. of second dose",
        "new_SecondDose_VacSite": "Place of vaccination of second dose",
        "new_province": "Province",
        "new_city": "City",
    }, inplace=True, errors="raise")

    if len(df):
        df['Scenario'] = df[scenarioList].apply(lambda x: myScenario(x), axis=1)


    df['Date of first dose'] = df.apply(lambda x: convertDateFormat(x.name, x['Date of first dose']),
                                        axis=1)
    df['Date of second dose'] = df.apply(lambda x: convertDateFormat(x.name, x['Date of second dose']),
                                         axis=1)

    df['Date of first dose'] = pd.to_datetime(df['Date of first dose'], errors='coerce')
    df['Date of second dose'] = pd.to_datetime(df['Date of second dose'], errors='coerce')

    df['Vaccine_days_gap'] = (df['Date of second dose'] - df['Date of first dose']).dt.days

    df['concats'] = df['Last Name'] + df['First Name'] + df['Middle Name'] + df['Employee Number']
    df['concats'] = df['concats'].str.replace(' ', '')

    df.drop_duplicates(subset=['concats'], keep='last', inplace=True)

    # Drop extra column
    df.drop(['concats', 'Vaccine_days_gap'],
            axis='columns', inplace=True)


    # df['is_GreaterEqual_date'] = df[['Date of first dose', 'Date of second dose']].apply(
    #     lambda x: isGreaterEqualDate(x),
    #     axis=1)
    #
    # df['is28_gap'] = df[['Date of first dose', 'Date of second dose', 'Select the vaccine brand.']].apply(lambda x: is28Gap(x),
    #                                                                          axis=1)

    # print(df)
    # # get number of columns exist
    # print(len(df.columns))

    return df


def convertDateFormat(name, param):
    if pd.notna(param):
        return dateutil.parser.parse(str(param)).strftime("%Y-%m-%d")
    else:
        return param


def get_value_from_listRow(row):
    row = list(filter(None, row))

    if row:
        return row[0]

    return


def myScenario(x):
    if x['Where did you get your initial doses of COVID-19 vaccines?'] == 'Lucio Tan Group of Companies (LTGC)':
        return 1
    elif x['Where did you get your initial doses of COVID-19 vaccines?'] == 'Elsewhere' and x[
        'Do you have prior registration with eZConsult via LTGC?'] == 'Yes':
        return 2
    elif x['Where did you get your initial doses of COVID-19 vaccines?'] == 'Elsewhere' and x[
        'Do you have prior registration with eZConsult via LTGC?'] == 'No':
        return 3

    pass


def duplicateTemplateLTGC(tempLTGC_Path, out, outputFilename):
    companyDir = out + "/"
    srcFile = companyDir + outputFilename + ".xlsx"

    if not os.path.isfile(srcFile):
        shutil.copy(tempLTGC_Path, srcFile)

    return companyDir + outputFilename + ".xlsx"
    pass


def writeExcel(df, templateFilePath, outPath, outFilename):
    # Create copy of template file and save it to out folder
    templateFile = duplicateTemplateLTGC(templateFilePath, outPath, outFilename)

    # Write df_master(consolidated/append data) to excel
    writer = pd.ExcelWriter(templateFile, engine='openpyxl', mode='a')
    writer.book = openpyxl.load_workbook(templateFile)
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    df.to_excel(writer, sheet_name="Form1", startrow=1, header=False, index=False)

    writer.save()

    # print(templateFile)

    pass


def Generate_File(comp_name, df, param):
    # print(len(df))
    folderName = os.path.join(compPath, comp_name)

    if not path.exists(folderName):
        os.makedirs(folderName)

    if param == 'a':
        templateFile = templateFileA
        Company_out_filename = comp_name + '_FileA_' + dateTime
    elif param == 'b':
        templateFile = templateFileB
        Company_out_filename = comp_name + '_FileB_' + dateTime

        df = df.loc[df['Scenario'] == 3].copy()

        # df.drop('Scenario', inplace=True, axis=1)
        df.drop('Scenario',
                axis='columns', inplace=True, errors='raise')

    elif param == 'n':
        templateFile = templateFilePath
        Company_out_filename = comp_name + '_' + dateTime

    writeExcel(df, templateFile, folderName, Company_out_filename)

    pass


def Generate_File_per_Company(master_df):
    groups = master_df.groupby('Company')
    fileA_column_arr = ['ID', 'Start time', 'Completion time', 'Email', 'Name',
                        'By filling out this form, I agree to participate in the booster vaccination program sponsored by LT Group, Inc., its subsidiaries and affiliates',
                        'Consent for Personal Data Collection', 'What is the name of your employer?', 'Employee Number',
                        'Last Name', 'First Name', 'Middle Name', 'Suffix', 'Birthdate',
                        'Where did you get your initial doses of COVID-19 vaccines?',
                        'Do you have prior registration with eZConsult via LTGC?',
                        'Enter your eZConsult Patient ID number.',
                        'Select the vaccine brand.', 'Date of first dose', 'Batch No. / Lot No. of first dose',
                        'Place of vaccination of first dose',
                        'Date of second dose',
                        'Batch No. / Lot No. of second dose', 'Place of vaccination of second dose',
                        'Select preferred vaccination site for your booster shot.',
                        'Company', 'Scenario']

    fileB_column_arr = ['ID', 'Start time', 'Completion time', 'Email', 'Name',
                        'By filling out this form, I agree to participate in the booster vaccination program sponsored by LT Group, Inc., its subsidiaries and affiliates',
                        'Consent for Personal Data Collection', 'What is the name of your employer?', 'Employee Number',
                        'Last Name', 'First Name', 'Middle Name', 'Suffix', 'Birthdate',
                        'Select your appropriate Priority Group',
                        'Are you a member of an indigenous group?', 'Are you a PWD (Person with Disability)?',
                        'Occupation',
                        'Mobile Number', 'Email Address', 'Region', 'Province', 'City', 'Barangay', 'Sex',
                        'Allergy to vaccines or components of vaccines?',
                        'With Comorbidity/ies?', 'Please specify comorbidity/ies',
                        'Company', 'Scenario']

    for comp, records in groups:
        FileA = records.reindex(columns=fileA_column_arr)
        FileB = records.reindex(columns=fileB_column_arr)

        Generate_File(comp, FileA, 'a')
        Generate_File(comp, FileB, 'b')
        Generate_File(comp, records, 'n')

    pass


def GenerateBackup():
    # Create zip file and delete folder company
    companyFiles = os.path.join(dirPath, 'company_' + dateTime)
    companyZipFilesPath = shutil.make_archive(companyFiles, 'zip', compPath)
    shutil.rmtree(compPath)

    toSendDirZip = os.path.join(toSend, os.path.split(companyZipFilesPath)[1])
    shutil.move(companyZipFilesPath, toSendDirZip)

    # Get the latest generated file from directory (For consolited file)
    outPathFiles_list = os.listdir(outPath)
    paths = [os.path.join(outPath, basename) for basename in outPathFiles_list]
    LatestConsoFilePath = max(paths, key=os.path.getctime)

    getFilename = os.path.split(max(paths, key=os.path.getctime))[1]

    # Get Filname from a path (For consolited file)
    # Copy the latest file from out director to destination folder (For consolited file)
    shutil.copy(LatestConsoFilePath, os.path.join(toSend, getFilename))

    toSendPath = os.path.join(dirPath, 'Vaccine_Booster_Conso_' + dateTime2)
    shutil.make_archive(toSendPath, 'zip', toSend)
    # shutil.rmtree(toSend)

    shutil.rmtree(toSend)

    pass


def isGreaterEqualDate(x):
    if x['Date of first dose'] >= x['Date of second dose']:
        return True

    return False


def is28Gap(x):
    vaccineBrand = ['Moderna', 'AstraZeneca']

    if x['Select the vaccine brand.'] in vaccineBrand:
        gap = (x['Date of second dose'] - x['Date of first dose']).days

        if gap >= 28:
            return True
        else:
            return False
    else:
        return

    pass


def Generate_Errors_greatEqualDate(df):

    df_date = df.loc[df['Is Ctrlnum Dup'] == True]

    arrid = []
    errMsg = []

    print(df_date)

    pass


if __name__ == '__main__':
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.width', None)

    today = datetime.today()
    dateTime = today.strftime("%m%d%y%H%M%S")
    dateTime2 = today.strftime("%m%d%y-%H%M")

    print("==============================================================")
    print("Running Scpirt: LTGC Booster shot Consolidation (Big File)......")
    print("==============================================================")

    # Set Dir Path
    # / Users / ran / Documents / Vaccine

    DocumentsFolder = os.path.join(os.path.expanduser("~"), "Documents")
    __rootPath = os.path.join(DocumentsFolder, "Vaccine")

    dirPath = CreateFolders('Booster_Conso', __rootPath)

    inPath = os.path.join(dirPath, "in")
    outPath = os.path.join(dirPath, "out")
    logPath = os.path.join(dirPath, "log")
    compPath = os.path.join(dirPath, "company")
    toSend = os.path.join(dirPath, "toSend_Booster_Signup_Conso")

    # 09465920619 Jao Layout

    templateFilePath = 'src/template/template.xlsx'
    templateFileA = 'src/template/file_a.xlsx'
    templateFileB = 'src/template/file_b.xlsx'

    outFilename = 'BoosterShot_Consolidated_' + dateTime

    # Get all Files in folder
    arrFilenames = os.listdir(inPath)
    arrdfFrames = []

    for inFile in arrFilenames:
        if not inFile == ".DS_Store":
            print("Reading: " + inFile + "......")

            arrdfFrames.append(getData(inFile))

    master_df = pd.concat(arrdfFrames)

    # Generate_Errors_greatEqualDate(master_df)

    Generate_File_per_Company(master_df)

    writeExcel(master_df, templateFilePath, outPath, outFilename)

    GenerateBackup()

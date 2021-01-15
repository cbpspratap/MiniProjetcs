'''
#Filename: SQL_To_Excel_Extract_Export.py
#Author: Chandra Bhanu Singh
#Description: Python code to run extract data using SQL query and export data to excel file at provided location
			  Update the parameters for Source SQL file, Target Data extract file and file name as required
#Created Date: 06-Sep-2020
#Version: v1.0
#Output File name: "output_file_name.xlsx"
#Output Directory: C:\Data\Output\
Revision History:
Last Update Date:
'''

import os
import pandas as pd
import pyodbc
import logging
import logging.handlers
from datetime import datetime

#Logging setup
#Update the log file path(LOG_FILENAME) as per requirement>>>>>>>>
LOG_FILENAME = datetime.now().strftime('C:/Data/Log/GOS_Meter_Export_Python_log_%d_%m_%Y_%H_%M_%S.log')
FORMAT = '%(asctime)-15s %(message)s'
logging.basicConfig(filename=LOG_FILENAME,format=FORMAT, level=logging.INFO)

#Logging
print(datetime.now().strftime(f'%d_%m_%Y %H:%M:%S - Logfile created with name {LOG_FILENAME}'))
logging.info(datetime.now().strftime(f'- Logfile created with name {LOG_FILENAME}'))

#Logging
print(datetime.now().strftime('%d_%m_%Y %H:%M:%S - All modules/packages imported'))
logging.info(datetime.now().strftime('- All modules/packages imported'))

#Establish connection with source Database
#Update the connection parameter(conn) as per requirement>>>>>>>>
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=SERVERNAME;'
                      'Database=DATABASENAME;'
                      'Trusted_Connection=yes;')

cursor = conn.cursor()

#Logging
print(datetime.now().strftime(f'%d_%m_%Y %H:%M:%S - Connection established with SQL Server'))
logging.info(datetime.now().strftime(f'- Connection established with SQL Server'))

#Path of the SQL file which needs to be imported
#Update the SQL file path(sqlpath) as per requirement>>>>>>>>
sqlpath="C:/Data/SQL/SQL_File.sql"

#Name of the file which is to be used for exporting data and logging
#Update the file name(tgtfilename) to be used for export and logging as per requirement>>>>>>>>
tgtfilename= "output_file_name"

#Importing SQL 
opensql= open(sqlpath)

#Reading the SQL file content
readsql=opensql.read()

#Logging
print(datetime.now().strftime(f'%d_%m_%Y %H:%M:%S - This script will export the data of {tgtfilename}'))
logging.info(datetime.now().strftime(f'- This script will export the data of {tgtfilename}'))

#Reading data and creating Pandas Dataframe
querydf=pd.read_sql_query(readsql,conn)

#Logging
print(datetime.now().strftime(f'%d_%m_%Y %H:%M:%S - SQL query run completed and data stored in variable'))
logging.info(datetime.now().strftime(f'- SQL query run completed and data stored in variable'))

#Adding header to pandas Data frame--NOT Required
#df = pd.DataFrame(querydf, columns = ['cEquipment_ID','Legacy_Number','Equipment_Desc','Current_Service_Operator','Equipment_Status','cEquipment_Make_Description','cEquipment_Model_Description','cEquiment_Type_Description','Manufactured_Date'])

#Logging
print(datetime.now().strftime(f'%d_%m_%Y %H:%M:%S - Data frame created and data stored in Data Frame'))
logging.info(datetime.now().strftime(f'- Data frame created and data stored in Data Frame'))

#Update the target file name with path as per requirement>>>>>>>>
targetpath= r'C:\Data\Output\output_file_name.xlsx'

#Exporting data to target location
querydf.to_excel(targetpath, index = False)

#Logging
print(datetime.now().strftime(f'%d_%m_%Y %H:%M:%S - Data frame exported to excel with file name: {tgtfilename}'))
logging.info(datetime.now().strftime(f'- Data frame exported to excel with file name: {tgtfilename}'))

#Logging
print(datetime.now().strftime(f'%d_%m_%Y %H:%M:%S - SUCCES: Excel for {tgtfilename} successfully exported to {targetpath} path'))
logging.info(datetime.now().strftime(f'- SUCCES: Excel for {tgtfilename} successfully exported to {targetpath} path'))

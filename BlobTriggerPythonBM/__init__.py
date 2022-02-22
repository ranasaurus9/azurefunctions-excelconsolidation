import logging
import io
import xlrd
import pandas as pd
import datetime
import azure.functions as func
from io import BytesIO
import xlsxwriter
from os import listdir
from os.path import isfile, join
import pyodbc


def consolidation_function(args):
    #some code that consolidates worksheet data into a dataframe
    return dataframe
    logging.info("consolidation complete")

async def main(inputblob: func.InputStream):
    logging.info(f"Python blob trigger function processed blob \n"
                 f"Name: {inputblob.name}\n")
    s1 = xlrd.open_workbook(file_contents=inputblob.read())

    # connection to SQL DB
    server = 'tcp:azureserver.database.windows.net,1433'
    database = 'SQL DB'
    username = 'username'
    driver = '{ODBC Driver 17 for SQL Server}'
    password = 'password'

    with pyodbc.connect('DRIVER='+driver+';SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password+';Authentication=ActiveDirectoryPassword') as conn:
        with conn.cursor() as cursor:
            print("Connected")

            header = ["A", "B", "C", "D"]
            rsheet = pd.DataFrame(columns=header)
            df = consolidation_function(args)
            df.columns = header
            cursor.fast_executemany = True
            sql_statement = "INSERT INTO Overview_template (A,B,C,D) VALUES (?,?,?,?)"
            cursor.execute(sql_statement)
            logging.info("file uploaded complete")

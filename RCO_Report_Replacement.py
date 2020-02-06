from botocore.config import Config
import win32com.client
import smtplib
from botocore.exceptions import NoCredentialsError
import boto3
import numpy
import pandas as pd
import pyodbc
import os
import logging
from logging.handlers import RotatingFileHandler
import traceback

# configure logging
logger = logging.getLogger("Rotating Log")
logger.setLevel(logging.ERROR)
handler = RotatingFileHandler("log.txt", maxBytes=10000, backupCount=5)
formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)

# add aws credentials
ACCESS_KEY = os.environ.get('AWS_AccessKey')
SECRET_KEY = os.environ.get('AWS_SecretKey')

# configure boto3 client & Proxy access / retries
s3 = boto3.client(
    's3',
    config=Config(
        proxies={
            'https': os.environ.get("https_proxy"),
            'http': os.environ.get("http_proxy")
        },
        connect_timeout=15,
        read_timeout=60,
        retries={'max_attempts': 2}
    ),
    aws_access_key_id=ACCESS_KEY,
    aws_secret_access_key=SECRET_KEY
)

# prepare Office client for PDF creation from spreadsheet
o = win32com.client.Dispatch("Excel.Application")
o.Visible = False
excelPath = r'.\Accepted_RCOs_Report.xlsx'
pdfPath = r'.\Accepted_RCOs_Report.pdf'

# prepare email notification
sender = os.environ.get('DPDAppsProd_Email')
receivers = [os.environ.get('Nick_Email'), os.environ.get('Dan_Email')]
password = os.environ.get('DPDAppsProd_password')

message = """Subject: RCO Report Upload Failed

The RCO Report upload failed.
"""
host = os.environ.get('smtp_host')
port = os.environ.get('smtp_port')
smtpObj = smtplib.SMTP(host, port)

# define connection string DONT POST THE PASSWORD STUFF PUBLICLY
database = os.environ.get('RCO_DB')

try:
    conn = pyodbc.connect(
        DRIVER='{ODBC Driver 13 for SQL Server}',
        SERVER=os.environ.get('PCPC_Server'),
        DATABASE=database,
        PORT=os.environ.get('PORT_VAR'),
        UID=os.environ.get('UID_VAR'),
        PWD=os.environ.get('PWD_VAR')
    )
except Exception as e:
    logger.error(str(e))
    logger.error(traceback.format_exc())

# write the sql query to select accepted RCOs
try:
    SQL_Query = pd.read_sql_query(
        f"SELECT Organization_Name, Organization_Address, Application_Date, Org_Type, Preffered_Contact_Method, Primary_Address, Primary_Email FROM {database}.dbo.RCO_Registration_Information WHERE Status='Accepted'", conn
    )
except Exception as e:
    logger.error(str(e))
    logger.error(traceback.format_exc())
    message += "\n Could Not Connect to SQL Server"
    smtpObj.sendmail(sender, receivers, message)
    smtpObj.quit()


# assign fields to SQL fields to dataframe and print dataframe
df = pd.DataFrame(SQL_Query, columns=['Organization_Name', 'Organization_Address',
                                      'Application_Date', 'Org_Type', 'Preffered_Contact_Method', 'Primary_Address', 'Primary_Email'])

# rename columns
df.rename(columns={'Organization_Name': 'RCO', 'Organization_Address': 'RCO Address', 'Application_Date': 'Application Date', 'Org_Type': 'Organization Type',
                   'Preffered_Contact_Method': 'Preferred Contact Method', 'Primary_Address': 'Contact Address', 'Primary_Email': 'Email'}, inplace=True)
df

# create a pandas excel writer using xlsxwriter as the engine
writer = pd.ExcelWriter(excelPath, engine='xlsxwriter',
                        date_format='mm dd yyyy', datetime_format='mm/dd/yyyy')

# convert dataframe to xlsxwriter excel object
df.to_excel(writer, sheet_name='Sheet1')

# get the xlsxwriter workbook and worksheet objects
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# set the column widths
worksheet.set_column('B:C', 50)
worksheet.set_column('D:F', 25)
worksheet.set_column('G:G', 50)
worksheet.set_column('H:H', 25)

# close pandas excel writer and output excel file
writer.save()

wb = o.Workbooks.Open(os.path.abspath(excelPath))
try:
    ws = wb.WorkSheets(1)
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesTall = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.Orientation = 2

    ws.PageSetup.TopMargin = .5
    ws.PageSetup.BottomMargin = .5
    ws.PageSetup.LeftMargin = 0.5
    ws.PageSetup.RightMargin = 0.5
    ws.PageSetup.HeaderMargin = 0
    ws.PageSetup.FooterMargin = 0

    excelRows = ws.UsedRange.Rows.Count
    ws.PageSetup.PrintArea = f'A1:H{excelRows}'

    wb.WorkSheets(1).Select()
    wb.ActiveSheet.ExportAsFixedFormat(0, os.path.abspath(pdfPath))
except Exception as e:
    logger.error(str(e))
    logger.error(traceback.format_exc())
    wb.Close(True)

wb.Close(True)
print(df)

# upload the excel doc to aws
errorMessage = message

try:
    s3.upload_file(pdfPath, 'dpd-rco-docs-prod',
                   'ReportOnAcceptedRCOs.pdf', ExtraArgs={'ACL': 'public-read'})

except Exception as e:
    logger.error(str(e))
    logger.error(traceback.format_exc())
    message += f"\n Could Not Upload Document To AWS S3 \n {str(e)}"
    smtpObj.sendmail(sender, receivers, message)
    smtpObj.quit()

print("Upload Complete")

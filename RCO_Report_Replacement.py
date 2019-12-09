# import os for environment variable secrets
import os

# import pyodbc to connect to sql database
import pyodbc

# import pandas to turn the sql database to a pandas dataframe
import pandas as pd

# import numpy because apparently pandas needs it
import numpy

# import boto3 for aws stuff
import boto3
from botocore.exceptions import NoCredentialsError
from botocore.config import Config	
boto3.resource('s3', config=Config(proxies=	
    {	
        'https': f'http://{os.environ.get("DPDAppsProd_Email")}:{os.environ.get("DPDAppsProd_password")}@proxy.phila.gov:8080',	
        'http' : f'http://{os.environ.get("DPDAppsProd_Email")}:{os.environ.get("DPDAppsProd_password")}@proxy.phila.gov:8080'	
    }))

# import smtplib for the email sending function
import smtplib

# import win32com to create pdfs
import win32com.client

o = win32com.client.Dispatch("Excel.Application")
o.Visible = False
excelPath = r'.\Accepted_RCOs_Report.xlsx'
pdfPath = r'.\Accepted_RCOs_Report.pdf'

sender = os.environ.get('DPDAppsProd_Email')
receivers = [os.environ.get('Nick_Email'), os.environ.get('Dan_Email')]
password = os.environ.get('DPDAppsProd_password')

message = """Subject: RCO Report Upload Failed

The RCO Report upload failed.
"""
host = os.environ.get('smtp_host')
port = os.environ.get('smtp_port')
smtpObj = smtplib.SMTP(host, port)

# add aws credentials
ACCESS_KEY = os.environ.get('AWS_AccessKey')
SECRET_KEY = os.environ.get('AWS_SecretKey')

# define connection string DONT POST THE PASSWORD STUFF PUBLICLY
conn = pyodbc.connect(
    DRIVER='{ODBC Driver 13 for SQL Server}',
    SERVER=os.environ.get('PCPC_Server'),
    DATABASE=os.environ.get('RCO_DB'),
    PORT=os.environ.get('PORT_VAR'),
    UID=os.environ.get('UID_VAR'),
    PWD=os.environ.get('PWD_VAR')
)

# write the sql query to select accepted RCOs
try:
    SQL_Query = pd.read_sql_query(
        "SELECT Organization_Name, Organization_Address, Application_Date, Org_Type, Preffered_Contact_Method, Primary_Address, Primary_Email FROM RCO_Registration.dbo.RCO_Registration_Information WHERE Status='Accepted'", conn
    )
except:
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

wb.Close(True)
print(df)

# upload the excel doc to aws


def upload_to_aws(local_file, bucket, s3_file):
    s3 = boto3.client('s3', aws_access_key_id=ACCESS_KEY,
                      aws_secret_access_key=SECRET_KEY)
    errorMessage = message
    try:
        s3.upload_file(local_file, bucket, s3_file)
        print("Upload Successful!")
        return True
    except FileNotFoundError:
        print("The file was not found")
        errorMessage += "\n The file was not found"
        smtpObj.sendmail(sender, receivers, errorMessage)
        smtpObj.quit()
        return False
    except NoCredentialsError:
        print("Credentials not available")
        errorMessage += "\n Credentials not available"
        smtpObj.sendmail(sender, receivers, errorMessage)
        smtpObj.quit()
        return False
    else:
        print("Upload failed.")
        errorMessage += "\n Upload failed"
        smtpObj.sendmail(sender, receivers, errorMessage)
        smtpObj.quit()


uploaded = upload_to_aws(pdfPath, 'dpd-rco-docs-prod',
                         'ReportOnAcceptedRCOs.pdf')

#import the modules
from inspect import getfile
import os
from tokenize import group
import pandas as pd
import numpy as np
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
import smtplib

def getFileLocation():
    return "./csv/daas_assignment.xlsx"
def getOutputLocation():
    return "./output/report.xlsx"



def mergeCsv():
    #read the path
    filePath ="./csv"
    cwd = os.path.abspath(filePath)
    #list all the files from the directory
    csv_files = os.listdir(cwd)
    combined_csv = pd.concat([pd.read_csv(filePath+"/"+file) for file in csv_files ])
    excelWriter = pd.ExcelWriter(getOutputLocation())
    # 4) convert csv to excel
    combined_csv.to_excel(
    excelWriter,
    index=False,
    sheet_name='grade')
    excelWriter.save()
    print("Done merging  report")
   

def procesExcel():
    df_data = pd.read_excel(getOutputLocation())
    df_append =df_data.pivot_table(index="Name", columns="Course",values="Score",
     aggfunc='sum', margins=True, margins_name='Total')
    print(df_append)

    with pd.ExcelWriter(getOutputLocation(),mode='a') as writer:  
        df_append.to_excel(writer, sheet_name='summary',index = True)

    return getOutputLocation();

def send_mail(send_from, send_to, subject, text, f=None,
              server="127.0.0.1"):
    assert isinstance(send_to, list)

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text))
    with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )
        # After the file is closed
    part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
    msg.attach(part)

    smtp = smtplib.SMTP(server)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.close()

def processData():
    mergeCsv()
    file =procesExcel()
    send_mail("from_user","hello@daas.ng",f=file)  #you can uncomment put the right details.

processData()


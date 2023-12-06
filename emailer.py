import csv
import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta
import logging
import os
from pylogix import PLC
import schedule
import shutil
import smtplib
import time
import warnings
import threading
import xlwings as xw



def initEmailServer():

    username = 'martinreahydroformbinemailer@gmail.com'
    password = 'SuperPassword'

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()

    try:
        server.login(username, password)

    except:
        logging.exception('login unsuccessful')
        return None

    return server



# def generateFilename():
#     t = time.localtime()
#     date = time.strftime("%Y-%m-%d_", t)
#     hour = (int(time.strftime("%H")) + 3) % 24
#     hour2 = hour - (hour % 3)
#     hour1 = hour2 - 3
#     if(hour1 < 10):
#         date = date + '0' + str(hour1)
#     else: 
#         date = date + str(hour1)
#     # date = date + ':00-'
#     date = date + '-'
#     if(hour2 < 10):
#         date = date + '0' + str(hour2)
#     else: 
#         date = date + str(hour2)
#     date = date + ':00'
#     file_name = '31_log'
#     return file_name



def sendEmail(server):

    bookName = '31_log'
    print(bookName)
    book = bookName.split(".")
    bookName = book[0]
    # Chech if server connection is active (it disconnects itself after a while)
    if not server or not server.noop()[0] == 250:
        print('Email server is down!')
        server = initEmailServer()

    senderEmail = 'martinreahydroformbinemailer@gmail.com'
    mailing_list = ['sankalp.kodandera@martinrea.com'] # add brian.rankin@martinrea.com and olivia.milnaric (and darren)
    receiverEmail = ", ".join(mailing_list)

    # Make an email with excel attachment
    message = MIMEMultipart("alternative")

    message["Subject"] = "Daily Tool-Crib Report"
    body = 'Please find the tool-crib report below.'
    with open(bookName, "rb") as f:
        excelAttachment = MIMEApplication(f.read(),_subtype="txt")
        excelAttachment.add_header('Content-Disposition','attachment',filename=bookName)
        message.attach(excelAttachment)

    message["From"] = senderEmail
    message["To"] = receiverEmail
    message.attach(MIMEText(body, 'plain'))

    # Send the email
    try:
        server.sendmail('martinreahydroformbinemailer@gmail.com', mailing_list, message.as_string())
        return True

    except:
        logging.exception('email sending unsuccessful')
        return False


current_hour = time.localtime().tm_hour

tag_list = ['WeldReworkReasonCode', 'WeldReworkId', 'WeldTechLogged']
t = time.localtime()
date = time.strftime("%Y-%m-%d_", t)

# Check if it's time to generate the file (every 8 hours starting at 7 am)

if current_hour == 7 or (current_hour - 7) % 8 == 0:

    print("he")


elif current_hour == 0:
    # Send files

    print("ha")


with PLC() as comm:
    comm.IPAddress = '10.10.14.12'
    ret = comm.Read(tag_list)

    # for testing 
    for r in ret:
        print(r.Value)

    comm.Close()

    if (ret[2]):

        write_data = [('WeldReworkReasonCode', ret[0]),
                    ('WeldReworkId', ret[1])]

        with open('31_log.csv', 'w') as csv_file: # need to change the name of the file 
            csv_file = csv.writer(csv_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            csv_file.writerows(write_data)
            time.sleep(1)
    else:
        print("yo")

while (1 == 1):
    schedule.every().day.at("11:46", tz="America/New_York").do(sendEmail)

import smtplib
import datetime
import shutil
import logging
import warnings
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta
import os
import time
import schedule
import threading
import xlwings as xw

import time
import os

def generateFilename(h):
    t = time.localtime()
    date = time.strftime("%Y-%m-%d_", t)
    hour = ((h) + 3) % 24
    hour2 = hour - (hour % 3)
    hour1 = hour2 - 3
    if(hour1 < 10):
        date = date + '0' + str(hour1)
    else: 
        date = date + str(hour1)
    date = date + '00-'
    if(hour2 < 10):
        date = date + '0' + str(hour2)
    else: 
        date = date + str(hour2)
    date = date + '00'
    file_name = "Tool_Crib_Log_" + date + ".csv"
    return file_name

def makeDir():
    t = time.localtime()
    month_year = time.strftime("%b-%Y", t)
    day = time.strftime("%d")
    dir = month_year
    if(not os.path.isdir(dir)):
        os.makedirs(dir)
    return dir

def runner():
    while True:
        schedule.run_pending()
        time.sleep(10)

def removeBook(path):
    if not os.path.exists(path):
        warnings.warn("Path does not exist. Remove book failed.")
        return

    os.remove(path)
    
def archiveBook():
    hour = int(time.strftime("%H"))
    small_file = generateFilename(hour - 1)
    t = time.localtime()
    date = time.strftime("%Y-%m-%d", t)
    big_file = "Tool_Crib_Log_" + date + ".csv"
    dir = makeDir()
    try: 
        with open(dir + '\\' + big_file, "x") as file:
            file.write("Time, ID, Name, Part Number, Part Name, Quantity, Cost, Extended Cost\n")
    except FileExistsError:
        pass
    file = open(dir + '\\' + small_file, "r")
    lines = file.readlines()
    bigger_file = open(dir + '\\' + big_file, "a")
    for line in lines:
        if(line != lines[0]):
            bigger_file.write(line)
    file.close()
    bigger_file.close()
    os.remove(dir + '\\' + small_file)
    # timeNow = datetime.now().strftime("%D %S").replace('/','-')
    # if not os.path.exists('./Archives'):
    #     warnings.warn("Path does not exist. Archive book failed.")
    #     os.makedirs('./Archives')
    #     return

    # bookName = 'TC-' +timeNow+ '.xlsx'
    # os.rename('kenobi1/test.xlsx', 'Archives/' + bookName)
    # return bookName


def initEmailServer():

    username = 'martinreahydroformbinemailer@gmail.com'
    password = 'uqemiariemwpsdhg'
    
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()

    try:
        server.login(username, password)

    except:
        logging.exception('login unsuccessful')
        return None

    return server

def sendEmail(server):
    t = time.localtime()
    # yesterday = timeNow - timedelta(1)
   
    dir = makeDir()
    # bookName = "Tool_Crib_Log-" + date + ".csv"
    hour = int(time.strftime("%H"))
    bookName = generateFilename(hour - 1)
    tim = bookName[30:]
    book = bookName.split(".")
    bookName = book[0]
    time.sleep(2)
    bookName = bookName + ".xlsx"
    timeName = datetime.now().strftime("%D %S").replace('/','-')
    # Chech if server connection is active (it disconnects itself after a while)
    if not server or not server.noop()[0] == 250:
        print('Email server is down!')
        server = initEmailServer()

    senderEmail = 'martinreahydroformbinemailer@gmail.com'
    #, 'hasan.khan@martinrea.com', 'siyeon.jung@martinrea.com' , 'sankalp.kodandera@martinrea.com'
    mailing_list = ['neil.patel@martinrea.com'] # add brian.rankin@martinrea.com and olivia.milnaric (and darren)
    receiverEmail = ", ".join(mailing_list)

    # Make an email with excel attachment
    message = MIMEMultipart("alternative")

    message["Subject"] = "Tool Crib Report - " + tim 
    body = 'Please find the tool-crib report below.'
    with open(dir + '\\' + bookName, "rb") as f:
        excelAttachment = MIMEApplication(f.read(),_subtype="txt")
        excelAttachment.add_header('Content-Disposition','attachment',filename=bookName)
        message.attach(excelAttachment)

    message["From"] = senderEmail
    message["To"] = receiverEmail
    message.attach(MIMEText(body, 'plain'))

    # Send the email
    try:
        server.sendmail('martinreahydroformbinemailer@gmail.com', mailing_list, message.as_string())
        os.remove(dir + '\\' + bookName)
        return True

    except:
        logging.exception('email sending unsuccessful')
        return False
    
def emailAndArchive():
    global database
    s = initEmailServer()
    sendEmail(s)
    archiveBook()
    # print("Book archived under name " + str(bname))
    s.quit()
    # copyMasterTable()
    # database = read_excel2("\\\hydfsrv\\Purchasing\\MRO\\")
    print("*Warcraft Peasant Voice* Job's done!")

schedule.every().day.at("07:00", tz="America/New_York").do(emailAndArchive)

t = threading.Thread(target=runner)
t.daemon = True
t.start()

print("Excel emailer schedule started")
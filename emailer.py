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

def runner():
    while True:
        schedule.run_pending()
        time.sleep(10)

def generateFilename(h):
    t = time.localtime()
    date = time.strftime("%Y-%m-%d_", t)
    shift = ""
    if(23 <= h or h < 7):
        shift = "Night"
    elif(7 <= h < 15):
        shift = "Morning"
    else:   
        shift = "Afternoon"
    date = date + shift + "_Shift"
    file_name = "Weld_Tally_Report_" + date + ".csv"
    return file_name

def makeDir():
    t = time.localtime()
    month_year = time.strftime("%b-%Y", t)
    dir = month_year
    if(not os.path.isdir(dir)):
        os.makedirs(dir)
    return dir

def removeBook(path):
    if not os.path.exists(path):
        warnings.warn("Path does not exist. Remove book failed.")
        return
    os.remove(path)
    
def archiveBook():
    t = time.localtime()
    date = time.strftime("%Y-%m-%d", t)
    big_file = "Weld_Tally_Report_" + date + ".csv"
    dir = makeDir()
    try: 
        with open(dir + '\\' + big_file, "x") as file:
            file.write("Time, WeldReworkId, WeldReworkReasonCode\n")
    except FileExistsError:
        pass
    copyFile(dir, big_file, generateFilename(7))
    copyFile(dir, big_file, generateFilename(15))
    copyFile(dir, big_file, generateFilename(23))

def copyFile(dir, big_file, small_file):
    file = open(dir + '\\' + small_file, "r")
    lines = file.readlines()
    bigger_file = open(dir + '\\' + big_file, "a")
    for line in lines:
        if(line != lines[0]):
            bigger_file.write(line)
    file.close()
    bigger_file.close()
    os.remove(dir + '\\' + small_file)

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
    dir = makeDir()
    t = time.localtime()
    date = time.strftime("%Y-%m-%d", t)
    bookNameA = generateFilename(7)
    bookNameN = generateFilename(15)
    bookNameM = generateFilename(23)
    if not server or not server.noop()[0] == 250:
        print('Email server is down!')
        server = initEmailServer()

    senderEmail = 'martinreahydroformbinemailer@gmail.com'
    # , 'hasan.khan@martinrea.com', 'siyeon.jung@martinrea.com' , 'sankalp.kodandera@martinrea.com'
    mailing_list = ['neil.patel@martinrea.com'] # add brian.rankin@martinrea.com and olivia.milnaric (and darren)
    receiverEmail = ", ".join(mailing_list)

    # Make an email with excel attachment
    message = MIMEMultipart("alternative")

    message["Subject"] = "Weld Tally Report - " + date 
    body = 'Please find the Weld Tally Report below.'
    
    try:
        f = open(dir + '\\' + bookNameM)
    except FileNotFoundError:
        pass
    else:
        with open(dir + '\\' + bookNameM, "rb") as f:
            csvAttachment = MIMEApplication(f.read(),_subtype="txt")
            csvAttachment.add_header('Content-Disposition','attachment',filename=bookNameM)
            message.attach(csvAttachment)

    try:
        f = open(dir + '\\' + bookNameA)
    except FileNotFoundError:
        pass
    else:
        with open(dir + '\\' + bookNameA, "rb") as f:
            csvAttachment = MIMEApplication(f.read(),_subtype="txt")
            csvAttachment.add_header('Content-Disposition','attachment',filename=bookNameA)
            message.attach(csvAttachment)
            
    try:
        f = open(dir + '\\' + bookNameN)
    except FileNotFoundError:
        pass
    else:        
        with open(dir + '\\' + bookNameN, "rb") as f:
            csvAttachment = MIMEApplication(f.read(),_subtype="txt")
            csvAttachment.add_header('Content-Disposition','attachment',filename=bookNameN)
            message.attach(csvAttachment)

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

def emailAndArchive():
    s = initEmailServer()
    sendEmail(s)
    archiveBook()
    s.quit()
    print("*Warcraft Peasant Voice* Job's done!")

schedule.every().day.at("12:46", tz="America/New_York").do(emailAndArchive)

t = threading.Thread(target=runner)
t.daemon = True
t.start()

print("Excel emailer schedule started")
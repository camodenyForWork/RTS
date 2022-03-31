from fpdf import FPDF
import gspread
import schedule
import time
import smtplib
from email.message import EmailMessage
from datetime import datetime, date
import os
from os.path import exists
import sqlite3
import threading
import pickle
import csv


class Master():
  db = []
  run = True
  supportEmail = False
  constantEmail = False
  autoPrinting = False


  #Updates Log File
  def LogUpdate(cDateTime, reason, message):
    with open('RTS-Server-Master-Log.txt', 'a') as file:
      file.write("\n"+cDateTime+" - "+reason+": "+message)

  #Grabs current date and time seperated by a comma
  def grabTime():
      now = datetime.now()
      current_time = now.strftime("%H:%M:%S")
      current_date = date.today()
      return str((str(current_date)+"|"+str(current_time)))
  
  #Initializes everything to return state
  def initProg():
    log_verification = os.path.exists('RTS-Server-Master-Log.txt')
    if(log_verification == False):
      f = open("RTS-Server-Master-Log.txt", "w")
      f.close()
      Master.LogUpdate(Master.grabTime(), "Error", "RTS-Server-Master-Log.txt file not found")
      Master.LogUpdate(Master.grabTime(), "Successful", "RTS-Server-Master-Log.txt file created")
    elif(log_verification):
      Master.LogUpdate(Master.grabTime(), "Successful", "RTS-Server-Master-Log.txt file found")
    
    pickleFile = open("RTS-VAR", "rb")
    Master.db = pickle.load(pickleFile)
    Master.supportEmail = Master.db[0]
    Master.constantEmail = Master.db[1]
    Master.autoPrinting = Master.db[2]
    Master.Recurrence.trackingID = Master.db[3]
    Master.Recurrence.lenEn = Master.db[4]
    Master.Recurrence.lenEnEm = Master.db[5]

    t1 = threading.Thread(target=Master.Recurrence.recurrence)
    t2 = threading.Thread(target=Master.ConsoleCommands.uInput)
    t1.start()
    t2.start()

  #Class to run all recurrence operations
  class Recurrence():
    prevDay = ""

    trackingID = 1

    #Amount of data entries last recorded
    lenEn = 0
    lenEnEm = 0

    #Google Sheets API Calls
    sa = gspread.service_account(filename="service_account.json")
    sh = sa.open("Device Repair Form")
    wks = sh.worksheet("Form Responses")
    wksEm = sh.worksheet("Device Repair Form")

    #Makes sure the scripts run at the proper times
    def recurrence():
      schedule.every(5).seconds.do(Master.Recurrence.checkForNewEntries)
      schedule.every().day.at("07:30").do(Master.Recurrence.dailyEmail)
      schedule.every().day.at("23:30").do(Master.Recurrence.dailyDate)

      #Runs the scheduler
      while(Master.run):
        schedule.run_pending()
        time.sleep(1)
    
    #Checks Google Sheets for new cell entry
    def checkForNewEntries():
      #Grabs amount of new entries currently
      try:
        entries = Master.Recurrence.wks.get("A2:A1000")

      #Compares the current check with the last check to find new entries
        if (len(entries) > Master.Recurrence.lenEn):
        #Updates last checked to be accurate
          Master.Recurrence.lenEn += 1

        #Grabs all necessary data from Google Sheets for storage

          first = (Master.Recurrence.wks.acell("B{}".format(Master.Recurrence.lenEn+1)).value)
          last = (Master.Recurrence.wks.acell("C{}".format(Master.Recurrence.lenEn+1)).value)
          formDate = (Master.Recurrence.wks.acell("D{}".format(Master.Recurrence.lenEn+1)).value)
          grade = (Master.Recurrence.wks.acell("E{}".format(Master.Recurrence.lenEn+1)).value)
          ID = (Master.Recurrence.wks.acell("F{}".format(Master.Recurrence.lenEn+1)).value)
          rfr = (Master.Recurrence.wks.acell("G{}".format(Master.Recurrence.lenEn+1)).value)
          details = (Master.Recurrence.wks.acell("H{}".format(Master.Recurrence.lenEn+1)).value)
          PW = (Master.Recurrence.wks.acell("I{}".format(Master.Recurrence.lenEn+1)).value)

        #Creates and prints PDF
          Master.Recurrence.pdfCreate(first+" "+last, formDate, grade, ID, rfr+": "+details,"h", PW)

        #Inserts data into forms table on rtsdb.db database // UPDATE JOTFORM
          Master.Recurrence.sqlInsert(first+" "+last, formDate, grade, PW, " ", ID, rfr, details)
      except Exception as error:
          Master.LogUpdate(Master.grabTime(), "Error", error)

      #Grabs amount of new entries currently
      try:
        entriesEm = Master.Recurrence.wksEm.get("A2:A1000")

      #Compares the current check with the last check to find new entries
        if (len(entriesEm) > Master.Recurrence.lenEnEm):
        #Updates last checked to be accurate
          Master.Recurrence.lenEnEm += 1

        #Grabs all necessary data from Google Sheets for storage
          first = (Master.Recurrence.wksEm.acell("B{}".format(Master.Recurrence.lenEnEm+1)).value)
          last = (Master.Recurrence.wksEm.acell("C{}".format(Master.Recurrence.lenEnEm+1)).value)
          formDate = (Master.Recurrence.wksEm.acell("D{}".format(Master.Recurrence.lenEnEm+1)).value)
          grade = (Master.Recurrence.wksEm.acell("E{}".format(Master.Recurrence.lenEnEm+1)).value)
          ID = (Master.Recurrence.wksEm.acell("F{}".format(Master.Recurrence.lenEnEm+1)).value)
          rfr = (Master.Recurrence.wksEm.acell("G{}".format(Master.Recurrence.lenEnEm+1)).value)
          details = (Master.Recurrence.wksEm.acell("H{}".format(Master.Recurrence.lenEnEm+1)).value)

        #Creates and prints PDF
          Master.Recurrence.pdfCreate(first+" "+last, formDate, grade, ID, rfr+": "+details,"e", "")

        #Inserts data into forms table on rtsdb.db database // UPDATE JOTFORM
          Master.Recurrence.sqlInsert(first+" "+last, formDate, grade, " ", " ", ID, rfr, details)
      except Exception as error:
          Master.LogUpdate(Master.grabTime(),"Error", error)

    #Sends email once a day at 07:30 cst
    def dailyEmail():
      if(Master.constantEmail):
        try:
          sqlC = sqlite3.connect('rtsdb')
          cursor = sqlC.cursor()
          Master.LogUpdate(Master.grabTime(), "Successful", "Successfully Connected to SQLite")

          sqlite_insert_query = f"""SELECT * FROM forms WHERE date='{Master.Recurrence.prevDay}'"""
          cursor.execute(sqlite_insert_query)
          
          forms = cursor.fetchall()
          screen_c = 0
          batt_c = 0
          slow_c = 0
          lost_c = 0
          other_c = 0
          total_c = 0
          for x in forms:
            if x[7] == "Broken Screen":
              screen_c += 1
              total_c += 1
            if x[7] == "Not Charging/Battery":
              batt_c += 1
              total_c += 1
            if x[7] == "Slow/Sluggish":
              slow_c += 1
              total_c += 1
            if x[7] == "Lost Charger":
              lost_c += 1
              total_c += 1
            if x[7] == "Other":
              other_c += 1
              total_c += 1

          Master.LogUpdate(Master.grabTime(), "Successful", "Attempting to send email")
          with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                smtp.ehlo()
                smtp.starttls()
                smtp.ehlo()
                smtp.login('rts@macademy.org', '%EW56Xb7!')
                msg = EmailMessage()
                msg['Subject'] = f"Total Devices: {total_c}. Daily checkup."
                msg['From'] = 'rts@macademy.org'
                msg['To'] = 'cpendergrass@macademy.org'
                msg.set_content(f"{date.today()} It's a new day. Get some coffee. Here's your checkup for the day \n\nTotal devices turned in yesterday: {total_c}\n\nRepair Breakdown:\nBroken Screens: {screen_c}\nNot Charging/Battery: {batt_c}\nSlow/Sluggish: {slow_c}\nLost Charger: {lost_c}\nOther (scary): {other_c}")
                smtp.send_message(msg)
                Master.LogUpdate(Master.grabTime(), "Successful", "E-Mail Sent")
        except sqlite3.Error as error:
            Master.LogUpdate(Master.grabTime(),"Error","Failed to insert data into SQlite table "+str(error))
        finally:
            if sqlC:
                sqlC.close()
                Master.LogUpdate(Master.grabTime(),"Successful","The SQLite connection is closed") 
    
    #Creates and prints off PDF
    def pdfCreate(name,currentDate,grade,asset,details,location, password):
      Master.LogUpdate(Master.grabTime(),"Successful","STARTING PDF CREATION")
      pdf = FPDF("P", "mm", "Letter")
      pdf.add_page()
      pdf.set_font("times", "", 12)
      pdf.set_right_margin(200)
      pdf.image("RepairForm.png", 0,0, 225, 291)
      pdf.text(43,80, name) #name
      pdf.text(170, 80, currentDate) #date
      pdf.text(53, 91.5, grade) #grade level
      pdf.text(50, 103, asset) #asset number
      pdf.text(27.5, 120, details) #details
      pdf.text(107, 91.5, password) #Password
      pdf.output("RepairForm.pdf")
      if(Master.autoPrinting):
        Master.LogUpdate(Master.grabTime(),"Successful","PDF CREATED.. ATTEMPTING TO PRINT")
        if(location == "e"):
          os.system("lpr -P HPED6EDE RepairForm.pdf")
          Master.LogUpdate(Master.grabTime(),"Successful","Printed to HPED6EDE")
        elif(location == "h"):
          os.system("lpr -P -E RepairForm.pdf")
          Master.LogUpdate(Master.grabTime(),"Successful","Printed to HP46D23E")
      
      if(Master.supportEmail):
        Master.LogUpdate(Master.grabTime(), "Successful", "Attempting to send email to support@macademy.org")
        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
          smtp.ehlo()
          smtp.starttls()
          smtp.ehlo()
          smtp.login('rts@macademy.org', '%EW56Xb7!')
          msg = EmailMessage()
          msg['Subject'] = "REPAIR FORM"
          msg['From'] = 'rts@macademy.org'
          msg['To'] = 'support@macademy.org'
          msg.set_content(f'Name: {name}\nStudentID: {asset}\nRFR: {details}')
        
          with open ('RepairForm.pdf', 'rb') as f:
            file_data = f.read()
          
          msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename="RepairForm.pdf")
          smtp.send_message(msg)
          Master.LogUpdate(Master.grabTime(),"Successful", "Email sent to support@macademy.org")
    
    #inserts data into the SQL table
    def sqlInsert(name, date, grade, password, deviceType, studentID, rfr, details):
      try:
        sqlC = sqlite3.connect('rtsdb')
        cursor = sqlC.cursor()
        Master.LogUpdate(Master.grabTime(),"Successful","Successfully Connected to SQLite")

        sqlite_insert_query = f"""INSERT INTO forms
                              (trackingID, name, date, grade, password, deviceType, studentID, rfr, details, OoC) 
                              VALUES 
                              ({Master.Recurrence.trackingID},'{name}','{date}','{grade}','{password}','{deviceType}','{studentID}','{rfr}','{details}',1)"""

        cursor.execute(sqlite_insert_query)
        sqlC.commit()
        Master.LogUpdate(Master.grabTime(), "Successful", "Record inserted successfully into forms table ")
        cursor.close()
        Master.Recurrence.trackingID += 1

      except sqlite3.Error as error:
          Master.LogUpdate(Master.grabTime(),"Error","Failed to insert data into SQlite table "+str(error))
      finally:
          if sqlC:
              sqlC.close()
              Master.LogUpdate(Master.grabTime(),"Successful","The SQLite connection is closed")
    
    def dailyDate():
      Master.Recurrence.prevDay = date.today()

  #Class for console commands
  class ConsoleCommands():
    def uInput():
      while(Master.run):
        userInput = input("Console Command: ")

        userInput = userInput.lower().split(" ")

        if(userInput[0] == "autoticket"):
          Master.ConsoleCommands.autoTicket(userInput[1])
        elif(userInput[0] == "help"):
          Master.ConsoleCommands.uInputHelp()
        elif(userInput[0] == "dailyemail"):
          Master.ConsoleCommands.dailyEmail(userInput[1])
        elif(userInput[0] == "autoprint"):
          Master.ConsoleCommands.autoPrinting(userInput[1])
        elif(userInput[0] == "shutdown" or userInput[0] == "exit" or userInput[0] == "quit"):
          Master.ConsoleCommands.shutDown()
        elif(userInput[0] == "backlog"):
          Master.ConsoleCommands.backLog()
        else:
          print("Sorry I don't recognize that command. Type help for a list of commands")
    
    def shutDown():
      print("Shutting Down")
      Master.LogUpdate(Master.grabTime(),"Successful","Server Shutting Down")
      Master.run = False
      Master.db = []
      Master.db.append(Master.supportEmail)
      Master.db.append(Master.constantEmail)
      Master.db.append(Master.autoPrinting)
      Master.db.append(Master.Recurrence.trackingID)
      Master.db.append(Master.Recurrence.lenEn)
      Master.db.append(Master.Recurrence.lenEnEm)
      pickleFile = open("RTS-VAR", "wb")
      pickle.dump(Master.db, pickleFile)
      pickleFile.close()
      Master.LogUpdate(Master.grabTime(),"Successful","Current state saved - See ya")
      print("That's all folks")

    def backLog():
      file_path = input("Please input the path to the file: ")
      with open(file_path, mode='r') as active:
        csv_reader = csv.reader(active, delimiter=',')
        for row in csv_reader:
          if(row[0] != "First Name"):
            Master.Recurrence.sqlInsert(row[0]+" "+row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8])
    
    def autoTicket(argument):
      if(argument == "enable"):
        Master.supportEmail = True
        print("Auto Ticket email enabled")
        Master.LogUpdate(Master.grabTime(), "Successful", "Auto Ticket email was enabled")
      elif(argument == "disable"):
        print("Auto Ticket email disabled")
        Master.supportEmail = False
        Master.LogUpdate(Master.grabTime(), "Successful", "Auto Ticket email was disabled")
      elif(argument == "status"):
      	if(Master.supportEmail):
      		print("Enabled")
      	else:
      		print("Disabled")
      else:
        print("Sorry I don't recognize that command. Type help Auto Ticket for a list of commands under auto ticket")
    
    def uInputHelp():
      print("autoticket\ndailyemail\nautoprint\nbacklog\n\nshutdown (safely shuts down)")
    
    def dailyEmail(argument):
      if(argument == "enable"):
        print("Daily Email enabled")
        Master.constantEmail = True
        Master.LogUpdate(Master.grabTime(), "Successful", "Daily Email was enabled")
      elif(argument == "disable"):
        print("Daily Email disabled")
        Master.constantEmail = False
        Master.LogUpdate(Master.grabTime(), "Successful", "Daily Email was disabled")
      elif(argument == "status"):
      	if(Master.constantEmail):
      		print("Enabled")
      	else:
      		print("Disabled")
      else:
        print("Sorry I don't recognize that command. Type help dailyemail for a list of commands under dailyemail")
    def autoPrinting(argument):
      if(argument == "enable"):
        print("Auto Printing enabled")
        Master.autoPrinting = True
        Master.LogUpdate(Master.grabTime(), "Successful", "Auto printing was enabled")
      elif(argument == "disable"):
        print("Auto Printing disabled")
        Master.autoPrinting = False
        Master.LogUpdate(Master.grabTime(), "Successful", "Auto printing was disabled")
      elif(argument == "status"):
      	if(Master.autoPrinting):
      		print("Enabled")
      	else:
      		print("Disabled")
      else:
        print("Sorry I don't recognize that command. Type help autoprinting for a list of commands under autoprinting")

  #Class for api calls
  class APICalls():

    def fetchLogin(user):
      print("fetchLogin")

    def fetchData(trackingID):
      print("fetchData")


from main import Master
import tkinter as tk
import threading
import schedule

def initProgram():
  schedule.every(30).seconds.do(checkForUpdates)

def checkForUpdates():
  print("checking")

t1 = threading.Thread(target=Master.initProg)
t2 = threading.Thread(target=initProgram)
t1.start()
t2.start()

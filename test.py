import datetime
from datetime import timedelta   
import json
import random
from openpyxl import Workbook
import string
import tkinter as tk
from tkcalendar import Calendar
from tkinter  import Listbox, ttk


staffList=[]
def gui():
   #TODO create gui
   window = tk.Tk()
   window.title("E-Roster")
   window.geometry("800x400") 

   #* date frame
   dateFrame = tk.LabelFrame(window,text = "Select Date")
   dateFrame.pack(fill="both", expand="yes", side=tk.LEFT)

   calendar = Calendar(dateFrame,selectmode="day")
   calendar.pack(fill = "both",padx = 10,pady =10)

   def generateStaff(house):
      if(house in "Daisy"):
         staffList = ["EN JUDIT","NA SYLVIA -I/C","NA EVANNY-I/C","NA SUKHWINDER-I/C","NA MARY ANN- I/C","NA RINI","SCHA NERISA"]
      elif(house in "Daffoldil"):
         staffList = ["EN ELIZABETH - I/C","SNA MELISSA - I/C","SNA RAISA - I/C","NA AMELINE -I/C","NA PATRICIA-I/C","NA PREETI-I/C","NA ARBITA I/C","NA GAGANDEEP KAUR","HA MON"]
      elif(house in "Gergera"):
         staffList = ["EN NEELA","EN MARJORY ","SNA JESSAMAY","NA SHILPA","NA TERESA","SHCA GINA","HCA NILAR","HA ZAR ZAR"]
      elif(house in "Tulip"):
         staffList = ["ACL RODNEY","EN BUENAFE I/C","EN STEEPHEN I/C","SNA JAYSON I/C","SNA JOFERSON I/C","NA KITH I/C","NA PRIYA I/C","NA ANOOP","HCA JESSIE","HA NILAR","HA HNIN"]
      
      for k in range(listBox.size()):
         listBox.delete(0,tk.END) 

      for i in staffList:
         listBox.insert(0,i) 

   def getDate():
      print(calendar.get_date())
      print(staffList)


   bt_submit = tk.Button(window, text="Generate Table", command= getDate)
   bt_submit.pack(side = tk.BOTTOM, fill = "both", expand = False)

   #* staff frame
   

   houseFrame = tk.LabelFrame(window, text="House")
   houseFrame.pack(fill = "both",expand = "yes",side= tk.LEFT)


   Daisy_btn = tk.Button(houseFrame,text="Daisy", command = lambda: generateStaff("Daisy"))
   Daisy_btn.pack(side = tk.TOP,fill = "both",expand = False)

   Daffoldil_btn = tk.Button(houseFrame,text="Daffoldil", command = lambda: generateStaff("Daffoldil"))
   Daffoldil_btn.pack(side = tk.TOP,fill = "both",expand = False)

   Gergera_btn = tk.Button(houseFrame,text="Gerger", command =  lambda: generateStaff("Gergera"))
   Gergera_btn.pack(side = tk.TOP,fill = "both",expand = False)

   Tulip_btn = tk.Button(houseFrame,text="Tulip", command =  lambda: generateStaff("Tulip"))
   Tulip_btn.pack(side = tk.TOP,fill = "both",expand = False)

   listBox = Listbox(houseFrame)
   listBox.pack()

   
   
   window.mainloop()




def startRoster(date):
   
   roster = {}
   staffa= ["Mark", "Amber", "Todd", "Anita", "Sandy","john","kent"]
   staff = list.copy(staffa)
   random.shuffle(staff)
   t_staff = []
   roles = ["AM","PM"]
   specialroles = ["SS","12H"]
   rolecounter = {
      "AM": 3,
      "PM" : 3,
   } 
   switch = 0
   alphabet = list(string.ascii_uppercase)
   pass

   # TODO create list of 14 days add date, day and staff into it
   for days in range(14):
      #ANCHOR to change
      datevar = datetime.datetime.now()+ timedelta(days=days)
      dateData = {
         "time":{
         "day" : datevar.strftime("%A"),
         "date" : datevar.strftime("%d/%m/%y")
         }
      }
      roster[days+1] = dateData 

      for stafflist in range(len(staff)):
         staffData = {
            staff[stafflist] : {
               "role": "",
               "12H": False,
               "SS" : False
            }
         }
         roster[days+1].update(staffData)


   #TODO night shift every two days 
   nightshift = {
      "role" : "NS"
   }
   for day in range(0,14,2):
      for nsrole in range(1,3): 
         roster[day+nsrole][staff[switch]].update(nightshift)
      switch+=1

   #TODO two rests days after two night shifts 
   switch = 0
   offday = {
      "role":"OD"
   }
   for day in range(0,14,2):
      for odrole in range(3,5): 
         if((day+odrole)>14):
            pass
         else:
            roster[day+odrole][staff[switch]].update(offday)
      switch +=1

   #TODO pre-add SS and 12H
   for days in range(len(roster)-1):
      t_staff = []
      for staf in range(len(staff)):
         if(roster[days+1][staff[staf]]["12H"]==True or roster[days+1][staff[staf]]["role"]=="NS" or roster[days+1][staff[staf]]["role"]=="OD"):
            pass
         else:
            t_staff.append(staff[staf])
      if(len(t_staff)!=0):
         randomStaff = random.choice(t_staff)
         print(randomStaff)
         for i in range (len(roster)):
            roster[i+1][randomStaff]["12H"] = True
      roster[days+1][randomStaff]["role"] = "12H"

   t_staff = []
   for days in range(len(roster)-1):
      t_staff = []
      for staf in range(len(staff)):
         if(roster[days+1][staff[staf]]["SS"]==True or roster[days+1][staff[staf]]["role"]=="NS" or roster[days+1][staff[staf]]["role"]=="OD"):
            pass
         else:
            t_staff.append(staff[staf])
      if(len(t_staff)!=0):
         randomStaff = random.choice(t_staff)
         print("ss: " + randomStaff)
         for i in range (len(roster)):
            roster[i+1][randomStaff]["SS"] = True
      roster[days+1][randomStaff]["role"] = "SS"



   #TODO randomize roles
   for days in range(0,len(roster)):
      for staffMember in range(len(staff)):
         staffrole = roster[days+1][staff[staffMember]]["role"]
         staffspecial = roster[days+1][staff[staffMember]]
         if(staffrole == "NS" or staffrole == "OD" or staffrole == "12H" or staffrole == "SS"):
               pass
         else:
            randomrole = random.randint(0,len(rolecounter)-1)
            assigned = roles[randomrole]

            if(assigned=="AM"):
               rolecounter["AM"] - 1
               if(rolecounter["AM"] == 0):
                  del rolecounter["AM"]
            elif(assigned=="PM"):
               rolecounter["PM"] - 1
               if(rolecounter["PM"] == 0):
                  del rolecounter["PM"]
            roledata = {
               
                  "role" : assigned
               
            }
            roster[days+1][staff[staffMember]].update(roledata)
   #print(json.dumps(roster, indent = 4)) 
   with open('data.txt', 'w') as outfile:
      json.dump(roster, outfile, indent =4)

   #TODO save roster to excel sheet
   workbook = Workbook()
   sheet = workbook.active

   #* add roles
   #roster[columns][row]["role"]
   for row in range(0,len(staffa)):
      rownumber = str(row + 1)
      for column in range (0,len(roster)):
         columnName = alphabet[column]
         cellName = columnName + rownumber
         sheet[cellName] = roster[column+1][staffa[row]]["role"]

   #* insert names column
   sheet.insert_cols(1)
   for names in range (len(staffa)):
      cellName = "A" + str(names + 1) 
      sheet[cellName] = staffa[names]

   #* insert date row
   sheet.insert_rows(1)
   sheet["A1"] = "Name of Staff"
   for rows in range (len(roster)):
      cellName = alphabet[rows+1] + "1"
      sheet[cellName] = roster[rows+1]["time"]["date"]



   workbook.save("roster.xlsx")


if(__name__ == "__main__"):
   gui()
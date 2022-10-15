import threading

import PySimpleGUI as sg

import pandas as pd

import time

from datetime import date

#Opens Excel,  may need to change engine and import VBA script.
workbook = pd.read_excel(r"N:\Python\Paint Out.xlsx")
workbook['Date'] = pd.to_datetime(workbook['Date'])
workbook.set_index('Date')
first_Date = workbook["Date"][0]
todays_date = date.today()
pdate = pd.to_datetime(todays_date)
print(pdate)
print(todays_date)

def splash():
    #this function is pointless until I find a decent splash image
    print("Are you not amused?")

if first_Date == todays_date:
    splash()
else:
    workbook.loc[len(workbook.index)] = [(pdate), 0, 0, 0, 0, 0]
    workbook.sort_values(by='Date', inplace=True, ascending=False)
    workbook.to_excel(r"N:\Python\Paint Out.xlsx", index=False)
    workbook.reset_index(drop=True, inplace=True)
    print(workbook)




# Opens a new thread, calls function saveto, once saveto is complete kills thread.
def savefile():
    t = threading.Timer(1,saveto())
    t.daemon = True
    t.start()
    t.cancel()


# Change to save file to excel as workbook. May need to call out VBA script as well.
def saveto():
    workbook.sort_values(by='Date', inplace=True, ascending=False)
    workbook.to_excel(r"N:\Python\Paint Out.xlsx", index=False)

    print(workbook)


# Need to add column for total output.
def add_one(column):
    workbook.at[0, (column)] += 1
    workbook.at[0, 'Total Out'] += 1



def sub_one(column):
    workbook.at[0, (column)] -= 1
    workbook.at[0, 'Total Out'] -=1



#Currently no counters to show users how many panels have been run. May need to add something.
layout = [
         [sg.Button("Add Tan Back", size=(15,1)),  sg.Button("Subtract Tan Back", size=(15,1))],
         [sg.Button("Add Tan Strike", size=(15,1)),sg.Button("Subtract Tan Strike", size=(15,1))],
         [sg.Button("Add Green Back", size=(15,1)),sg.Button("Subtract Green Back", size=(15,1))],
         [sg.Button("Add Green Strike", size=(15,1)),sg.Button("Subtract Green Strike", size=(15,1))],
         [sg.Button("Exit", size=(15,1)) ],
         ]

# Create the window
window = sg.Window(title="Paint Tally",layout=layout, margins=(150, 150))


# Create an event loop.
while True:
    event, values = window.read(timeout= 120*1000)
    if event == sg.TIMEOUT_KEY:
        savefile()
    if event == "Add Tan Back":
        add_one("Tan Backs Out")
    if event == "Subtract Tan Back":
        sub_one("Tan Backs Out")
    if event == "Add Tan Strike":
        add_one('Tan Strikes Out')
    if event == "Subtract Tan Strike":
        sub_one('Tan Strikes Out')
    if event == "Add Green Back":
        add_one('Green Backs Out')
    if event == "Subtract Green Back":
        sub_one('Green Backs Out')
    if event == "Add Green Strike":
        add_one('Green Strikes Out')
    if event == "Subtract Green Strike":
        sub_one('Green Strikes Out')
    if event == "Exit":
        saveto()
        break


window.close()
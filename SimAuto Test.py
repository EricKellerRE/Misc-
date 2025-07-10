import pandas as pd
import os
import win32com.client

def ListToString(tarList):
    tarList = tarList
    listString = "["
    for item in tarList:
        listString = listString + str(item) + ", "
    listString = listString[:-2]
    listString = listString + "]"
    return listString
objectType = "Area"
AreaNum = 1
AreaName = "1"
fieldList = ["AreaName", "AreaNum"]
fieldList = ListToString(fieldList)
valueList = [AreaNum, AreaName]
valuelist = ListToString(valueList)
sim = win32com.client.Dispatch("pwrworld.SimulatorAuto")
sim.RunScriptCommand("CreateNewCase")
#CreateCommand = "CreateData(" + objectType + ", " + fieldList + str(valueList) + ")"
#sim.RunScriptCommand('LoadCSV("C:\\Users\\erickeller\\Documents\\Grid-Workshop-1\\BusVar.csv", YES);')
sim.RunScriptCommand('CreateData(bus, BusNum, 1);')
#sim.RunScriptCommand(CreateCommand)
# ---- SAVE CASE ----
path = os.path.splitext(os.path.join(os.getcwd(), 'Test'))
new_name = "{}_new{}".format(path[0], path[1])
sim.SaveCase(new_name, 'PWB', True), "Saved as {}".format(new_name)
sim.CloseCase(), "Closed Case"
sim = None
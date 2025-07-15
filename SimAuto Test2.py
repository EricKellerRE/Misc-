import pandas as pd
import os
import win32com.client

objectType = "Bus"
fieldList = ["AreaNum", "BusName", "BusName_NomVolt", "BusNomVolt", "BusNum", "ZoneNum"]
valueList = ["A1", "B1", "B1_138", 138, 1, 1]
sim = win32com.client.Dispatch("pwrworld.SimulatorAuto")
sim.RunScriptCommand("CreateNewCase")
sim.RunScriptCommand(f"CreateData({objectType}, {fieldList}, {valueList})")
# ---- SAVE CASE ----
path = os.path.splitext(os.path.join(os.getcwd(), 'Test'))
new_name = "{}_new{}".format(path[0], path[1])
sim.SaveCase(new_name, 'PWB', True), "Saved as {}".format(new_name)
sim.CloseCase(), "Closed Case"
sim = None
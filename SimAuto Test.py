import pandas as pd
import os
import win32com.client
import math
import locale
import numpy as np
''
#Clever trick. Alternate to declaring variables in the definition
'''class MyClass:
    __slots__ = ['x', 'y']
    
    def __init__(self, x, y):
        self.x = x
        self.y = y
m = MyClass(10, 20)
print(MyClass.__slots__)
print(f"x = {m.x}, y = {m.y}")
'''
#__new__ pattern. 
'''class A:
    def __new__(cls):
        print("Creating instance")
        return super(A, cls).__new__(cls)
    
    def __init__(self):
        print("Initializing instance")'''

'''busList = []
class bus():
    def __init__(self):
        self.AreaNum = 1
        self.BusName = 'B1'
        self.BusName_NomVolt = 'B1_138'
        self.BusNomVolt = 138
        self.BusNum = 1
        self.ZoneNum = 1
testClass = bus()
print(testClass.__dict__.keys())
print(testClass.__dict__.values())'''

def RelativeFileSearch(file):
    found = 0
    path = None
    cwd = os.getcwd()
    while cwd != 'C:\\' :
        cwd=os.path.dirname(cwd)
    for root, dirs, files in os.walk(cwd, topdown=False):
            if found == 0:
                for items in files:
                    if items == file:
                        print("Found: ", file)
                        path = os.path.join(root, file)
                        print(path)
                        print("---------")
                        found = 1
                    else:
                         pass
            else: 
                 pass
    return path

def create_df_data_to_powerworld(pw_object, df, object_type):    
    field_list = df.columns.tolist()    
    field_list = str(field_list).replace("'", "")    
    # print(field_list)    
    for sublist in df.values.tolist():
        # convert flattened list to string
        flattened_list = str(sublist).replace("'", "")        
        # print(flattened_list)        
        r = pw_object.RunScriptCommand(f"CreateData({object_type}, {field_list}, {flattened_list})")   

#Mine -EK
def DFCleaner(dataframe, column):
    print('Cleaning: ', column)
    cleanedDF = pd.DataFrame()
    dataframe.fillna({column: float(0)}, inplace = True)
    for index, row in dataframe.iterrows():
        #print('Value: ', row[column], 'Type: ', type(row[column]))
        cleanedDF.loc[index, column] = row[column] if row[column] != ' ' else 1.0
        #print('Value: ', row[column], 'Processed Type: ', type(row[column]))
    print('Cleaned: ', column)
    print('----------')
    return cleanedDF

ZoneNames = ['Slack Bus', 'AK', 'AL','AR', 'AZ', 'CA', 'CO', 'CT', 'DC', 'DE', 'FL', 'GA', 'HI','IA', 'ID', 'IL', 'IN', 'KS', 'KY', 'LA', 'MA', 'ME', 'MD', 'MI', 'MN','MO', 'MS', 'MT', 'NC', 'ND', 'NE', 'NH', 'NJ', 'NM', 'NV', 'NY', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC', 'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WI', 'WV', 'WY']
ZoneNums= [999, 1, 2, 4, 5, 6, 8, 9, 10, 11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 44, 45, 46, 47, 48, 49, 50, 51, 53, 54, 55, 56]
states = dict(zip(ZoneNames, ZoneNums))
#Pull all of the keys from the dataframe into a list. We're going to construct a new one and use the list so created to delete the rest. 
df = pd.read_excel("C:\\Users\\erickeller\\Downloads\\eia8602024ER\\3_1_Generator_Y2024_Early_Release.xlsx", skiprows=2)
df = df[:-1]
gendf = pd.DataFrame()
busdf = pd.DataFrame()
#object types are "Bus" and "Gen"
#Come back to this. Not all info is found the same file. 
#GenFields = ['AreaName', 'AreaNum', 'BusName', 'BusNum', 'GenAGCAble', 'GenAVRAble', 'GenFuelType', 'GenID', 'GenMVRMax', 'GenMVRMin', 'GenMvrSetPoint', 'GenMWMax', 'GenMWMin', 'GenMWSetPoint', 'GenStatus', 'GenUnitType', 
#             'GenVoltSet', 'Latitude', 'Longitude', 'PowerFactor', 'SubName', 'ZoneName', 'ZoneNum']
GenFields = ['PowerFactor' 'GenMWMin' 'AreaName' 'AreaNum' 'BusName' 'BusNum' 'GenAGCAble' 'GenAVRAble' 'GenFuelType' 'GenID' 'GenMVRMax' 'GenMVRMin' 'GenMvrSetPoint' 'GenMWMax' 'GenMWMin' 'GenMWSetPoint' 'GenStatus' 'GenUnitType' 
             'GenVoltSet' 'SubName' 'ZoneName' 'ZoneNum']
plantDf = pd.read_excel(RelativeFileSearch('2___Plant_Y2024__Early_Release.xlsx'), skiprows = 2)
plantLatitude = dict(zip(plantDf['Plant Name'], plantDf['Latitude']))
plantLongitude = dict(zip(plantDf['Plant Name'], plantDf['Longitude']))
#str, int, str, int, YES/NO, YES/NO, 2 char str, real, real, real, real, real, real, Open/Closed, two-char string, real (bus voltage), real, real, string, string, intl
BusFields = ["AreaNum", "BusName", "BusName_NomVolt", "BusNomVolt", "BusNum", "ZoneName", "ZoneNum"]

LineFields= []
gendf['AreaNum'] = busdf['AreaNum'] = df['Utility ID']
gendf['BusName'] = busdf['BusName'] = df['Plant Name'] + df['Generator ID']
gendf['BusName_NomVolt'] = busdf['BusName_NomVolt'] = busdf['BusName']+'_138'
gendf['BusNomVolt'] = busdf['BusNomVolt'] = 138
gendf['BusNum'] = busdf['BusNum'] = df.index
gendf['ZoneName'] = busdf['ZoneName'] = df['State']
gendf['ZoneNum'] = busdf['ZoneNum'] = df['State'].map(states)

df.fillna({'Nameplate Power Factor': float(1)}, inplace = True)
gendf['PowerFactor'] = DFCleaner(df, 'Nameplate Power Factor')
df.fillna({'Minimum Load (MW)': float(0)}, inplace = True)
gendf['GenMWMin'] = DFCleaner(df, 'Minimum Load (MW)')
df.fillna({'Summer Capacity (MW)': float(1000)}, inplace = True)
gendf['GenMWMax'] = DFCleaner(df, 'Summer Capacity (MW)')
gendf['GenFueltype'] = df['Energy Source 1']
#gendf['SubName'] = df['Plant Name']
#gendf['Latitude'] = df['Plant Name'].map(plantLatitude)
#gendf['Longitude'] = df['Plant Name'].map(plantLongitude)
gendf['GenID'] = '01'
gendf['GenAGCAble'] = 'YES'
gendf['GenAVRAble'] = 'YES'
gendf['GenMVRMax'] = np.sqrt(df['Nameplate Capacity (MW)']**2-(df['Nameplate Capacity (MW)']/gendf['PowerFactor']**2))
gendf.fillna({'GenMVRMax': 1000}, inplace = True)
for index, row in gendf.iterrows():
    row['GenMVRMax'] = row['GenMVRMax'] if not 0 else 1000
    print(row['GenMVRMax'])
#Lagging. Add 0.05 to Nameplate Power Factor for min
gendf['GenMVRMin'] = 0-np.sqrt((gendf['GenMWMin']/(gendf['PowerFactor']+0.05))**2-gendf['GenMWMin']**2)
gendf.fillna({'GenMVRMin': 0}, inplace = True)
gendf['GenMvrSetpoint'] = (gendf['GenMVRMax'] + gendf['GenMVRMin']) / 2
gendf['GenMWSetPoint'] = (gendf['GenMWMax']+gendf['GenMWMin'])/2
gendf['GenStatus'] = 'Open'
gendf['GenVoltSet'] = 138
gendf['GenUnitType'] = df['Prime Mover']
gendf['ZoneName'] = df['State']
gendf['ZoneNum'] = df['State'].map(states)

sim = win32com.client.Dispatch("pwrworld.SimulatorAuto")
sim.RunScriptCommand("CreateNewCase")
create_df_data_to_powerworld(sim, busdf, 'Bus')
BusFieldsSlack = ["AreaNum", "BusName", "BusName_NomVolt", "BusNomVolt", "BusNum", "ZoneName", "ZoneNum"]
slackBus = [99999, 'slack', 'slack_138', 138, 99999, 'slack', 99999]
busFieldsList = str(BusFieldsSlack).replace("'", "")
busValuesList = str(slackBus).replace("'", "")
check = sim.RunScriptCommand(f"CreateData({'Bus'}, {busFieldsList}, {busValuesList})")     
create_df_data_to_powerworld(sim, gendf, 'Gen')   
# ---- SAVE CASE ----
path = os.path.splitext(os.path.join(os.getcwd(), 'Test'))
new_name = "{}_new{}".format(path[0], path[1])
sim.SaveCase(new_name, 'PWB', True), "Saved as {}".format(new_name)
sim.CloseCase(), "Closed Case"
sim = None
 
#df = pd.read_excel(r"C:\Users\erickeller\Downloads\eia8602024ER\3_1_Generator_Y2024_Early_Release.xlsx", skiprows=2)
#df = df[:-1]
#print(df.columns.values.tolist())
#df['rfid'] = df['Plant Code'].astype(str) + '_' + df['Generator ID']
#print(df['rfid'])
#df['Category'] = df['Experience'].apply(

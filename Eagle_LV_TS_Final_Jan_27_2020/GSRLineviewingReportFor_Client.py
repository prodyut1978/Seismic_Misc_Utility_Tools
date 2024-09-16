import os
import pandas as pd
import glob
import datetime
import csv
import openpyxl
import PySimpleGUI as sg
import pickle
import tkinter.ttk as ttk
from tkinter.filedialog import asksaveasfile
from tkinter.filedialog import askopenfilenames
from tkinter import simpledialog
from tkinter import*
import tkinter.messagebox
from tkinter import filedialog
import sys

LV_TS_path = r'C:\LV_TS_Report'
if not os.path.exists(LV_TS_path):
    os.makedirs(LV_TS_path)

LV_TS_path_Support_File = r'C:\LV_TS_Report\Support_Files'
if not os.path.exists(LV_TS_path_Support_File):
    os.makedirs(LV_TS_path_Support_File)

LV_TS_path_Parameter_File = r'C:\LV_TS_Report\Support_Files\LVParameter'
if not os.path.exists(LV_TS_path_Parameter_File):
    os.makedirs(LV_TS_path_Parameter_File)

LV_TS_path_Support_File_BattVolt = r'C:\LV_TS_Report\Support_Files\LV_BattVoltage'
if not os.path.exists(LV_TS_path_Support_File_BattVolt):
    os.makedirs(LV_TS_path_Support_File_BattVolt)

LV_TS_path_Support_File_Impedence = r'C:\LV_TS_Report\Support_Files\LV_Impedence'
if not os.path.exists(LV_TS_path_Support_File_Impedence):
    os.makedirs(LV_TS_path_Support_File_Impedence)

if not os.path.exists("C:\LV_TS_Report\Support_Files\LVParameter\LV_fixed_parameters"):
        PickleJobName    = ''
        PickleClientName = ''
        PickleCrewName   = ''
else:
        Pickle_in           = open("C:\LV_TS_Report\Support_Files\LVParameter\LV_fixed_parameters","rb")
        pickle_dict         = pickle.load(Pickle_in)
        PickleJobName       = pickle_dict[1]
        PickleClientName    = pickle_dict[2]
        PickleCrewName      = pickle_dict[3]
    

Default_Date_today   = datetime.date.today()
layout = [[sg.Text('Enter Desired BattVoltage Value (Default = 16.0) :',      size=(42, 1)), sg.InputText(16.0)],
          [sg.Text('Enter High Impedence Value for GSRX3 (Default = 9465) :', size=(42, 1)), sg.InputCombo((9465, 950),default_value=9465,  size=(8, 1))],
          [sg.Text('Enter Low Impedence Value for GSRX3 (Default = 7128) :',  size=(42, 1)), sg.InputCombo((7128, 750),default_value=7128,  size=(8, 1))],          
          [sg.Text('Enter High Impedence Value for GSR-1C (Default = 9465) :', size=(42, 1)), sg.InputCombo((9465, 950),default_value=9465,  size=(8, 1))],          
          [sg.Text('Enter Low Impedence Value for GSR-1C (Default = 7128) :',  size=(42, 1)), sg.InputCombo((7128, 750),default_value=7128,  size=(8, 1))],   
          
          [sg.Text('Enter Job Name :',    size=(25, 1), auto_size_text=True), sg.InputText(PickleJobName, size=(33,1))],
          [sg.Text('Enter Client Name :', size=(25, 1), auto_size_text=True), sg.InputText(PickleClientName, size=(33,1))],
          [sg.Text('Enter Crew Name :', size=(25, 1), auto_size_text=True), sg.InputText(PickleCrewName, size=(33,1))],
          [sg.Text('Production Date (YYYY-MM-DD) :', size=(25, 1), auto_size_text=True), sg.InputText(Default_Date_today, size=(33,1))],
          [sg.Submit(), sg.Cancel()]]

window = sg.Window('LineViewer QC Parameters:', auto_size_text=True, default_element_size=(10, 1)).Layout(layout)      
event, values = window.Read()

if event is None or event == 'Cancel':
        sg.PopupAutoClose('Exiting LV QC',line_width=60)
else:        
        DesiredBattVoltage              = float(values[0])
        High_Threshold_Imp_GSRX3        = float(values[1])
        Low_Threshold_Imp_GSRX3         = float(values[2])        
        High_Threshold_Imp_GSR_1C       = float(values[3])            
        Low_Threshold_Imp_GSR_1C        = float(values[4])            
        PickleJobName                   = (values[5])
        PickleClientName                = (values[6])
        PickleCrewName                  = (values[7])
        pickle_dict = {1:PickleJobName, 2:PickleClientName, 3:PickleCrewName}
        pickle_out  = open("C:\LV_TS_Report\Support_Files\LVParameter\LV_fixed_parameters","wb")
        pickle.dump(pickle_dict, pickle_out)
        pickle_out.close()

        JobName         = 'Job Name        :  ' + str(PickleJobName)
        JobName_FileName= " " + str(PickleJobName)+ " "
        ClientName      = 'Client Name    :  '  + str(PickleClientName)
        CrewName        = 'Crew Name          : '   + str(PickleCrewName)
        PreparedDate    = 'Production Date : '    + (values[8])

        root = Tk()
        root.fileList = askopenfilenames(initialdir = "/", title = "Import Geospace LineViewer Files" , filetypes=[('CSV File', '*.csv'), ('Excel File', ('*.xls', '*.xlsx'))])
        Length_fileList  =  len(root.fileList)
        if Length_fileList >0:
                dfList = []
                LVList = []
                for filename in root.fileList:
                    df                   = pd.read_csv(filename, sep=',',low_memory=False)
                    df                   = df.iloc[:,:]
                    CaseSN               = df.loc[:,'Case SN']
                    DeviceType           = df.loc[:,' Device Type']
                    Deployment_Time      = df.loc[:, 'Deployment Time (UTC)']
                    Line                 = df.loc[:,'Line']
                    Station              = df.loc[:,' Station']
                    BattVoltage          = df.loc[:,' BatteryVoltage']
                    Ch1Impedance         = df.loc[:,'Ch1 Impedance(Ohms)']
                    Ch2Impedance         = df.loc[:,'Ch2 Impedance(Ohms)']
                    Ch3Impedance         = df.loc[:,'Ch3 Impedance(Ohms)']
                    LastScanTime         = df.loc[:,' Last Scan Time (UTC)']
                    LineViewerName       = df.loc[:,'LineViewerName']
                    Recording_State      = df.loc[:,'Record State']
                    ScriptChecksum       = df.loc[:,'Script Checksum']
                    ParamChecksum        = df.loc[:,' Param Checksum']
                    TestResultsHealth    = df.loc[:,'TestResultsHealth']
                    GpsHealth            = df.loc[:,'GpsHealth']
                    MemoryHealth         = df.loc[:,'MemoryHealth']
                    GsrHealth            = df.loc[:,'GsrHealth']
                    CurrentActivity      = df.loc[:,'CurrentActivity']
                    Latitude             = df.loc[:,'Latitude (DecDeg)']
                    Longitude            = df.loc[:,' Longitude (DecDeg)']
                    Altitude             = df.loc[:,' Altitude (m)']

                    column_names = [CaseSN, DeviceType, Deployment_Time, Line, Station, BattVoltage,
                                    Ch1Impedance, Ch2Impedance, Ch3Impedance, LastScanTime, LineViewerName,
                                    Recording_State,ScriptChecksum,ParamChecksum,TestResultsHealth,GpsHealth,
                                    MemoryHealth,GsrHealth,CurrentActivity,Latitude,Longitude,Altitude]
                    catdf = pd.concat (column_names,axis=1,ignore_index =True)
                    dfList.append(catdf)
                    LVList.append (filename)

                # Combine_All_LV_Files and Selected Column
                concatDf = pd.concat(dfList,axis=0)
                concatDf.rename(columns={0:'CaseSN', 1:'DeviceType', 2:'Deployment_Time', 3:'Line', 4:'Station', 5:'BattVoltage',
                                         6:'Ch1Impedence',7:'Ch2Impedence',8:'Ch3Impedence',9:'LastScanTime',10:'LineViewerName',
                                         11:'Recording_State',12:'ScriptChecksum',13:'ParamChecksum',14:'TestResultsHealth',
                                         15:'GpsHealth',16:'MemoryHealth',17:'GsrHealth',18:'CurrentActivity',
                                         19:'Latitude',20:'Longitude',21:'Altitude'},inplace = True)
                LV_Rep = pd.DataFrame(concatDf)
                LV_Rep["QC_Comments"]     = LV_Rep.shape[0]*[" "]
                root.destroy()

                # Export_Combined Report
                os.chdir = ("C:\\LV_TS_Report")
                outfile_Modified =("C:\\LV_TS_Report\\Support_Files\\Combined_LV_Report.csv")
                LV_Rep.to_csv(outfile_Modified,index=None)

                # LV QC Total Accomplished-Modified
                LV_QC_Valid  = LV_Rep.loc[:,['Line','Station','CaseSN','DeviceType','Deployment_Time','LastScanTime',
                                             'BattVoltage','LineViewerName','Ch1Impedence','Ch2Impedence','Ch3Impedence',
                                             'Latitude','Longitude','Altitude',
                                             'Recording_State','ScriptChecksum','ParamChecksum','TestResultsHealth',
                                             'GpsHealth','MemoryHealth','GsrHealth','CurrentActivity','QC_Comments']]
                LV_QC_Valid  = LV_QC_Valid[pd.to_numeric(LV_QC_Valid.Line,   errors='coerce').notnull()]
                LV_QC_Valid  = LV_QC_Valid[pd.to_numeric(LV_QC_Valid.Station,errors='coerce').notnull()]
                LV_QC_Valid  = LV_QC_Valid[pd.to_numeric(LV_QC_Valid.CaseSN, errors='coerce').notnull()]
                LV_QC_Valid['Line']    = LV_QC_Valid.Line.astype (int)
                LV_QC_Valid['Station'] = LV_QC_Valid.Station.astype (int)
                LV_QC_Valid['CaseSN']  = LV_QC_Valid.CaseSN.astype (int)
                LV_QC_accomp_Detailed_Rep = pd.DataFrame(LV_QC_Valid)
                

                # LV QC Report without invalid line station and case serial number
                LV_QC_accomp_valid_LN_ST = pd.DataFrame(LV_QC_Valid)
                LV_QC_accomp_valid_LN_ST = LV_QC_accomp_valid_LN_ST[(LV_QC_accomp_valid_LN_ST.Line != -1)&
                                          (LV_QC_accomp_valid_LN_ST.Station != -1)]

                LV_QC_accomp_valid_LN_ST['Line_Station_Combined'] = (LV_QC_accomp_valid_LN_ST['Line'].map(str)+LV_QC_accomp_valid_LN_ST['Station'].map(str))
                LV_QC_accomp_valid_LN_ST['Line_Station_Combined']  = LV_QC_accomp_valid_LN_ST.Line_Station_Combined.astype (float)
                LV_QC_accomp_valid_LN_ST  = LV_QC_accomp_valid_LN_ST.loc[:,
                                             ['Line','Station','Line_Station_Combined','CaseSN','DeviceType',
                                              'Deployment_Time','LastScanTime','BattVoltage','LineViewerName',
                                              'Ch1Impedence','Ch2Impedence','Ch3Impedence','Latitude','Longitude','Altitude',                              
                                              'Recording_State','ScriptChecksum','ParamChecksum','TestResultsHealth',
                                              'GpsHealth','MemoryHealth','GsrHealth','CurrentActivity','QC_Comments']]

                outfile_LV_QC_accomp_valid_LN_ST   =("C:\\LV_TS_Report\\Support_Files\\LV_TS_accomplished_Valid_Line_Station.csv")
                LV_QC_accomp_valid_LN_ST.to_csv(outfile_LV_QC_accomp_valid_LN_ST,index=None)


                # Battery Voltage Fail Check############## 
                LV_Rep_BattVoltage_Check  = pd.DataFrame(LV_QC_accomp_valid_LN_ST)
                
                LV_Rep_BattVoltage_Check = LV_Rep_BattVoltage_Check[(LV_Rep_BattVoltage_Check.DeviceType == ' 1-C GSR')|
                                                        (LV_Rep_BattVoltage_Check.DeviceType == ' 1-C GSX')|
                                                        (LV_Rep_BattVoltage_Check.DeviceType == ' GSR X3')]
                LV_Rep_BattVoltage_Check = LV_Rep_BattVoltage_Check.reset_index(drop=True)

                def trans_Batt_Volt_Change(y):
                    if y == ' OPEN':
                        return 0
                    elif y == 'OPEN':
                        return 0
                    elif y == 'Off':
                        return 0
                    elif y == ' Off':
                        return 0
                    elif y == ' ?':
                        return 0
                    elif y == '?':
                        return 0
                    elif y == '-':
                        return 0
                    elif y == ' -':
                        return 0
                    elif y == 'Unknown':
                        return 0
                    elif y == ' Unknown':
                        return 0
                    else:
                        return y

                LV_Rep_BattVoltage_Check['Batt_Voltage'] = LV_Rep_BattVoltage_Check['BattVoltage'].apply(trans_Batt_Volt_Change)
                LV_Rep_BattVoltage_Check['Batt_Voltage'] = LV_Rep_BattVoltage_Check.Batt_Voltage.astype (float)

                LV_Rep_BattVoltage_Fail = pd.DataFrame(LV_Rep_BattVoltage_Check)
                LV_Rep_BattVoltage_Pass = pd.DataFrame(LV_Rep_BattVoltage_Check)

                LV_Rep_BattVoltage_Fail = LV_Rep_BattVoltage_Fail[LV_Rep_BattVoltage_Fail.Batt_Voltage < DesiredBattVoltage]
                LV_Rep_BattVoltage_Pass = LV_Rep_BattVoltage_Pass[LV_Rep_BattVoltage_Pass.Batt_Voltage >= DesiredBattVoltage]
                LV_Rep_BattVoltage_Fail = LV_Rep_BattVoltage_Fail.loc[:,
                                    ['Line','Station','Line_Station_Combined','CaseSN','DeviceType','BattVoltage','Batt_Voltage','Deployment_Time','LastScanTime',
                                     'Ch1Impedence', 'Ch2Impedence','Ch3Impedence','Latitude','Longitude','Altitude','LineViewerName' ,'QC_Comments']]

                LV_Rep_BattVoltage_Pass = LV_Rep_BattVoltage_Pass.loc[:,
                                    ['Line','Station','Line_Station_Combined','CaseSN','DeviceType','BattVoltage','Batt_Voltage','Deployment_Time','LastScanTime',
                                     'Ch1Impedence', 'Ch2Impedence','Ch3Impedence','Latitude','Longitude','Altitude','LineViewerName','QC_Comments']]

                LV_Rep_BattVoltage_Fail = LV_Rep_BattVoltage_Fail.drop_duplicates(['Line_Station_Combined'],keep='last')
                LV_Rep_BattVoltage_Fail = LV_Rep_BattVoltage_Fail.reset_index(drop=True)
                LV_Rep_BattVoltage_Pass = LV_Rep_BattVoltage_Pass.drop_duplicates(['Line_Station_Combined'],keep='last')
                LV_Rep_BattVoltage_Pass = LV_Rep_BattVoltage_Pass.reset_index(drop=True)

                LV_Rep_BattVoltage_Fail_M = pd.merge(LV_Rep_BattVoltage_Fail, LV_Rep_BattVoltage_Pass, how='left', on=['Line_Station_Combined','Line_Station_Combined'])
                LV_Rep_BattVoltage_Fail_M.drop(columns=['Batt_Voltage_y','Line_y','Station_y','DeviceType_y','Deployment_Time_y',
                                                        'Latitude_y','Longitude_y','Altitude_y', 'QC_Comments_y'],axis=1,inplace=True)

                LV_Rep_BattVoltage_Fail_M["QC_Comments_x"] = LV_Rep_BattVoltage_Fail_M.shape[0]*["Right- After_QC"]

                
                ### Client Statistics Report- Batt_Volt
                Batt_Volt_RNG_Six = LV_Rep_BattVoltage_Pass[(LV_Rep_BattVoltage_Pass.Batt_Voltage > 16.5)].count()
                Batt_Volt_RNG_Six = Batt_Volt_RNG_Six['Line_Station_Combined']

                Batt_Volt_RNG_Five = LV_Rep_BattVoltage_Pass[(LV_Rep_BattVoltage_Pass.Batt_Voltage > 16.4)&
                                                (LV_Rep_BattVoltage_Pass.Batt_Voltage <= 16.5)].count()
                Batt_Volt_RNG_Five = Batt_Volt_RNG_Five['Line_Station_Combined']

                Batt_Volt_RNG_Four = LV_Rep_BattVoltage_Pass[(LV_Rep_BattVoltage_Pass.Batt_Voltage > 16.3)&
                                                (LV_Rep_BattVoltage_Pass.Batt_Voltage <= 16.4)].count()
                Batt_Volt_RNG_Four = Batt_Volt_RNG_Four['Line_Station_Combined']

                Batt_Volt_RNG_Three = LV_Rep_BattVoltage_Pass[(LV_Rep_BattVoltage_Pass.Batt_Voltage > 16.2)&
                                                (LV_Rep_BattVoltage_Pass.Batt_Voltage <= 16.3)].count()
                Batt_Volt_RNG_Three = Batt_Volt_RNG_Three['Line_Station_Combined']

                Batt_Volt_RNG_Two   = LV_Rep_BattVoltage_Pass[(LV_Rep_BattVoltage_Pass.Batt_Voltage >= 16)&
                                                (LV_Rep_BattVoltage_Pass.Batt_Voltage <= 16.2)].count()
                Batt_Volt_RNG_Two   = Batt_Volt_RNG_Two['Line_Station_Combined']

                # Check Batt Volt Fail_Stats
                LV_Rep_BattVoltage_Fail_Client   = LV_Rep_BattVoltage_Fail_M[(LV_Rep_BattVoltage_Fail_M.CaseSN_y.isnull())]

                Batt_Volt_RNG_One = LV_Rep_BattVoltage_Fail_Client[(LV_Rep_BattVoltage_Fail_Client.Batt_Voltage_x > 15.9)&
                                                (LV_Rep_BattVoltage_Fail_Client.Batt_Voltage_x < 16.0)].count()
                Batt_Volt_RNG_One = Batt_Volt_RNG_One['Line_Station_Combined']
                Batt_Volt_RNG_Zero = LV_Rep_BattVoltage_Fail_Client[(LV_Rep_BattVoltage_Fail_Client.Batt_Voltage_x > 15.7)&
                                                                    (LV_Rep_BattVoltage_Fail_Client.Batt_Voltage_x <= 15.9)].count()
                Batt_Volt_RNG_Zero = Batt_Volt_RNG_Zero['Line_Station_Combined']

                Batt_Volt_RNG_MOne = LV_Rep_BattVoltage_Fail_Client[(LV_Rep_BattVoltage_Fail_Client.Batt_Voltage_x > 15.5)&
                                                                    (LV_Rep_BattVoltage_Fail_Client.Batt_Voltage_x <= 15.7)].count()
                Batt_Volt_RNG_MOne = Batt_Volt_RNG_MOne['Line_Station_Combined']

                Batt_Volt_RNG_Mtwo = LV_Rep_BattVoltage_Fail_Client[(LV_Rep_BattVoltage_Fail_Client.Batt_Voltage_x > 15.3)&
                                                                    (LV_Rep_BattVoltage_Fail_Client.Batt_Voltage_x <= 15.5)].count()
                Batt_Volt_RNG_Mtwo = Batt_Volt_RNG_Mtwo['Line_Station_Combined']

                Batt_Volt_RNG_Mthree = LV_Rep_BattVoltage_Fail_Client[(LV_Rep_BattVoltage_Fail_Client.Batt_Voltage_x <=  15.3)].count()
                                                                    
                Batt_Volt_RNG_Mthree = Batt_Volt_RNG_Mthree['Line_Station_Combined']

                LV_QC_Batt_Stat_Client = pd.DataFrame({'Below 15.3':[Batt_Volt_RNG_Mthree],
                                                       '15.3 - 15.5':[Batt_Volt_RNG_Mtwo],
                                                       '15.5 - 15.7':[Batt_Volt_RNG_MOne],
                                                       '15.7 - 15.9':[Batt_Volt_RNG_Zero],
                                                       '15.9 - 16.0':[Batt_Volt_RNG_One],
                                                       '16.0 - 16.2':[Batt_Volt_RNG_Two],
                                                       '16.2 - 16.3':[Batt_Volt_RNG_Three],                                   
                                                       '16.3 - 16.4':[Batt_Volt_RNG_Four],
                                                       '16.4 - 16.5':[Batt_Volt_RNG_Five],
                                                       '16.5 and Above':[Batt_Volt_RNG_Six]},index=None)

                LV_QC_Batt_Stat_Client= LV_QC_Batt_Stat_Client.T

                LV_QC_Batt_Stat_Client = LV_QC_Batt_Stat_Client.reset_index(drop=False)
                
                LV_QC_Batt_Stat_Client.rename(columns = {'index':'Battery Voltage (V)', 0:'Number of Receivers'},inplace = True)
                LV_QC_Batt_Stat_Client                = LV_QC_Batt_Stat_Client.reset_index(drop=True)
                Comments                              = ['LOW', 'LOW','LOW', 'LOW','LOW', 'AVERAGE', 'HIGH', 'HIGH','HIGH', 'HIGH']
                Index                                 = [1,2,3,4,5,6,7,8,9,10]
                LV_QC_Batt_Stat_Client['Comments']    = Comments
                LV_QC_Batt_Stat_Client['Index']       = Index
                LV_QC_Batt_Stat_Client = LV_QC_Batt_Stat_Client.loc[:,
                                           ['Index','Battery Voltage (V)','Number of Receivers','Comments']]

                LV_Rep_BattVoltage_Fail_M.drop(columns=['Batt_Voltage_x'],axis=1,inplace=True)

                ### Battery Voltage Fail/Pass Export Report
                
                outfile_BattVoltage_FAIL =("C:\\LV_TS_Report\\Support_Files\\LV_BattVoltage\\BattVoltageFAIL_LVReport.csv")
                LV_Rep_BattVoltage_Fail_M.to_csv(outfile_BattVoltage_FAIL,index=None)

                outfile_BattVoltage_PASS =("C:\\LV_TS_Report\\Support_Files\\LV_BattVoltage\\BattVoltagePASS_LVReport.csv")
                LV_Rep_BattVoltage_Pass.to_csv(outfile_BattVoltage_PASS,index=None)

                outfile_LV_QC_Batt_Stat_Client =("C:\\LV_TS_Report\\LV_QC_Batt_Stat_Client.csv")
                LV_QC_Batt_Stat_Client.to_csv(outfile_LV_QC_Batt_Stat_Client,index=False)


                # Impedence Fail Check ############## Impedence Fail Check ###########
                LV_Rep  = pd.DataFrame(LV_QC_accomp_valid_LN_ST)
                LV_Rep  = LV_Rep[(LV_Rep.DeviceType == ' 1-C GSR')|
                                 (LV_Rep.DeviceType == ' 1-C GSX')|
                                 (LV_Rep.DeviceType == ' GSR X3')]
                
                LV_Rep = LV_Rep.reset_index(drop=True)
                                                        
                def trans_Ch_Impedance_Change(x):
                    if x == ' OPEN':
                        return 99999999
                    elif x == 'OPEN':
                        return 99999999
                    elif x == 'Off':
                        return 88888888
                    elif x == ' Off':
                        return 88888888
                    elif x == ' ?':
                        return 77777777
                    elif x == '?':
                        return 77777777
                    elif x == '-':
                        return 66666666
                    elif x == ' -':
                        return 66666666
                    elif x == 'Unknown':
                        return 66666666
                    elif x == ' Unknown':
                        return 66666666
                    else:
                        return x      

                # Apply filter to change OPEN, off, -, NA values..... xxxxxxx
                LV_Rep['Ch1_Impedance'] = LV_Rep['Ch1Impedence'].apply(trans_Ch_Impedance_Change)
                LV_Rep['Ch2_Impedance'] = LV_Rep['Ch2Impedence'].apply(trans_Ch_Impedance_Change)
                LV_Rep['Ch3_Impedance'] = LV_Rep['Ch3Impedence'].apply(trans_Ch_Impedance_Change)

                outfile_LV_ImpedenceValidLineStationReport=("C:\\LV_TS_Report\\Support_Files\\LV_Impedence\\LV_ImpedenceValidLineStationReport.csv")
                LV_Rep.to_csv(outfile_LV_ImpedenceValidLineStationReport,index=None)


                # Channge 'Ch1_Impedance','Ch2_Impedance','Ch3_Impedance' to object to int

                LV_Rep_GSRX3  = LV_Rep[(LV_Rep.DeviceType == ' GSR X3')]
                LV_Rep_GSRX3  = LV_Rep_GSRX3.reset_index(drop=True)
                
                LV_Rep_GSRX3['Ch1_Impedance'] = LV_Rep_GSRX3.Ch1_Impedance.astype (float)
                LV_Rep_GSRX3['Ch2_Impedance'] = LV_Rep_GSRX3.Ch2_Impedance.astype (float)
                LV_Rep_GSRX3['Ch3_Impedance'] = LV_Rep_GSRX3.Ch3_Impedance.astype (float)

                LV_Rep_GSRX1  = LV_Rep[(LV_Rep.DeviceType == ' 1-C GSR')|
                                 (LV_Rep.DeviceType == ' 1-C GSX')]
                LV_Rep_GSRX1  = LV_Rep_GSRX1.reset_index(drop=True)
                LV_Rep_GSRX1['Ch1_Impedance'] = LV_Rep_GSRX1.Ch1_Impedance.astype (float)

                outfile_LV_Rep_GSRX1=("C:\\LV_TS_Report\\Support_Files\\LV_Impedence\\LV_Rep_GSRX1_Only.csv")
                LV_Rep_GSRX1.to_csv(outfile_LV_Rep_GSRX1,index=None)

                outfile_LV_Rep_GSRX3=("C:\\LV_TS_Report\\Support_Files\\LV_Impedence\\LV_Rep_GSRX3_Only.csv")
                LV_Rep_GSRX3.to_csv(outfile_LV_Rep_GSRX3,index=None)

                # Search 'Ch1_Impedance','Ch2_Impedance','Ch3_Impedance' impedence FAILED
                LV_Impedence_FAIL_GSRX3  =    LV_Rep_GSRX3[(LV_Rep_GSRX3.Ch1_Impedance <Low_Threshold_Imp_GSRX3)|(LV_Rep_GSRX3.Ch1_Impedance >High_Threshold_Imp_GSRX3)|
                                                     (LV_Rep_GSRX3.Ch2_Impedance <Low_Threshold_Imp_GSRX3)|(LV_Rep_GSRX3.Ch2_Impedance >High_Threshold_Imp_GSRX3)|
                                                     (LV_Rep_GSRX3.Ch3_Impedance <Low_Threshold_Imp_GSRX3)|(LV_Rep_GSRX3.Ch3_Impedance >High_Threshold_Imp_GSRX3)]                                                                            

                LV_Impedence_PASS_GSRX3  =    LV_Rep_GSRX3[(LV_Rep_GSRX3.Ch1_Impedance >=Low_Threshold_Imp_GSRX3)&(LV_Rep_GSRX3.Ch1_Impedance <=High_Threshold_Imp_GSRX3)&
                                                     (LV_Rep_GSRX3.Ch2_Impedance >=Low_Threshold_Imp_GSRX3)&(LV_Rep_GSRX3.Ch2_Impedance <=High_Threshold_Imp_GSRX3)&
                                                     (LV_Rep_GSRX3.Ch3_Impedance >=Low_Threshold_Imp_GSRX3)&(LV_Rep_GSRX3.Ch3_Impedance <=High_Threshold_Imp_GSRX3)]


                outfile_GSR_X3_LV_Impedence_FAIL_Report_With_Duplicated =("C:\\LV_TS_Report\\Support_Files\\LV_Impedence\\GSR_X3_LV_Impedence_FAIL_Report_With Duplicated.csv")
                LV_Impedence_FAIL_GSRX3.to_csv(outfile_GSR_X3_LV_Impedence_FAIL_Report_With_Duplicated,index=None)

                outfile_GSR_X3_LV_Impedence_PASS_Report_With_Duplicated =("C:\\LV_TS_Report\\Support_Files\\LV_Impedence\\GSR_X3_LV_Impedence_PASS_Report_With Duplicated.csv")
                LV_Impedence_PASS_GSRX3.to_csv(outfile_GSR_X3_LV_Impedence_PASS_Report_With_Duplicated,index=None)
                

                # Filtering GSR-X3 
                LV_Impedence_FAIL_GSR_X3 = LV_Impedence_FAIL_GSRX3.loc[:,
                                           ['Line','Station','Line_Station_Combined','CaseSN','DeviceType','Deployment_Time', 'LastScanTime',
                                            'BattVoltage','Ch1Impedence', 'Ch2Impedence','Ch3Impedence', 'LineViewerName',
                                            'Latitude','Longitude','Altitude','QC_Comments','Ch1_Impedance','Ch2_Impedance','Ch3_Impedance']]
                LV_Impedence_PASS_GSRX3  = LV_Impedence_PASS_GSRX3.loc[:,
                                           ['Line','Station','Line_Station_Combined','CaseSN','DeviceType','Deployment_Time', 'LastScanTime',
                                            'BattVoltage','Ch1Impedence', 'Ch2Impedence','Ch3Impedence', 'LineViewerName',
                                            'Latitude','Longitude','Altitude','QC_Comments']]
                


                LV_Impedence_FAIL_GSR_X3   = LV_Impedence_FAIL_GSR_X3[LV_Impedence_FAIL_GSR_X3.DeviceType == ' GSR X3']
                LV_Impedence_PASS_GSRX3    = LV_Impedence_PASS_GSRX3 [LV_Impedence_PASS_GSRX3.DeviceType == ' GSR X3']

                LV_Impedence_FAIL_GSR_X3   = LV_Impedence_FAIL_GSR_X3.drop_duplicates(['Line_Station_Combined'],keep='last')
                LV_Impedence_FAIL_GSR_X3   = LV_Impedence_FAIL_GSR_X3.reset_index(drop=True)
                
                LV_Impedence_PASS_GSRX3    = LV_Impedence_PASS_GSRX3.drop_duplicates(['Line_Station_Combined'],keep='last')
                LV_Impedence_PASS_GSRX3    = LV_Impedence_PASS_GSRX3.reset_index(drop=True)
                
                LV_Impedence_FAIL_GSR_X3_M = pd.merge(LV_Impedence_FAIL_GSR_X3, LV_Impedence_PASS_GSRX3, how='left', on=['Line_Station_Combined','Line_Station_Combined'])
                LV_Impedence_FAIL_GSR_X3_M.drop(columns = ['Line_y','Station_y','DeviceType_y','Deployment_Time_y',
                                                           'Latitude_y','Longitude_y','Altitude_y','QC_Comments_y'],axis=1,inplace=True)
                LV_Impedence_FAIL_GSR_X3_M["QC_Comments_x"] = LV_Impedence_FAIL_GSR_X3_M.shape[0]*["Right- After_QC"]


                outfile_LV_Impedence_FAIL_GSR_X3_M =("C:\\LV_TS_Report\\Support_Files\\LV_Impedence\\GSR_X3_LV_Impedence_FAIL_Report.csv")
                LV_Impedence_FAIL_GSR_X3_M.to_csv(outfile_LV_Impedence_FAIL_GSR_X3_M,index=None)

                outfile_LV_Impedence_PASS_GSRX3 =("C:\\LV_TS_Report\\Support_Files\\LV_Impedence\\GSR_X3_LV_Impedence_PASS_Report.csv")
                LV_Impedence_PASS_GSRX3.to_csv(outfile_LV_Impedence_PASS_GSRX3,index=None)


                ### Client Statistics Report GSR_X3- Impedence
                GSR_X3_Impedence_RNG_Two       = LV_Impedence_PASS_GSRX3.CaseSN.count()
                GSR_X3_Impedence_Fail_Client   = LV_Impedence_FAIL_GSR_X3_M[(LV_Impedence_FAIL_GSR_X3_M.CaseSN_y.isnull())]

                GSR_X3_Impedence_RNG_Zero      = GSR_X3_Impedence_Fail_Client[(GSR_X3_Impedence_Fail_Client.Ch1_Impedance >= 66666666)|
                                                (GSR_X3_Impedence_Fail_Client.Ch2_Impedance >= 66666666)|
                                                (GSR_X3_Impedence_Fail_Client.Ch3_Impedance >= 66666666)].count()
                GSR_X3_Impedence_RNG_Zero      = GSR_X3_Impedence_RNG_Zero['Line_Station_Combined']


                GSR_X3_Impedence_RNG_One       = GSR_X3_Impedence_Fail_Client[(GSR_X3_Impedence_Fail_Client.Ch1_Impedance < Low_Threshold_Imp_GSRX3)|
                                                (GSR_X3_Impedence_Fail_Client.Ch2_Impedance < Low_Threshold_Imp_GSRX3)|
                                                (GSR_X3_Impedence_Fail_Client.Ch3_Impedance < Low_Threshold_Imp_GSRX3)]

                GSR_X3_Impedence_RNG_One       = GSR_X3_Impedence_RNG_One[(GSR_X3_Impedence_RNG_One.Ch1_Impedance < 66666666)&
                                                (GSR_X3_Impedence_RNG_One.Ch2_Impedance < 66666666)&
                                                (GSR_X3_Impedence_RNG_One.Ch3_Impedance < 66666666)].count()

                GSR_X3_Impedence_RNG_One       = GSR_X3_Impedence_RNG_One['Line_Station_Combined']

                GSR_X3_Impedence_RNG_Three     = GSR_X3_Impedence_Fail_Client[(GSR_X3_Impedence_Fail_Client.Ch1_Impedance > High_Threshold_Imp_GSRX3)|
                                                (GSR_X3_Impedence_Fail_Client.Ch2_Impedance > High_Threshold_Imp_GSRX3)|
                                                (GSR_X3_Impedence_Fail_Client.Ch3_Impedance > High_Threshold_Imp_GSRX3)]

                GSR_X3_Impedence_RNG_Three     = GSR_X3_Impedence_RNG_Three[(GSR_X3_Impedence_RNG_Three.Ch1_Impedance < 66666666)&
                                                (GSR_X3_Impedence_RNG_Three.Ch2_Impedance < 66666666)&
                                                (GSR_X3_Impedence_RNG_Three.Ch3_Impedance < 66666666)]
                
                GSR_X3_Impedence_RNG_Three     = GSR_X3_Impedence_RNG_Three[(GSR_X3_Impedence_RNG_Three.Ch1_Impedance > Low_Threshold_Imp_GSRX3)&
                                                (GSR_X3_Impedence_RNG_Three.Ch2_Impedance > Low_Threshold_Imp_GSRX3)&
                                                (GSR_X3_Impedence_RNG_Three.Ch3_Impedance > Low_Threshold_Imp_GSRX3)].count()                
                                                 
                GSR_X3_Impedence_RNG_Three      = GSR_X3_Impedence_RNG_Three['Line_Station_Combined']
                Int_Low_Threshold_Imp_GSR_3C    = int(Low_Threshold_Imp_GSRX3)
                Int_High_Threshold_Imp_GSR_3C   = int(High_Threshold_Imp_GSRX3)
                
                Heading_GSR_X3_Impedence_RNG_One   = 'Less Than ' + str(Int_Low_Threshold_Imp_GSR_3C)
                Heading_GSR_X3_Impedence_RNG_Two   = str(Int_Low_Threshold_Imp_GSR_3C) + ' - ' + str(Int_High_Threshold_Imp_GSR_3C)
                Heading_GSR_X3_Impedence_RNG_Three = str(Int_High_Threshold_Imp_GSR_3C) + ' Up'
                

                GSR_X3_LV_QC_Impedence_Stat_Client = pd.DataFrame({'0':[GSR_X3_Impedence_RNG_Zero],Heading_GSR_X3_Impedence_RNG_One:[GSR_X3_Impedence_RNG_One],
                                                   Heading_GSR_X3_Impedence_RNG_Two:[GSR_X3_Impedence_RNG_Two],Heading_GSR_X3_Impedence_RNG_Three:[GSR_X3_Impedence_RNG_Three]},index=None)                                  
                                                   
                GSR_X3_LV_QC_Impedence_Stat_Client = GSR_X3_LV_QC_Impedence_Stat_Client.T
                GSR_X3_LV_QC_Impedence_Stat_Client = GSR_X3_LV_QC_Impedence_Stat_Client.reset_index(drop=False)
                
                GSR_X3_LV_QC_Impedence_Stat_Client.rename(columns = {'index':'Impedance (Ohms)', 0:'Number of Receivers'},inplace = True)
                GSR_X3_LV_QC_Impedence_Stat_Client = GSR_X3_LV_QC_Impedence_Stat_Client.reset_index(drop=True)
                Comments = ['OPEN', 'LOW', 'IMPEDENCE RANGE', 'HIGH']
                Index    = [1,2,3,4]
                GSR_X3_LV_QC_Impedence_Stat_Client['Comments'] = Comments
                GSR_X3_LV_QC_Impedence_Stat_Client['Index']    = Index
                GSR_X3_LV_QC_Impedence_Stat_Client = GSR_X3_LV_QC_Impedence_Stat_Client.loc[:,
                                           ['Index','Impedance (Ohms)','Number of Receivers','Comments']]   

                outfile_GSR_X3_LV_QC_Impedence_Stat_Client =("C:\\LV_TS_Report\\GSR_X3_LV_QC_Impedence_Stat_Client.csv")
                GSR_X3_LV_QC_Impedence_Stat_Client.to_csv(outfile_GSR_X3_LV_QC_Impedence_Stat_Client,index=False)


                # Filtering GSR-X1
                LV_Impedence_FAIL_GSRX1  =    LV_Rep_GSRX1[(LV_Rep_GSRX1.Ch1_Impedance <Low_Threshold_Imp_GSR_1C)|(LV_Rep_GSRX1.Ch1_Impedance >High_Threshold_Imp_GSR_1C)]

                LV_Impedence_PASS_GSRX1  =    LV_Rep_GSRX1[(LV_Rep_GSRX1.Ch1_Impedance >=Low_Threshold_Imp_GSR_1C)&(LV_Rep_GSRX1.Ch1_Impedance <=High_Threshold_Imp_GSR_1C)]


                outfile_GSR_X1_LV_Impedence_FAIL_Report_With_Duplicated =("C:\\LV_TS_Report\\Support_Files\\LV_Impedence\\GSR_X1_LV_Impedence_FAIL_Report_With Duplicated.csv")
                LV_Impedence_FAIL_GSRX1.to_csv(outfile_GSR_X1_LV_Impedence_FAIL_Report_With_Duplicated,index=None)

                outfile_GSR_X1_LV_Impedence_PASS_Report_With_Duplicated =("C:\\LV_TS_Report\\Support_Files\\LV_Impedence\\GSR_X1_LV_Impedence_PASS_Report_With Duplicated.csv")
                LV_Impedence_PASS_GSRX1.to_csv(outfile_GSR_X1_LV_Impedence_PASS_Report_With_Duplicated,index=None)

                LV_Impedence_FAIL_GSR_X1 = LV_Impedence_FAIL_GSRX1.loc[:,
                                           ['Line','Station','Line_Station_Combined','CaseSN','DeviceType','Deployment_Time','LastScanTime',
                                            'BattVoltage','Ch1Impedence','LineViewerName','Latitude','Longitude','Altitude',
                                            'QC_Comments','Ch1_Impedance']]

                LV_Impedence_PASS_GSRX1 = LV_Impedence_PASS_GSRX1.loc[:,
                                           ['Line','Station','Line_Station_Combined','CaseSN','DeviceType','Deployment_Time','LastScanTime',
                                            'BattVoltage','Ch1Impedence','LineViewerName','Latitude','Longitude','Altitude','QC_Comments']]        
                
                LV_Impedence_FAIL_GSR_X1   = LV_Impedence_FAIL_GSR_X1[(LV_Impedence_FAIL_GSR_X1.DeviceType == ' 1-C GSR')|
                                            (LV_Impedence_FAIL_GSR_X1.DeviceType == ' 1-C GSX')]
                LV_Impedence_FAIL_GSR_X1   = LV_Impedence_FAIL_GSR_X1.drop_duplicates(['Line_Station_Combined'],keep='last')
                LV_Impedence_FAIL_GSR_X1   = LV_Impedence_FAIL_GSR_X1.reset_index(drop=True)
                LV_Impedence_PASS_GSRX1    = LV_Impedence_PASS_GSRX1.drop_duplicates(['Line_Station_Combined'],keep='last')
                LV_Impedence_PASS_GSRX1    = LV_Impedence_PASS_GSRX1.reset_index(drop=True)
                
                LV_Impedence_FAIL_GSR_X1_M = pd.merge(LV_Impedence_FAIL_GSR_X1, LV_Impedence_PASS_GSRX1,
                                             how='left', on=['Line_Station_Combined','Line_Station_Combined'])
                
                LV_Impedence_FAIL_GSR_X1_M.drop(columns =  ['Line_y','Station_y','DeviceType_y','Deployment_Time_y',
                                                            'Latitude_y','Longitude_y','Altitude_y', 'QC_Comments_y'], axis=1, inplace=True)
                LV_Impedence_FAIL_GSR_X1_M["QC_Comments_x"] = LV_Impedence_FAIL_GSR_X1_M.shape[0]*["Right- After_QC"]


                outfile_LV_Impedence_FAIL_GSR_X1_M =("C:\\LV_TS_Report\\Support_Files\\LV_Impedence\\GSR_X1_LV_Impedence_FAIL_Report.csv")
                LV_Impedence_FAIL_GSR_X1_M.to_csv(outfile_LV_Impedence_FAIL_GSR_X1_M,index=None)

                outfile_LV_Impedence_PASS_GSRX1 =("C:\\LV_TS_Report\\Support_Files\\LV_Impedence\\GSR_X1_LV_Impedence_PASS_Report.csv")
                LV_Impedence_PASS_GSRX1.to_csv(outfile_LV_Impedence_PASS_GSRX1,index=None)


                ### Client Statistics Report GSR_X1- Impedence

                GSR_X1_Impedence_RNG_Two       = LV_Impedence_PASS_GSRX1.CaseSN.count()
                GSR_X1_Impedence_Fail_Client   = LV_Impedence_FAIL_GSR_X1_M[(LV_Impedence_FAIL_GSR_X1_M.CaseSN_y.isnull())]
                GSR_X1_Impedence_RNG_Zero      = GSR_X1_Impedence_Fail_Client[(GSR_X1_Impedence_Fail_Client.Ch1_Impedance >= 66666666)].count()
                GSR_X1_Impedence_RNG_Zero      = GSR_X1_Impedence_RNG_Zero['Line_Station_Combined']
                GSR_X1_Impedence_RNG_One       = GSR_X1_Impedence_Fail_Client[(GSR_X1_Impedence_Fail_Client.Ch1_Impedance < Low_Threshold_Imp_GSR_1C)].count()
                GSR_X1_Impedence_RNG_One       = GSR_X1_Impedence_RNG_One['Line_Station_Combined']
                GSR_X1_Impedence_RNG_Three     = GSR_X1_Impedence_Fail_Client[((GSR_X1_Impedence_Fail_Client.Ch1_Impedance > High_Threshold_Imp_GSR_1C)&
                                                (GSR_X1_Impedence_Fail_Client.Ch1_Impedance < 66666666))].count()
                GSR_X1_Impedence_RNG_Three     = GSR_X1_Impedence_RNG_Three['Line_Station_Combined']
                Int_Low_Threshold_Imp_GSR_1C    = int(Low_Threshold_Imp_GSR_1C)
                Int_High_Threshold_Imp_GSR_1C   = int(High_Threshold_Imp_GSR_1C)
                
                Heading_GSR_X1_Impedence_RNG_One   = 'Less Than ' + str(Int_Low_Threshold_Imp_GSR_1C)
                Heading_GSR_X1_Impedence_RNG_Two   = str(Int_Low_Threshold_Imp_GSR_1C) + ' - ' + str(Int_High_Threshold_Imp_GSR_1C)
                Heading_GSR_X1_Impedence_RNG_Three = str(Int_High_Threshold_Imp_GSR_1C) + ' Up'
                GSR_X1_LV_QC_Impedence_Stat_Client = pd.DataFrame({'0':[GSR_X1_Impedence_RNG_Zero], Heading_GSR_X1_Impedence_RNG_One:[GSR_X1_Impedence_RNG_One],
                                                   Heading_GSR_X1_Impedence_RNG_Two:[GSR_X1_Impedence_RNG_Two],Heading_GSR_X1_Impedence_RNG_Three:[GSR_X1_Impedence_RNG_Three]},index=None)                                  
                GSR_X1_LV_QC_Impedence_Stat_Client = GSR_X1_LV_QC_Impedence_Stat_Client.T
                GSR_X1_LV_QC_Impedence_Stat_Client = GSR_X1_LV_QC_Impedence_Stat_Client.reset_index(drop=False)        
                GSR_X1_LV_QC_Impedence_Stat_Client.rename(columns = {'index':'Impedance (Ohms)', 0:'Number of Receivers'},inplace = True)        
                GSR_X1_LV_QC_Impedence_Stat_Client = GSR_X1_LV_QC_Impedence_Stat_Client.reset_index(drop=True)
                Comments = ['OPEN', 'LOW', 'IMPEDENCE RANGE', 'HIGH']
                Index    = [1,2,3,4]
                GSR_X1_LV_QC_Impedence_Stat_Client['Comments'] = Comments
                GSR_X1_LV_QC_Impedence_Stat_Client['Index']    = Index
                GSR_X1_LV_QC_Impedence_Stat_Client = GSR_X1_LV_QC_Impedence_Stat_Client.loc[:,
                                           ['Index','Impedance (Ohms)','Number of Receivers','Comments']]        
                outfile_GSR_X1_LV_QC_Impedence_Stat_Client =("C:\\LV_TS_Report\\GSR_X1_LV_QC_Impedence_Stat_Client.csv")
                GSR_X1_LV_QC_Impedence_Stat_Client.to_csv(outfile_GSR_X1_LV_QC_Impedence_Stat_Client,index=False)


                ### Export Client Statistics_REPORT
                from datetime import datetime
                def get_LV_Daily_LV_QC_Stat_Client_datetime():
                    return JobName_FileName + datetime.now().strftime("%Y%m%d") + " Lineviewer Report" + ".xlsx"
                root = Tk()
                root.filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select File Name To Export Lineviewer QC Statistics For Client" ,
                                                                     filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
                if len(root.filename) >0:
                        LV_QC_Client             = get_LV_Daily_LV_QC_Stat_Client_datetime()
                        LV_QC_Client_path        = root.filename + LV_QC_Client
                        XLSX_writer_LV_QC_Client = pd.ExcelWriter(LV_QC_Client_path)
                        LV_QC_Batt_Stat_Client.to_excel(XLSX_writer_LV_QC_Client,'BattStat',index=False, startrow=25)
                        GSR_X3_LV_QC_Impedence_Stat_Client.to_excel(XLSX_writer_LV_QC_Client,'GSRX3Stat',index=False, startrow=25)
                        GSR_X1_LV_QC_Impedence_Stat_Client.to_excel(XLSX_writer_LV_QC_Client,'GSRX1Stat',index=False, startrow=25)

                        header_BattStat = '&L&G'+'&C&18 LineViewer GSR/GSX Battery Report' 
                        header_GSX3Imp  = '&L&G'+'&C&18 LineViewer  GSR X3 Impedence Report' 
                        header_GSX1Imp  = '&L&G'+'&C&18 LineViewer 1-C GSR/GSX Impedence Report'
                        footer = ('&CEAGLE CANADA SEISMIC SERVICES ULC' + '\n'
                                  + '6806 Railway Street SE Calgary, AB T2H 3A8' + '\n' +
                                    'Ph: (403) 263-7770  Fax: 403 263 7776 Web : www.eaglecanada.ca')

                        workbook             = XLSX_writer_LV_QC_Client.book
                        worksheet_BattStat   = XLSX_writer_LV_QC_Client.sheets['BattStat']
                        worksheet_GSRX3Stat  = XLSX_writer_LV_QC_Client.sheets['GSRX3Stat']
                        worksheet_GSRX1Stat  = XLSX_writer_LV_QC_Client.sheets['GSRX1Stat']

                        worksheet_BattStat.set_margins(0.4, 0.4, 1.6, 1.1)
                        worksheet_GSRX3Stat.set_margins(0.4, 0.4, 1.6, 1.1)
                        worksheet_GSRX1Stat.set_margins(0.4, 0.4, 1.6, 1.1)

                        worksheet_BattStat.set_header(header_BattStat,{'image_left':"eagle logo.jpg"})
                        worksheet_BattStat.set_footer(footer)
                        worksheet_GSRX3Stat.set_header(header_GSX3Imp,{'image_left':"eagle logo.jpg"})                
                        worksheet_GSRX3Stat.set_footer(footer)
                        worksheet_GSRX1Stat.set_header(header_GSX1Imp,{'image_left':"eagle logo.jpg"})                
                        worksheet_GSRX1Stat.set_footer(footer)

                        workbook.formats[0].set_align('center')
                        workbook.formats[0].set_font_size(11)
                        workbook.formats[0].set_bold(True)
                        workbook.formats[0].set_border(4)

                        worksheet_BattStat.print_area('A1:D37')
                        worksheet_BattStat.print_across()
                        worksheet_BattStat.fit_to_pages(1, 1)
                        worksheet_BattStat.set_paper(9)
                        worksheet_BattStat.set_start_page(1)
                        worksheet_BattStat.hide_gridlines(1)

                        worksheet_GSRX3Stat.print_area('A1:D37')
                        worksheet_GSRX3Stat.print_across()
                        worksheet_GSRX3Stat.fit_to_pages(1, 1)
                        worksheet_GSRX3Stat.set_paper(9)
                        worksheet_GSRX3Stat.set_start_page(1)
                        worksheet_GSRX3Stat.hide_gridlines(1)

                        worksheet_GSRX1Stat.print_area('A1:D37')
                        worksheet_GSRX1Stat.print_across()
                        worksheet_GSRX1Stat.fit_to_pages(1, 1)
                        worksheet_GSRX1Stat.set_paper(9)
                        worksheet_GSRX1Stat.set_start_page(1)
                        worksheet_GSRX1Stat.hide_gridlines(1)

                        worksheet_BattStat.set_column('A:A',15)
                        worksheet_BattStat.set_column('B:B', 35)
                        worksheet_BattStat.set_column('C:C', 25)
                        worksheet_BattStat.set_column('D:D', 15)

                        worksheet_GSRX3Stat.set_column('A:A',15)
                        worksheet_GSRX3Stat.set_column('B:B', 30)
                        worksheet_GSRX3Stat.set_column('C:C', 25)
                        worksheet_GSRX3Stat.set_column('D:D', 20)
                        
                        worksheet_GSRX1Stat.set_column('A:A',15)
                        worksheet_GSRX1Stat.set_column('B:B', 30)
                        worksheet_GSRX1Stat.set_column('C:C', 25)
                        worksheet_GSRX1Stat.set_column('D:D', 20)

                        chart_object = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
                        
                        ## Making Plot For Batt Volt
                        chart_object.add_series({ 
                            'name':       ['BattStat', 25, 2],   
                            'categories': ['BattStat', 26, 1, 35, 1],    
                            'values':     ['BattStat', 26, 2, 35, 2], 'gap': 50}) 
                        chart_object.set_title({'name': 'GSR Battery Volt vs Total Number Plot', 'name_font': {'size': 17, 'bold': True}})
                        chart_object.set_size({'width': 647, 'height': 430})                                  
                        chart_object.set_y_axis({
                            'name': 'Number of Receivers',
                            'name_font': {'size': 16, 'bold': True},
                            'num_font':  {'size': 10, 'bold': True, 'rotation': - 45},
                            'major_gridlines': {
                                          'visible': True,
                                          'line': {'width': 1.25, 'dash_type': 'dash'}},})
                        
                        
                        chart_object.set_x_axis({'name': 'Battery Voltage (V)', 'num_font':  {'size': 10, 'bold': True, 'rotation': - 45}, 'name_font': {'size': 15, 'bold': True},
                                          'major_gridlines': {
                                          'visible': True,
                                          'line': {'width': 1.25, 'dash_type': 'dash'}},})
                        chart_object.set_style(11)
                        chart_object.set_legend({'none': True}) 
                        worksheet_BattStat.insert_chart('A3', chart_object,  
                                        {'x_offset': 1, 'y_offset': 15})

                        ## Making Plot For GSX-3
                        GSRX3Statchart_object = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
                        GSRX3Statchart_object.add_series({ 
                            'name':       ['GSRX3Stat', 25, 2],   
                            'categories': ['GSRX3Stat', 26, 1, 29, 1],    
                            'values':     ['GSRX3Stat', 26, 2, 29, 2], 'gap': 50})

                        GSRX3Statchart_object.set_title({'name': 'GSR X3 Impedence vs Total Number Plot', 'name_font': {'size': 17, 'bold': True}})
                        GSRX3Statchart_object.set_size({'width': 647, 'height': 430})                                  
                        GSRX3Statchart_object.set_y_axis({
                            'name': 'Number of Receivers',
                            'name_font': {'size': 16, 'bold': True},
                            'num_font':  {'size': 10, 'bold': True, 'rotation': - 45},
                            'major_gridlines': {
                                          'visible': True,
                                          'line': {'width': 1.25, 'dash_type': 'dash'}},})                
                        
                        GSRX3Statchart_object.set_x_axis({'name': 'GSR X3 Impedance (Ohms)', 'num_font':  {'size': 10, 'bold': True, 'rotation': - 45}, 'name_font': {'size': 15, 'bold': True},
                                          'major_gridlines': {
                                          'visible': True,
                                          'line': {'width': 1.25, 'dash_type': 'dash'}},})
                        GSRX3Statchart_object.set_style(11)
                        GSRX3Statchart_object.set_legend({'none': True}) 
                        worksheet_GSRX3Stat.insert_chart('A3', GSRX3Statchart_object,  
                                        {'x_offset': 1, 'y_offset': 10})

                        ## Making Plot For GSX-1
                        GSRX1Statchart_object = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
                        GSRX1Statchart_object.add_series({ 
                            'name':       ['GSRX1Stat', 25, 2],   
                            'categories': ['GSRX1Stat', 26, 1, 29, 1],    
                            'values':     ['GSRX1Stat', 26, 2, 29, 2], 'gap': 50})

                        GSRX1Statchart_object.set_title({'name': ' 1-C GSR/GSX Impedence vs Total Number Plot', 'name_font': {'size': 17, 'bold': True}})
                        GSRX1Statchart_object.set_size({'width': 647, 'height': 430})                
                        GSRX1Statchart_object.set_y_axis({
                            'name': 'Number of Receivers',
                            'name_font': {'size': 16, 'bold': True},
                            'num_font':  {'size': 10, 'bold': True, 'rotation': - 45},
                            'major_gridlines': {
                                          'visible': True,
                                          'line': {'width': 1.25, 'dash_type': 'dash'}},})
                                       
                        GSRX1Statchart_object.set_x_axis({'name': '1-C GSR/GSX Impedance (Ohms)', 'num_font':  {'size': 10, 'bold': True, 'rotation': - 45}, 'name_font': {'size': 15, 'bold': True},
                                          'major_gridlines': {
                                          'visible': True,
                                          'line': {'width': 1.25, 'dash_type': 'dash'}},})
                        GSRX1Statchart_object.set_style(11)
                        GSRX1Statchart_object.set_legend({'none': True}) 
                        worksheet_GSRX1Stat.insert_chart('A3', GSRX1Statchart_object,  
                                        {'x_offset': 1, 'y_offset': 10})


                        cell_format_1 = workbook.add_format({
                                                        'bold': True,
                                                        'text_wrap': True,
                                                        'valign': 'top'})
                        cell_format_1.set_align('left')
                        cell_format_1.set_font_size(12)
                        worksheet_BattStat.merge_range('A1:B1', JobName, cell_format_1)
                        worksheet_BattStat.merge_range('A2:B2', ClientName, cell_format_1)
                        worksheet_BattStat.merge_range('C1:D1', CrewName, cell_format_1)
                        worksheet_BattStat.merge_range('C2:D2', PreparedDate, cell_format_1)

                        worksheet_GSRX3Stat.merge_range('A1:B1', JobName, cell_format_1)
                        worksheet_GSRX3Stat.merge_range('A2:B2', ClientName, cell_format_1)
                        worksheet_GSRX3Stat.merge_range('C1:D1', CrewName, cell_format_1)
                        worksheet_GSRX3Stat.merge_range('C2:D2', PreparedDate, cell_format_1)
                        
                        worksheet_GSRX1Stat.merge_range('A1:B1', JobName, cell_format_1)
                        worksheet_GSRX1Stat.merge_range('A2:B2', ClientName, cell_format_1)
                        worksheet_GSRX1Stat.merge_range('C1:D1', CrewName, cell_format_1)
                        worksheet_GSRX1Stat.merge_range('C2:D2', PreparedDate, cell_format_1)
                       
                        XLSX_writer_LV_QC_Client.save()
                        XLSX_writer_LV_QC_Client.close()
                        tkinter.messagebox.showinfo("BoxScriptSummary Report Export Message","Lineviewer QC Statistics Report Saved as Excel")
                        root.destroy()
                        
                else:
                        tkinter.messagebox.showinfo("Lineviewer QC Statistics Report Export Message","Please Select Lineviewer QC Statistics Report File Name")


        else:
            tkinter.messagebox.showinfo("LineViewer CSV File Import Message","Please Select LineViewer CSV File Folder")
window.Close()

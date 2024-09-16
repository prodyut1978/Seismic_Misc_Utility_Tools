import os
import pandas as pd
import glob
import datetime
import csv
import openpyxl
import numpy as np
from PyPDF2 import PdfFileMerger, PdfMerger
import matplotlib.pyplot as plt
import datetime
from datetime import datetime
import math
import tkinter.ttk as ttk
from tkinter.filedialog import asksaveasfile
from tkinter.filedialog import askopenfilenames
from tkinter import simpledialog
from tkinter import*
import tkinter.messagebox
from tkinter import filedialog
import sys
sys.setrecursionlimit(15000)

GPS_Batt_QC_path = r'C:\GPS_Batt_QC'
if not os.path.exists(GPS_Batt_QC_path):
    os.makedirs(GPS_Batt_QC_path)

GPS_Batt_Support_path = r'C:\GPS_Batt_QC\Support_Files'
if not os.path.exists(GPS_Batt_Support_path):
    os.makedirs(GPS_Batt_Support_path)

GPS_Batt_DetailTemp_path = r'C:\GPS_Batt_QC\Support_Files\DetailTempFiles'
if not os.path.exists(GPS_Batt_DetailTemp_path):
    os.makedirs(GPS_Batt_DetailTemp_path)
DetailTempfilelist = [ f for f in os.listdir(GPS_Batt_DetailTemp_path) if f.endswith(".pdf") ]
for f in DetailTempfilelist:
    os.remove(os.path.join(GPS_Batt_DetailTemp_path, f))

##Reading GPS Batt Files
root = Tk()
root.directory = filedialog.askdirectory(parent=root,initialdir="/path/to/start",title='Please Select The GSR GPS Batt Volt directory')
Length_Directory  =  len(root.directory)
if Length_Directory >0:
    os.chdir(root.directory)
    fileList_Initial = sorted(glob.glob("*.TXT"))
    Length_fileList_Initial = len(fileList_Initial)+1
    Recursion_Limit = 50*(Length_fileList_Initial)+100
    root.destroy()
    dfList   = []
    for filename in fileList_Initial:
        sys.setrecursionlimit(Recursion_Limit)
        df = pd.read_csv(filename, sep=',', low_memory=False, skiprows=(0,1),header=None, encoding = 'unicode_escape')
        df1          = pd.read_csv(filename, sep=',', low_memory=False, header=None, nrows=1, encoding = 'unicode_escape')
        BattNumber   = df1.loc[:,0]
        BattNumber   = BattNumber[0]
        Batt_Number  = BattNumber[30:48]
        df           = df.iloc[:,:]
        Time_UTC     = df.loc[:,0]
        Batt_Volt    = (df.loc[:,13])
        Batt_Temp    = (df.loc[:,14])
        NodeNumber   = (filename[0:-4])
        column_names = [Time_UTC, Batt_Volt, Batt_Temp]
        catdf        = pd.concat (column_names,axis=1,ignore_index =True)
        catdf.rename(columns ={0:'Deployment_Duration_UTC',
                              1:'BatteryVoltage',2:'BatteryTemp'},inplace = True)    
        catdf  = catdf[pd.to_numeric(catdf.BatteryVoltage, errors='coerce').notnull()]
        catdf['BatteryVoltage']         = (catdf.loc[:,['BatteryVoltage']]).astype(float)
        catdf["Node_Number"] = catdf.shape[0]*[NodeNumber]
        catdf["Batt_Number"] = catdf.shape[0]*[Batt_Number]
        catdf  = catdf[pd.to_numeric(catdf.BatteryVoltage, errors='coerce').notnull()]    
        catdf  = catdf[(catdf.BatteryVoltage > 0)]
        catdf  = pd.DataFrame(catdf)
        catdf['Node_Number']             = (catdf.loc[:,['Node_Number']]).astype(int)
        catdf['Batt_Number']             = (catdf.loc[:,['Batt_Number']]).astype(int)
        catdf['BatteryVoltage']          = (catdf.loc[:,['BatteryVoltage']]).astype(float)
        catdf['BatteryVoltage']          =  catdf['BatteryVoltage'] *0.001
        catdf['Deployment_Duration_UTC'] = pd.to_datetime(catdf.Deployment_Duration_UTC)
        Batt_Voly_DF    = pd.DataFrame(catdf)
        Batt_Voly_DF    = Batt_Voly_DF.reset_index(drop=True)
        Batt_Voly_DF['DuplicatedEntries'] = Batt_Voly_DF.sort_values(by =['Node_Number']).duplicated(['Deployment_Duration_UTC'],keep='last')
        Batt_Voly_DF    = Batt_Voly_DF.loc[Batt_Voly_DF.DuplicatedEntries == False]
        Batt_Voly_DF    = Batt_Voly_DF.reset_index(drop=True)    
        Batt_Volt_MIN   = round(Batt_Voly_DF['BatteryVoltage'].min(),1)
        Batt_Volt_MAX   = round(Batt_Voly_DF['BatteryVoltage'].max(),1)
        Batt_Volt_DROP  = (Batt_Volt_MAX - Batt_Volt_MIN)
        Batt_Volt_DROP_PERCENT   =   round(100*(Batt_Volt_DROP/Batt_Volt_MAX),1)
        Batt_Volt_START       =   Batt_Voly_DF ['Deployment_Duration_UTC'].min()
        Batt_Volt_END         =   Batt_Voly_DF ['Deployment_Duration_UTC'].max()
        Batt_Volt_DURATION    =  ( Batt_Volt_END - Batt_Volt_START).days
        BattDF_Plot = Batt_Voly_DF.plot(x='Deployment_Duration_UTC',
                       y=['BatteryVoltage','BatteryTemp'],
                       layout=None,  subplots=True, 
                       title=(" Battery #: " + str(Batt_Number) + '    ' + 
                              " Unit #: " + str(NodeNumber)+ '    ' + " Volt Drop (max): " + str(Batt_Volt_DROP_PERCENT) + '%' ),
                                legend=True, ylim=None, fontsize=4)
        plt.xlabel("Deployment Time Samples" + " ( Duration : " + str (Batt_Volt_DURATION) + " Days )" , fontsize=6)
        BattDF_Plot[0].set_ylabel('Battery Voltage (Volts)', fontsize=6)
        BattDF_Plot[1].set_ylabel('Battery Temp (°C)', fontsize=6)
        plt.rcParams.update({'figure.max_open_warning': 0})                
        def get_Node_Number():
            return "Node_Number -" + NodeNumber + ".pdf"
        Plt_get_Node_Number = get_Node_Number()
        Plt_get_Node_path = "C:\\GPS_Batt_QC\\Support_Files\\DetailTempFiles\\" + Plt_get_Node_Number
        plt.savefig((Plt_get_Node_path),dpi=10,orientation='portrait',
                    bbox_inches='tight',metadata={'Title': 'Nodes Battery Discharge Curve'})
        dfList.append(catdf)        

    concatDf                = pd.concat(dfList,axis=0)
    GPS_Batt_Rep            = pd.DataFrame(concatDf)
    GPS_Batt_Rep            = GPS_Batt_Rep.reset_index(drop=True)
    GPS_Batt_Rep['DuplicatedEntries'] = GPS_Batt_Rep.sort_values(by =['Node_Number']).duplicated(['Deployment_Duration_UTC'],keep='last')
    GPS_Batt_Rep = GPS_Batt_Rep.loc[GPS_Batt_Rep.DuplicatedEntries == False]
    GPS_Batt_Rep = GPS_Batt_Rep.reset_index(drop=True)

    ## Generating Batt Volt Statistics Report
    BattVoltMean           =   round(GPS_Batt_Rep.groupby('Node_Number').BatteryVoltage.mean(),1)
    BattVoltMin            =   round(GPS_Batt_Rep.groupby('Node_Number').BatteryVoltage.min(),1)
    BattVoltEnd            =   round(GPS_Batt_Rep.groupby('Node_Number').BatteryVoltage.nth(-1),1)
    BattVoltMax            =   round(GPS_Batt_Rep.groupby('Node_Number').BatteryVoltage.max(),1)
    BattVoltDrop           =   round((BattVoltMax - BattVoltMin),1)
    BattVoltDropPercent    =   round(100*(BattVoltDrop/BattVoltMax),1)
    DeployedDayStart       =   GPS_Batt_Rep.groupby('Node_Number').Deployment_Duration_UTC.nth(0)
    DeployedDayEnd         =   GPS_Batt_Rep.groupby('Node_Number').Deployment_Duration_UTC.nth(-1)
    DeploymentDuration     =  (DeployedDayEnd - DeployedDayStart).astype(str)
    DeploymentDuration     =  (DeploymentDuration.str.slice(0,9)) + ' Hours'
    BattSN                 =  (GPS_Batt_Rep.groupby('Node_Number').Batt_Number.unique()).astype(int)

    GPS_Batt_Stat          = [BattSN, DeployedDayStart, DeployedDayEnd, DeploymentDuration,
                              BattVoltMax, BattVoltMin, BattVoltEnd, BattVoltMean, BattVoltDrop,
                              BattVoltDropPercent]
    GPS_Batt_Stat          = pd.concat (GPS_Batt_Stat,axis=1,ignore_index =True)
    GPS_Batt_Stat.reset_index(inplace=True)
    GPS_Batt_Stat.rename(columns={0:'BattSN', 1:'DeployedDay',2:'PickupDay',
                                  3:'DeploymentDuration',4:'BattVoltMax',
                                  5:'BattVoltMin',6:'BattVoltEnd', 7:'BattVoltMean',
                                  8:'BattVoltDrop', 9:'% of BattVoltDrop'},inplace = True)
    GPS_Batt_Stat         = pd.DataFrame(GPS_Batt_Stat)


    ## Exporting Batt Voltage Stat File
    def get_BattStat_datetime():
        return "Batt Voltage Stat Report - " + datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx"
          
    get_BattStatReport   = get_BattStat_datetime()
    StatReport_path = "C:\\GPS_Batt_QC\\" + get_BattStatReport
    XLSX_writer = pd.ExcelWriter(StatReport_path)
    GPS_Batt_Stat.to_excel(XLSX_writer,'BattVoltStat',index=False)
    XLSX_writer.save()
    XLSX_writer.close()


    ## Exporting Detail Pdf Plot File
    iPlot = tkinter.messagebox.askyesno("GPS Battery Voltage Plot And Stat Export", "Do you Like To Save GPS Battery Voltage Plot And Stat Report?")
    if iPlot >0:
        ## Collecting All PDF Plot Files 
        os.chdir("C:\\GPS_Batt_QC\\Support_Files\\DetailTempFiles")
        files = [f for f in os.listdir('.') if os.path.isfile(f) and f.endswith('.pdf')]
        def get_Plot_datetime():
            return "GSR Batt Voltage Plot Detail Report - " + datetime.now().strftime("%Y%m%d-%H%M%S")+ ".pdf"
        def get_BattStat_datetime():
            return "GSR Batt Voltage Stat Report - " + datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx"
        
        Plot_datetime_apply    = get_Plot_datetime()
        Stat_datetime_apply    = get_BattStat_datetime()
        
        root = Tk()
        root.filename =  filedialog.asksaveasfilename(initialdir = "/", title = "Save Batt Voltage Plot File As PDF ", filetypes = (("PDF files","*.pdf"),("all files","*.*")))
        
        Plot_path = root.filename + Plot_datetime_apply
        Stat_path = root.filename + Stat_datetime_apply

        ## Making PDF For Plot Files Combined
        merger = PdfMerger()
        for pdf in files:
            merger.append(open(pdf, 'rb'))
        with open(Plot_path, "wb") as fout:
            merger.write(fout)
        merger.close()
        fout.close()

        ## Making Batt Stat Report    
        XLSX_writer = pd.ExcelWriter(Stat_path)
        GPS_Batt_Stat.to_excel(XLSX_writer,'BattVoltStat',index=False)
        workbook       = XLSX_writer.book
        worksheet      = XLSX_writer.sheets['BattVoltStat']
        worksheet.set_landscape()
        worksheet.set_margins(0.6, 0.6, 1.6, 1.1)
        worksheet.print_across()                                  
        worksheet.set_paper(9)
        worksheet.set_start_page(1)
        workbook.formats[0].set_align('center')
        workbook.formats[0].set_font_size(11)
        workbook.formats[0].set_bold(True)
        workbook.formats[0].set_border(1)
        worksheet.set_column(0, 10, 21)   
        XLSX_writer.save()
        XLSX_writer.close()
        tkinter.messagebox.showinfo("Batt Voltage Export Report"," Batt Voltage Plot Report Saved as PDF And Stat Report Saved As Excel")
        root.destroy()


else:
    tkinter.messagebox.showinfo("Batt Voltage Import Message","Please Select Batt Voltage File Folder")
    









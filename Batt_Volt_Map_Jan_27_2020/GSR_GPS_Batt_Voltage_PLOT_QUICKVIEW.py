import os
import pandas as pd
import glob
import datetime
import csv
import openpyxl
import matplotlib.pyplot as plt
import numpy as np
from PyPDF2 import PdfFileMerger
import datetime
from datetime import datetime
import math
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

GPS_Batt_QuickTemp_path = r'C:\GPS_Batt_QC\Support_Files\QuickViewTempFiles'
if not os.path.exists(GPS_Batt_QuickTemp_path):
    os.makedirs(GPS_Batt_QuickTemp_path)
Quickfilelist = [ f for f in os.listdir(GPS_Batt_QuickTemp_path) if f.endswith(".pdf") ]
for f in Quickfilelist:
    os.remove(os.path.join(GPS_Batt_QuickTemp_path, f))

##Reading GPS Batt Files
root = Tk()
root.directory = filedialog.askdirectory(parent=root,initialdir="/path/to/start",title='Please select The GSR Batt Voltage GPS File directory')
Length_Directory  =  len(root.directory)
if Length_Directory >0:
    os.chdir(root.directory)
    fileList_Initial = sorted(glob.glob("*.TXT"))
    Length_fileList_Initial = len(fileList_Initial)+1
    Recursion_Limit = 50*(Length_fileList_Initial)+100
    root.destroy()

    def divide_chunks(l, n): 
        for i in range(0, len(l), n):  
            yield l[i:i + n]
    Count_PerPage = 16  
    fileListByChunk = list(divide_chunks(fileList_Initial, Count_PerPage))

    countLoop  = 0
    for x in list(fileListByChunk):
        countLoop = countLoop+1
        StringCountLoop = str(countLoop)
        dfList   = []
        for filename in list(x):
            sys.setrecursionlimit(Recursion_Limit)
            df = pd.read_csv(filename, sep=',', low_memory=False, skiprows=(0,1),header=None, encoding = 'unicode_escape')
            df1          = pd.read_csv(filename, sep=',', low_memory=False, header=None, nrows=1, encoding = 'unicode_escape')
            BattNumber   = df1.loc[:,0]
            BattNumber   = BattNumber[0]
            Batt_Number  = BattNumber[30:48]
            df           = df.iloc[:,:]
            Time_UTC     = df.loc[:,0]
            Batt_Volt    = (df.loc[:,13])
            NodeNumber   = (filename[0:-4])
            column_names = [Time_UTC, Batt_Volt]
            catdf        = pd.concat (column_names,axis=1,ignore_index =True)
            catdf.rename(columns ={0:'Deployment_Duration_UTC',
                                  1:'Battery_Voltage'},inplace = True)
            catdf  = catdf[pd.to_numeric(catdf.Battery_Voltage, errors='coerce').notnull()]
            catdf['Battery_Voltage']         = (catdf.loc[:,['Battery_Voltage']]).astype(float)
            catdf["Node_Number"] = catdf.shape[0]*[NodeNumber]
            catdf["Batt_Number"] = catdf.shape[0]*[Batt_Number]
            catdf  = catdf[pd.to_numeric(catdf.Battery_Voltage, errors='coerce').notnull()]    
            catdf  = catdf[(catdf.Battery_Voltage > 0)]
            catdf  = pd.DataFrame(catdf)
            catdf['Node_Number']             = (catdf.loc[:,['Node_Number']]).astype(int)
            catdf['Batt_Number']             = (catdf.loc[:,['Batt_Number']]).astype(int)
            catdf['Battery_Voltage']         = (catdf.loc[:,['Battery_Voltage']]).astype(float)
            catdf['Battery_Voltage']         =  catdf['Battery_Voltage'] *0.001
            catdf['Deployment_Duration_UTC'] = pd.to_datetime(catdf.Deployment_Duration_UTC)
            dfList.append(catdf)
            
        concatDf                = pd.concat(dfList,axis=0)
        GPS_Batt_Rep            = pd.DataFrame(concatDf)
        GPS_Batt_Rep            = GPS_Batt_Rep.reset_index(drop=True)

        GPS_Batt_Rep['DuplicatedEntries'] = GPS_Batt_Rep.sort_values(by =['Node_Number']).duplicated(['Deployment_Duration_UTC'],keep='last')
        GPS_Batt_Rep = GPS_Batt_Rep.loc[GPS_Batt_Rep.DuplicatedEntries == False]
        GPS_Batt_Rep = GPS_Batt_Rep.reset_index(drop=True)
        
        GPS_Batt_Rep_Glance  = GPS_Batt_Rep.loc[:,['Batt_Number','Battery_Voltage']]
        grouped=GPS_Batt_Rep_Glance.groupby('Batt_Number')
        LengthGroup = len(grouped)
        Ncolumns = 4
        Nrows= int(math.ceil(len(grouped)/Ncolumns))
        fig,axs=plt.subplots(nrows = Nrows, ncols=Ncolumns,constrained_layout = True)    
        plt.rcParams.update({'figure.max_open_warning': 0})
        fig.suptitle("Quick View Of Multiple GSR Battery Voltage Plot")

        for (name,GPS_Batt_Rep_Glance),ax in zip(grouped,axs.flat):
            GPS_Batt_Rep_Glance.groupby('Batt_Number').Battery_Voltage.plot.line(y='Battery_Voltage',ax=ax,
            sharex=False, sharey=True, legend=True, fontsize=5)
            ax.title.set_size(6)
            ax.set_ylabel('Batt Volts' ,fontsize=6)
            ax.set_xlabel('Time Samples',fontsize=6)
            ax.legend(fontsize=6)

        def get_Batt_Number():
            return "Batt_Voltage Plot Quick_View" + StringCountLoop + ".pdf"
        Plt_get_Batt_Number = get_Batt_Number()
        Plt_get_Batt_path = "C:\\GPS_Batt_QC\\Support_Files\\QuickViewTempFiles\\" + Plt_get_Batt_Number
        plt.savefig((Plt_get_Batt_path),dpi=10,orientation='portrait',
                        bbox_inches='tight',metadata={'Title': 'Nodes Battery Discharge Curve'})

    ## Exporting Detail Pdf Plot File
    os.chdir("C:\\GPS_Batt_QC\\Support_Files\\QuickViewTempFiles")
    files = [f for f in os.listdir('.') if os.path.isfile(f) and f.endswith('.pdf')]
    def get_Plot_datetime():
        return "Batt Voltage Plot QuickView Report - " + datetime.now().strftime("%Y%m%d-%H%M%S")+ ".pdf"
    Plot_datetime_apply = get_Plot_datetime()

    root = Tk()
    root.filename =  filedialog.asksaveasfilename(initialdir = "/", title = "Save Batt Voltage Plot File As PDF ", filetypes = (("PDF files","*.pdf"),("all files","*.*")))

    Plot_path = root.filename + Plot_datetime_apply

    merger = PdfFileMerger()
    for pdf in files:
        merger.append(pdf)
    with open(Plot_path, "wb") as fout:
        merger.write(fout)
    merger.close()
    tkinter.messagebox.showinfo("Batt Voltage Export Report"," Batt Voltage Plot Report Saved as PDF")
    root.destroy()
else:
    tkinter.messagebox.showinfo("Batt Voltage Import Message","Please Select Batt Voltage File Folder")

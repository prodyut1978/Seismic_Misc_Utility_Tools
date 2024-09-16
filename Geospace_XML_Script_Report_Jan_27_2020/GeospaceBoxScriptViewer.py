import os
import pandas as pd
import glob
import datetime
import openpyxl
import datetime
from datetime import datetime
import tkinter.ttk as ttk
from tkinter.filedialog import asksaveasfile
from tkinter.filedialog import askopenfilenames
from tkinter import*
import tkinter.messagebox
from tkinter import filedialog
import sys
from xml.dom import minidom
import codecs
import re
from PIL import Image


myImage = Image.open("eagle logo.jpg")
Vib_VibProductionReport_path = r'C:\XMLRestrictedFolder'
if not os.path.exists(Vib_VibProductionReport_path):
    os.makedirs(Vib_VibProductionReport_path)
myImage = myImage.save("C:\\XMLRestrictedFolder\\eagle logo.jpg")

root = Tk()
root.directory = filedialog.askdirectory(parent=root,initialdir="/path/to/start",title='Please Select Directory Where XML Script Located')
Length_Directory  =  len(root.directory)
if Length_Directory >0:
    os.chdir(root.directory)
    fileList = glob.glob("*_Script.xml")
    xmldoc = minidom.parse(fileList[0])
    root.destroy()
    ## Getting Channel Gain
    GainCodeCh1 = xmldoc.getElementsByTagName('GainCodeCh1')
    GainCodeCh1 = GainCodeCh1[0]
    GainCodeCh1 = GainCodeCh1.firstChild
    my_GainCodeCh1 = GainCodeCh1.data
    my_GainCodeCh1 = int(re.search(r'\d+', my_GainCodeCh1).group(0))
    my_GainCodeCh1 = str(my_GainCodeCh1) + ' DB'

    GainCodeCh2 = xmldoc.getElementsByTagName('GainCodeCh2')
    GainCodeCh2 = GainCodeCh2[0]
    GainCodeCh2 = GainCodeCh2.firstChild
    my_GainCodeCh2 = GainCodeCh2.data
    my_GainCodeCh2 = int(re.search(r'\d+', my_GainCodeCh2).group(0))
    my_GainCodeCh2 = str(my_GainCodeCh2) + ' DB'

    GainCodeCh3 = xmldoc.getElementsByTagName('GainCodeCh3')
    GainCodeCh3 = GainCodeCh3[0]
    GainCodeCh3 = GainCodeCh3.firstChild
    my_GainCodeCh3 = GainCodeCh3.data
    my_GainCodeCh3 = int(re.search(r'\d+', my_GainCodeCh3).group(0))
    my_GainCodeCh3 = str(my_GainCodeCh3) + ' DB'


    GainCodeCh4 = xmldoc.getElementsByTagName('GainCodeCh4')
    GainCodeCh4 = GainCodeCh4[0]
    GainCodeCh4 = GainCodeCh4.firstChild
    my_GainCodeCh4 = GainCodeCh4.data
    my_GainCodeCh4 = int(re.search(r'\d+', my_GainCodeCh4).group(0))
    my_GainCodeCh4 = str(my_GainCodeCh4) + ' DB'


    Gain_Vector = pd.DataFrame({'Gain Channel Number': ['CH-1', 'CH-2' ,'CH-3','CH-4'],
                                'Applied Channel Gain': [my_GainCodeCh1,my_GainCodeCh2,
                                                 my_GainCodeCh3, my_GainCodeCh4]})
    ## Getting Lowcut

    LowCutCode1 = xmldoc.getElementsByTagName('LowCutCode1')
    LowCutCode1 = LowCutCode1[0]
    LowCutCode1 = LowCutCode1.firstChild
    my_LowCutCode1 = LowCutCode1.data
    my_LowCutCode1 = int(list(filter(str.isdigit, my_LowCutCode1))[0])
    my_LowCutCode1 = str(my_LowCutCode1) + ' Hz'

    LowCutCode2 = xmldoc.getElementsByTagName('LowCutCode2')
    LowCutCode2 = LowCutCode2[0]
    LowCutCode2 = LowCutCode2.firstChild
    my_LowCutCode2 = LowCutCode2.data
    my_LowCutCode2 = int(list(filter(str.isdigit, my_LowCutCode2))[0])
    my_LowCutCode2 = str(my_LowCutCode2) + ' Hz'

    LowCutCode3 = xmldoc.getElementsByTagName('LowCutCode3')
    LowCutCode3 = LowCutCode3[0]
    LowCutCode3 = LowCutCode3.firstChild
    my_LowCutCode3 = LowCutCode3.data
    my_LowCutCode3 = int(list(filter(str.isdigit, my_LowCutCode3))[0])
    my_LowCutCode3 = str(my_LowCutCode3) + ' Hz'

    LowCutCode4 = xmldoc.getElementsByTagName('LowCutCode4')
    LowCutCode4 = LowCutCode4[0]
    LowCutCode4 = LowCutCode4.firstChild
    my_LowCutCode4 = LowCutCode4.data
    my_LowCutCode4 = int(list(filter(str.isdigit, my_LowCutCode4))[0])
    my_LowCutCode4 = str(my_LowCutCode4) + ' Hz'

    LowCut_Vector = pd.DataFrame({'Low Cut Channel Number': ['CH-1', 'CH-2' ,'CH-3','CH-4'],
                                'Applied Low Cut Filter': [my_LowCutCode1, my_LowCutCode2,
                                                   my_LowCutCode3, my_LowCutCode4]})


    ## Getting Alias Filter

    AliasCode = xmldoc.getElementsByTagName('AliasCode')
    AliasCode = AliasCode[0]
    AliasCode = AliasCode.firstChild
    my_AliasCode = AliasCode.data

    AliasVector = pd.DataFrame({'Alias Filter': [my_AliasCode]})

    ## Getting Sample Rate
    RateCode = xmldoc.getElementsByTagName('RateCode')
    RateCode = RateCode[0]
    RateCode = RateCode.firstChild
    my_RateCode = RateCode.data
    my_RateCode = int(re.search(r'\d+', my_RateCode).group(0))
    if my_RateCode == 1:
        my_RateCode = int(1/(my_RateCode))
    elif my_RateCode == 2:
        my_RateCode = (1/(my_RateCode))
    elif my_RateCode == 4:
        my_RateCode = (1/(my_RateCode))
    else:
        my_RateCode = int(1000*(1/(my_RateCode)))
        
    my_RateCode = str(my_RateCode) + ' ms'
    RateCodeVector = pd.DataFrame({'Sample Rate (ms)': [my_RateCode]})


    ## Getting GPS Mode
    GpsMode = xmldoc.getElementsByTagName('GpsMode')
    GpsMode = GpsMode[0]
    GpsMode = GpsMode.firstChild
    my_GpsMode = GpsMode.data

    GpsModeVector = pd.DataFrame({'GPS Mode': [my_GpsMode]})

    ## Getting Name And Loop Mode And Start date Time
    loopMode = xmldoc.getElementsByTagName('loopMode')
    loopMode = loopMode[0]
    loopMode = loopMode.firstChild
    my_loopMode = loopMode.data

    Name = xmldoc.getElementsByTagName('Name')
    Name = Name[0]
    Name = Name.firstChild
    my_Name = Name.data

    StartDateTime = xmldoc.getElementsByTagName('StartDateTime')
    StartDateTime = StartDateTime[0]
    StartDateTime = StartDateTime.firstChild
    my_StartDateTime = StartDateTime.data

    ## Getting Time Record On

    IntervalCount = xmldoc.getElementsByTagName('IntervalCount')
    IntervalCount = IntervalCount[0]
    IntervalCount = IntervalCount.firstChild
    my_IntervalCount = IntervalCount.data

    TimeUnit = xmldoc.getElementsByTagName('TimeUnit')
    TimeUnit = TimeUnit[0]
    TimeUnit = TimeUnit.firstChild
    my_TimeUnit = TimeUnit.data

    IntervalCount1 = xmldoc.getElementsByTagName('IntervalCount')
    IntervalCount1 = IntervalCount1[1]
    IntervalCount1 = IntervalCount1.firstChild
    my_IntervalCount1 = IntervalCount1.data


    TimeUnit1 = xmldoc.getElementsByTagName('TimeUnit')
    TimeUnit1 = TimeUnit1[1]
    TimeUnit1 = TimeUnit1.firstChild
    my_TimeUnit1 = TimeUnit1.data

    CombinedVector = pd.DataFrame({'Script Name': [my_Name], 'Script Start Date And Time':[my_StartDateTime], 'Recording On Per Day': [my_IntervalCount + ' ' + my_TimeUnit],
                               'Recording Off Per Day': [my_IntervalCount1 + ' ' + my_TimeUnit1]})

    Combined_GPS_Alias_Rate = pd.DataFrame({'Alias Filter (Linear Or Minimum Phase)': [my_AliasCode], 'Sample Rate (ms)': [my_RateCode], 'GPS Mode (Cycle or Always On)': [my_GpsMode], 'Loop Mode': [my_loopMode] })


     #### Export Vib Production Summary Report

    def get_Script_Rep_datetime():
        return " - Geospace Box Script Summary -" + datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx"
    root = Tk()
    root.filename               = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select File Name For Geospace Box Script Summary Report" ,
                                  filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))

    if len(root.filename) >0:
        ScriptSummary   = get_Script_Rep_datetime()
        outfile_ScriptSummary = root.filename + ScriptSummary
        XLSX_writer = pd.ExcelWriter(outfile_ScriptSummary)

        CombinedVector.to_excel(XLSX_writer,'BoxScriptSummary',index=False, startrow=3)
        Combined_GPS_Alias_Rate.to_excel(XLSX_writer,'BoxScriptSummary',index=False, startrow=8)
        Gain_Vector.to_excel(XLSX_writer,'BoxScriptSummary',index=False, startrow=13, startcol=0)
        LowCut_Vector.to_excel(XLSX_writer,'BoxScriptSummary',index=False, startrow=13, startcol=2)

        workbook             = XLSX_writer.book
        worksheet_Front      = XLSX_writer.sheets['BoxScriptSummary']
        header1 = '&L&G'+'&CEagle Canada Seismic Services ULC' + '\n' + '6806 Railway Street SE' + '\n' + 'Calgary, AB T2H 3A8' + '\n' +  'Ph: (403) 263-7770'
        worksheet_Front.set_header(header1,{'image_left':"C:\\XMLRestrictedFolder\\eagle logo.jpg"})
        footer1 = ('&LDate : &D')
        worksheet_Front.set_footer(footer1)
        worksheet_Front.set_margins(0.4, 0.4, 1.6, 1.1)
        worksheet_Front.set_v_pagebreaks([4])
        worksheet_Front.print_area('A1:D34')
        worksheet_Front.print_across()
        worksheet_Front.fit_to_pages(1, 1)                                    
        worksheet_Front.set_paper(9)
        worksheet_Front.set_start_page(1)
        worksheet_Front.hide_gridlines(1)
        worksheet_Front.set_page_view()
        workbook.formats[0].set_align('center')
        workbook.formats[0].set_font_size(11)
        workbook.formats[0].set_bold(True)
        workbook.formats[0].set_border(2)

        worksheet_Front.set_column('A:A',35)
        worksheet_Front.set_column('B:B', 32)
        worksheet_Front.set_column('C:C', 28)
        worksheet_Front.set_column('D:D', 23)
        cell_format_Left = workbook.add_format({
                                                'bold': True,
                                                'text_wrap': True,
                                                'valign': 'top',
                                                'border': 1})
        cell_format_Left.set_align('left')
        cell_format_Left.set_font_size(12)

        worksheet_Front.merge_range('A22:B23', " Parameter Sheet Revision : ", cell_format_Left)
        worksheet_Front.merge_range('A24:B26', " Client Name : ", cell_format_Left)
        worksheet_Front.merge_range('A27:B29', " Client Rep Name : ", cell_format_Left)
        worksheet_Front.merge_range('A30:B32', " Signature Client Rep : ", cell_format_Left)

        worksheet_Front.merge_range('C22:D23', " Merge Operator Name : ", cell_format_Left)
        worksheet_Front.merge_range('C24:D26', " Program Or Project Name : ", cell_format_Left)
        worksheet_Front.merge_range('C27:D29', " Operation Supervisor Name: ", cell_format_Left)
        worksheet_Front.merge_range('C30:D32', " Signature Operation Supervisor : ", cell_format_Left)
        cell_format_Summary = workbook.add_format({
                                                'bold': True,
                                                'text_wrap': True,
                                                'valign': 'top'})
        cell_format_Summary.set_align('center')
        cell_format_Summary.set_font_size(16)
        cell_format_Summary.set_underline(1)
        worksheet_Front.merge_range('A1:D2', "Geospace Box Script Details", cell_format_Summary)
        worksheet_Front.set_page_view()
        XLSX_writer.save()
        XLSX_writer.close()
        tkinter.messagebox.showinfo("BoxScriptSummary Report Export Message","BoxScriptSummary Report Saved as Excel")
        root.destroy()
    else:
        tkinter.messagebox.showinfo("BoxScriptSummary Report Export Message","Please Select BoxScriptSummary Report File Name")
        
else:
    tkinter.messagebox.showinfo("BoxScript XML File Import Message","Please Select BoxScript XML File Folder")
    















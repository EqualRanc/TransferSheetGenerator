# -*- coding: utf-8 -*-
"""
Created on Thu Oct 14 09:59:15 2021
@author: EqualRanc
"""

import xlwings as xw
import pandas as pd
import os
import PySimpleGUI as sg
import datetime

values=[]

def process_xl(tabs, fullname):
    try:
        excel_app = xw.App(visible=False)
        filepath = os.path.expanduser(fullname)
        if not os.path.exists(filepath):
            return filepath
        excel_book = excel_app.books.open(filepath)
        sheetend = tabs + 1
        df = {}
        for number in range(1, sheetend):
            tabname = 'data ' + str(number)
            sheet = excel_book.sheets(tabname)
            df[tabname] = sheet[sheet.used_range.address].options(pd.DataFrame, index=False, header=True).value
            df[tabname] = df[tabname].fillna(0)
        excel_book.close()
        excel_app.quit()
        return df
    except:
        window['chemstatus'].print('Issue processing chemical database file, please check it.')
        window['assaystatus'].print('Issue processing chemical database file, please check it.')

def possplitter(s):
    letterz = s.rstrip('0123456789')
    numz = s[len(letterz):]
    return [letterz, numz]

def compare_checkboxlist(checkboxlist):
    for sublist in checkboxlist:
        checkboxtemp = [j for i in checkboxlist for j in i]
        if checkboxtemp[0] == checkboxlist[0][:4]:
            checkboxlist2 = checkboxtemp
        else:
            checkboxlist2 = checkboxlist
        return checkboxlist2

def ctablist():
    #Create list of data tab names where user entered a transfer volume greater than 0
    checkboxlist = []
    if values['-VNN-'] != '0':
        checkboxlist.append(["data 5", "data 6", "data 7", "data 8"])
    if values['-VN-'] != '0':
        checkboxlist.append(["data 9", "data 10", "data 11", "data 12"])
    if values['-VZZ-'] != '0':
        checkboxlist.append(["data 1", "data 2", "data 3"])
    if values['-VZ-'] != '0':
        checkboxlist.append(["data 4"])
    if values['-VO-'] != '0':
        checkboxlist.append(["data 13", "data 14"])
    if values['-VAT-'] != '0':
        checkboxlist.append(["data 15"])
    if values['-VAH-'] != '0':
        checkboxlist.append(["data 24"])
    if values['-VCA-'] != '0':
        checkboxlist.append(["data 16", "data 17", "data 18", "data 19", "data 20", "data 21", "data 22", "data 23"])
    if values['-VKT-'] != '0':
       checkboxlist.append(["data 25"]) 
    if values['-VBO-'] != '0':
        checkboxlist.append(["data 26", "data 27"])

    #Begin concatenating the desired chemical classes' data tabs and initializing the transfer sheet
    checkboxlist = [j for i in checkboxlist for j in i]
    if len(checkboxlist) == 0:
        window['chemstatus'].print("Please pick at least one chemical class (Chemical Tab).")
    chemfullsheet = pd.concat(oner, keys=checkboxlist)
    return chemfullsheet

def atablist():
            assaycheckboxlist = []
            if values['-ZZ2-'] == True:
                assaycheckboxlist.append(["data 1", "data 2", "data 3"])
            if values['-Z2-'] == True:
                assaycheckboxlist.append(["data 4"])
            if values['-NN2-'] == True:
                assaycheckboxlist.append(["data 5", "data 6", "data 7", "data 8"])
            if values['-N2-'] == True:
                assaycheckboxlist.append(["data 9", "data 10", "data 11", "data 12"])
            if values['-O2-'] == True:
                assaycheckboxlist.append(["data 13", "data 14"])
            if values['-AT2-'] == True:
                assaycheckboxlist.append(["data 15"])
            if values['-CA2-'] == True:
                assaycheckboxlist.append(["data 16", "data 17", "data 18", "data 19", "data 20", "data 21", "data 22", "data 23"])
            if values['-AH2-'] == True:
                assaycheckboxlist.append(["data 24"])
            if values['-KT2-'] == True:
                assaycheckboxlist.append(["data 25"])
            if values['-BO2-'] == True:
                assaycheckboxlist.append(["data 26", "data 27"])

            assaycheckboxlist = [j for i in assaycheckboxlist for j in i]
            if len(assaycheckboxlist) == 0:
                window['assaystatus'].print("Please pick at least one chemical class (Assay Tab).")            
            
            #Filters database tabs and joins them into one table based on chemical class selections
            assayfullsheet = pd.concat(oner, keys=assaycheckboxlist)
            return assayfullsheet

def zformat(chemfullsheet):
    #Swaps 1536 and 384W locations for Z and 1Z classes to ensure they load onto transfer sheet
    wellswap = {'384W ML Well': '1536W ML Well', '1536W ML Well':'384W ML Well'}
    plateswap = {'384W ML': '1536W ML', '1536W ML': '384W ML'}
    chemfullsheet.update(chemfullsheet.loc[(chemfullsheet['Class'].isin(['Z','1Z'])), ['384W ML Well', '1536W ML Well']].rename(wellswap,axis=1))
    chemfullsheet.update(chemfullsheet.loc[(chemfullsheet['Class'].isin(['Z','1Z'])), ['384W ML', '1536W ML']].rename(plateswap,axis=1))
    chemslice = chemfullsheet.loc[:, ['Class','Molecule Name','Batch Name','384W ML','384W ML Well', '1536W LL', '1536W LL Well']]
    chemslice["Transfer Volume"] = 0
    return chemslice

def aformat(assayfullsheet):
    #Logic to check whether user is running 1536 or 384W assay
    if values['-A1536-'] == True:
        alphabetcheck = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF"]
        numcheck = list(range(1,49))
        assayslice = assayfullsheet.loc[:, ['Class','Molecule Name','Batch Name','1536W LL','1536W LL Well','1536W ZI','1536W ZI Well']]
        return alphabetcheck, numcheck, assayslice
    elif values['-A384-'] == True:
        alphabetcheck = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
        numcheck = list(range(1,25))
        assayslice = assayfullsheet.loc[:, ['Class','Molecule Name','Batch Name','1536W LL','1536W LL Well','384W ZI','384W ZI Well']]
        return alphabetcheck, numcheck, assayslice
    else:
        window['assaystatus'].print("Please pick an assay plate format, 1536 or 384 well plate.")
    
def ctvol(chemslice):
    #Begin filling in values for user-entered transfer volumes
    if values['-VNN-'] != '0':
        if float(values['-VNN-']) > 2000 or float(values['-VNN-']) < 2.5:
            window['chemstatus'].print("XA transfer volume is not compatible with the Echo. Setting to default of 50nL.")
            chemslice.loc[(chemslice['Class'] == 'XA'),"Transfer Volume"] = '50'
        else:
            chemslice.loc[(chemslice['Class'] == 'XA'),"Transfer Volume"] = values['-VNN-']   
    if values['-VN-'] != '0':
        if float(values['-VN-']) > 2000 or float(values['-VN-']) < 2.5:
            window['chemstatus'].print("XB transfer volume is not compatible with the Echo. Setting to default of 50nL.")
            chemslice.loc[(chemslice['Class'] == 'XB'),"Transfer Volume"] = '50'
        else:
            chemslice.loc[(chemslice['Class'] == 'XB'),"Transfer Volume"] = values['-VN-']
    if values['-VZZ-'] != '0':
        if float(values['-VZZ-']) > 2000 or float(values['-VZZ-']) < 2.5:
            window['chemstatus'].print("Z transfer volume is not compatible with the Echo. Setting to default of 50nL.")
            chemslice.loc[(chemslice['Class'] == 'Z'),"Transfer Volume"] = '50'
        else:
            chemslice.loc[(chemslice['Class'] == 'Z'),"Transfer Volume"] = values['-VZZ-']
    if values['-VZ-'] != '0':
        if float(values['-VZ-']) > 2000 or float(values['-VZ-']) < 2.5:
            window['chemstatus'].print("1Z transfer volume is not compatible with the Echo. Setting to default of 50nL.")
            chemslice.loc[(chemslice['Class'] == '1Z'),"Transfer Volume"] = '50'
        else:
            chemslice.loc[(chemslice['Class'] == '1Z'),"Transfer Volume"] = values['-VZ-']
    if values['-VO-'] != '0':
        if float(values['-VO-']) > 2000 or float(values['-VO-']) < 2.5:
            window['chemstatus'].print("PYT transfer volume is not compatible with the Echo. Setting to default of 50nL.")
            chemslice.loc[(chemslice['Class'] == 'PYT'),"Transfer Volume"] = '50'
        else:
            chemslice.loc[(chemslice['Class'] == 'PYT'),"Transfer Volume"] = values['-VO-']
    if values['-VAT-'] != '0':
        if float(values['-VAT-']) > 2000 or float(values['-VAT-']) < 2.5:
            window['chemstatus'].print("Q transfer volume is not compatible with the Echo. Setting to default of 50nL.")
            chemslice.loc[(chemslice['Class'] == 'Q'),"Transfer Volume"] = '50'
        else:
            chemslice.loc[(chemslice['Class'] == 'Q'),"Transfer Volume"] = values['-VAT-']
    if values['-VAH-'] != '0':
        if float(values['-VAH-']) > 2000 or float(values['-VAH-']) < 2.5:
            window['chemstatus'].print("TGTT transfer volume is not compatible with the Echo. Setting to default of 50nL.")
            chemslice.loc[(chemslice['Class'] == 'TGTT'),"Transfer Volume"] = '50'
        else:
            chemslice.loc[(chemslice['Class'] == 'TGTT'),"Transfer Volume"] = values['-VAH-']
    if values['-VCA-'] != '0':
        if float(values['-VCA-']) > 2000 or float(values['-VCA-']) < 2.5:
            window['chemstatus'].print("ZO transfer volume is not compatible with the Echo. Setting to default of 50nL.")
            chemslice.loc[(chemslice['Class'] == 'ZO'),"Transfer Volume"] = '50'
        else:
            chemslice.loc[(chemslice['Class'] == 'ZO'),"Transfer Volume"] = values['-VCA-']
    if values['-VKT-'] != '0':
        if float(values['-VKT-']) > 2000 or float(values['-VKT-']) < 2.5:
            window['chemstatus'].print("BR transfer volume is not compatible with the Echo. Setting to default of 50nL.")
            chemslice.loc[(chemslice['Class'] == 'BR'),"Transfer Volume"] = '50'
        else:
            chemslice.loc[(chemslice['Class'] == 'BR'),"Transfer Volume"] = values['-VKT-']
    if values['-VBO-'] !='0':
        if float(values['-VBO-']) > 2000 or float(values['-VBO-']) < 2.5:
            window['chemstatus'].print("W transfer volume is not compatible with the Echo. Setting to default of 50nL.")
            chemslice.loc[(chemslice['Class'] == 'W'),"Transfer Volume"] = '50'
        else:
            chemslice.loc[(chemslice['Class'] == 'W'),"Transfer Volume"] = values['-VBO-'] 
    return chemslice

def atvol(assayslice, assayxvol):
    if assayxvol > 2000 or assayxvol < 2.5:
        window['assaystatus'].print("Incompatible volume entered, setting to default of 50nL.")
        assayslice['Transfer Volume'] = '50'
    else:
        assayslice['Transfer Volume'] = assayxvol
    xbw = int(values['-xbw-'])+1
    bbw = int(values['-bbw-'])+1
    hiw = int(values['-hiw-'])+1
    low = int(values['-low-'])+1
    
    if values['-xbv-'] != 'N/A':
        xbv = int(values['-xbv-'])
        assayslice.loc[(assayslice['Molecule Name'] == 'XB-01'), 'Transfer Volume'] = xbv
    if values['-bbv-'] != 'N/A':
        bbv = int(values['-bbv-'])
        assayslice.loc[(assayslice['Molecule Name'] == 'BB8900'), 'Transfer Volume'] = bbv
    if values['-hiv-'] != 'N/A':
        hiv = int(values['-hiv-'])
        assayslice.loc[(assayslice['Molecule Name'] == 'EC98'), 'Transfer Volume'] = hiv
    if values['-lov-'] != 'N/A':
        lov = int(values['-lov-'])
        assayslice.loc[(assayslice['Molecule Name'] == 'EC57'), 'Transfer Volume'] = lov
                
    if xbw != 32:
        xbwrange = list(range(xbw,33))
        assayslice.loc[(assayslice['Batch Name'].isin(xbwrange)) & (assayslice['Molecule Name'] == 'XB-01'), 'Transfer Volume'] = 0
    if bbw != 32:
        bbwrange = list(range(bbw,33))
        assayslice.loc[(assayslice['Batch Name'].isin(bbwrange)) & (assayslice['Molecule Name'] == 'BB8900'), 'Transfer Volume'] = 0
    if hiw != 32:
        hiwrange = list(range(hiw,33))
        assayslice.loc[(assayslice['Batch Name'].isin(hiwrange)) & (assayslice['Molecule Name'] == 'EC98'), 'Transfer Volume'] = 0
    if low != 32:
        lowrange = list(range(low,33))
        assayslice.loc[(assayslice['Batch Name'].isin(lowrange)) & (assayslice['Molecule Name'] == 'EC57'), 'Transfer Volume'] = 0
    return assayslice

def actrl(selxb, selbb, selhi, sello, alphabetcheck, numcheck, assayslice):
    #For xb
    if selxb != 'None':
        xbcheck = possplitter(selxb)
        xblet = str(xbcheck[0])
        xbnum = int(xbcheck[1])
        if xblet in alphabetcheck and xbnum in numcheck:
            assayslice.loc[(assayslice['Molecule Name'] == "XB-01"), "1536W LL Well"] = selxb
        else:
            window['assaystatus'].print("Please enter a valid plate well position for xb or type 'None'.")
    else:
        window['assaystatus'].print("Entire column for xb will be used.")
                
    #For BB8900 (positive control)    
    if selbb != 'None':
        bbcheck = possplitter(selbb)
        bblet = str(bbcheck[0])
        bbnum = int(bbcheck[1])
        if bblet in alphabetcheck and bbnum in numcheck:
            assayslice.loc[(assayslice['Molecule Name'] == "BB8900"), "1536W LL Well"] = selbb
        else:
            window['assaystatus'].print("Please enter a valid plate well position for bb8900 or type 'None'.")
    else:
        window['assaystatus'].print("Entire column for bb/ctrl will be used.")
                
    #For EC98 control (assay)
    if selhi != 'None':
        hicheck = possplitter(selhi)
        hilet = str(hicheck[0])
        hinum = int(hicheck[1])
        if hilet in alphabetcheck and hinum in numcheck:
            assayslice.loc[(assayslice['Molecule Name'] == "EC98"), "1536W LL Well"] = selhi
        else:
            window['assaystatus'].print("Please enter a valid plate well position for hi or type 'None'.")
    else:
        window['assaystatus'].print("Entire column for high control will be used.")  
    
    #For EC57 control (assay)    
    if sello != 'None':
        locheck = possplitter(sello)
        lolet = str(locheck[0])
        lonum = int(locheck[1])
        if lolet in alphabetcheck and lonum in numcheck:
            assayslice.loc[(assayslice['Molecule Name'] == "EC57"), "1536W LL Well"] = sello
        else:
            window['assaystatus'].print("Please enter a valid plate well position for lo or type 'None'.")
    else:
        window['assaystatus'].print("Entire column for low control will be used.")  
    
    return assayslice

def csheetclean(chemslice):
    #Clean up the dataframe and finalize it
    if values['-removeempties-'] == True:
        chemslice = chemslice[(chemslice['Class'] != 'empty') & (~chemslice['Molecule Name'].isin(['XB-01', 'BB8900', 'EC57', 'EC98']))]
        chemfinal = chemslice.loc[:, ['Class','Molecule Name','Batch Name','384W ML','384W ML Well', '1536W LL', '1536W LL Well', 'Transfer Volume']]
    else:
        chemslice.loc[chemslice['Class'] == 'empty', 'Transfer Volume'] = 0
        chemfinal = chemslice.loc[:, ['Class','Molecule Name','Batch Name','384W ML','384W ML Well', '1536W LL', '1536W LL Well', 'Transfer Volume']]
        
    #Create the transfer sheet
    outfile = os.path.dirname(values['-mf-']) + '/' + str(datetime.date.today().isoformat()) + '_' + '%s chem Transfer Sheet.csv' % chemmetadata
    chemfinal.columns = ['Class', 'Molecule Name', 'Batch Name', 'Source Plate Barcode', 'Source Well', 'Destination Plate Barcode', 'Destination Well', 'Transfer Volume']
    return chemfinal, outfile

def asheetclean(assayslice):
    #Clean up the transfer sheet to remove empty plate positions, and export the .csv file
    if values['-aremoveempties-'] == True:
        assayfinal = assayslice[(assayslice['Class'] != 'empty') & (assayslice['Transfer Volume'] != 0)]
    else:
        assayslice.loc[assayslice['Class'] == 'empty', 'Transfer Volume'] = 0
        assayfinal = assayslice
        
    if values['-A1536-'] == True:
        assayfiletype = '1536W'
    elif values['-A384-'] == True:
        assayfiletype = '384W'
    else:
        assayfiletype = 'Unknown Plate Type'
    outfile = os.path.dirname(values['-file-']) + '/' + str(datetime.date.today().isoformat()) + '_' + '%s' % ametadata + '_' + str(assayfiletype) + '_' + 'Assay Transfer Sheet.csv'
    assayfinal.columns = ['Class', 'Molecule Name', 'Batch Name', 'Source Plate Barcode', 'Source Well', 'Destination Plate Barcode', 'Destination Well', 'Transfer Volume']
    return assayfinal, outfile

def swindow():
    sg.theme('DarkTeal12')

    #Define checkboxes for chemical classes in Assay Transfer Tab
    assaycheckboxes = [[sg.Checkbox(':XA', default=False, key="-NN2-"), sg.Checkbox(':XB', default=False, key="-N2-"),
                        sg.Checkbox(':1Z', default=False, key="-Z2-"), sg.Checkbox(':Z', default=False, key="-ZZ2-"),
                        sg.Checkbox(':PYT', default=False, key="-O2-"), sg.Checkbox(':Q', default=False, key="-AT2-"),
                        sg.Checkbox(':ZO', default=False, key="-CA2-"), sg.Checkbox(':TGTT', default=False, key="-AH2-"),
                        sg.Checkbox(':BR', default=False, key="-KT2-"), sg.Checkbox(':W', default=False, key="-BO2-")]
                       ]

    #Define inputs for chemical class volumes in Assay Transfer Tab, will implement if needed.... currently it is not being used
    #assayvolumes = [[sg.Frame('XA:', [[sg.Input(0, key="-VNN2-",size=(5,1))]]), sg.Frame('XB:', [[sg.Input(0, key="-VN2-",size=(5,1))]])],
    #                [sg.Frame('Z:', [[sg.Input(0, key="-VZZ2-",size=(5,1))]]), sg.Frame('1Z:', [[sg.Input(0, key="-VZ2-",size=(5,1))]])],
    #                [sg.Frame('PYT:', [[sg.Input(0, key="-VO2-",size=(5,1))]]), sg.Frame('Q:', [[sg.Input(0, key="-VAT2-",size=(5,1))]])],
    #                [sg.Frame('ZO:', [[sg.Input(0, key="-VCA2-",size=(5,1))]]), sg.Frame('TGTT:', [[sg.Input(0, key="-VAH2-",size=(5,1))]])],
    #                [sg.Frame('BR:', [[sg.Input(0, key="-VKT2-",size=(5,1))]]), sg.Frame('W:', [[sg.Input(0, key="-VBO2-",size=(5,1))]])]
    #                ]

    #Define inputs for chemical class volumes in Chemical Transfer Tab
    chemvolumes = [[sg.Frame('XA:', [[sg.Input('0', key="-VNN-",size=(5,1))]]), sg.Frame('XB:', [[sg.Input('0', key="-VN-",size=(5,1))]])],
                   [sg.Frame('Z:', [[sg.Input('0', key="-VZZ-",size=(5,1))]]), sg.Frame('1Z:', [[sg.Input('0', key="-VZ-",size=(5,1))]])],
                   [sg.Frame('PYT:', [[sg.Input('0', key="-VO-",size=(5,1))]]), sg.Frame('Q:', [[sg.Input('0', key="-VAT-",size=(5,1))]])],
                   [sg.Frame('ZO:', [[sg.Input('0', key="-VCA-",size=(5,1))]]), sg.Frame('TGTT:', [[sg.Input('0', key="-VAH-",size=(5,1))]])],
                   [sg.Frame('BR:', [[sg.Input('0', key="-VKT-",size=(5,1))]]), sg.Frame('W:', [[sg.Input('0', key="-VBO-",size=(5,1))]])]
                   ]

    #Define layout of Chemical Transfer Tab
    chemxfer = [
        [sg.Text("Chemical Transfer Sheet Generator",font='Any 18')],
        [sg.Frame('Browse to database file:', [[sg.Input(key="-mf-"), sg.FileBrowse(target="-mf-")]])],
        [sg.Frame('Enter volumes for desired chemical classes in nanoliters:',chemvolumes)],
        [sg.Text("Enter metadata, e.g. xb ID, LL type, assay info, target, assay concentration: ")],
        [sg.Input(key="-chemmetadata-")],
        [sg.Checkbox('Remove empty rows?', default=False, key="-removeempties-")],
        [sg.Button("Generate Chemical Transfer Sheet"), sg.Button("Exit (Chemical Tab)")]
        ]

    #Define layout of Assay Transfer Tab
    assayxfer = [
        [sg.Text("Assay Transfer Sheet Generator",font='Any 18')],          
        [sg.Frame('Browse to database file:', [[sg.Input(key="-file-"), sg.FileBrowse(target="-file-")]])],
        [sg.Frame('Select among the following classes:', assaycheckboxes)],
        [sg.Frame("Enter transfer volume in nanoliters:",[[sg.Input(key='-avol-')]])],
        [sg.Frame("Enter control information (anything left as default assumes all rows in control columns will be used):",
              [
                  [sg.Text("Column 45 = xb, Column 46 = bb, Column 47 = EC98 Control, Column 48 = EC57 Control")],
                  [sg.Frame('xb:',[
                      [sg.Text("Position:"),sg.Input('None',size=(5,1),key="-xbp-")],
                      [sg.Text("No. of Wells:"),sg.Input(32,size=(5,1),key="-xbw-")],
                      [sg.Text("Vol. to Xfer:"),sg.Input('N/A',size=(5,1),key="-xbv-")]]),
                      sg.Frame('BB8900:',[
                          [sg.Text("Position:"),sg.Input('None',size=(5,1),key="-bbp-")],
                          [sg.Text("No. of Wells:"),sg.Input(32,size=(5,1),key="-bbw-")],
                          [sg.Text("Vol. to Xfer:"),sg.Input('N/A',size=(5,1),key="-bbv-")]]),
                      sg.Frame('EC98 Ctrl:',[
                          [sg.Text("Position:"),sg.Input('None',size=(5,1),key="-hip-")],
                          [sg.Text("No. of Wells:"),sg.Input(32,size=(5,1),key="-hiw-")],
                          [sg.Text("Vol. to Xfer:"),sg.Input('N/A',size=(5,1),key="-hiv-")]]),
                      sg.Frame('EC57 Ctrl:',[
                          [sg.Text("Position:"),sg.Input('None',size=(5,1),key="-lop-")],
                          [sg.Text("No. of Wells:"),sg.Input(32,size=(5,1),key="-low-")],
                          [sg.Text("Vol. to Xfer:"),sg.Input('N/A',size=(5,1),key="-lov-")]])
                   ]
               ]
              )
         ],
        [sg.Frame("Enter metadata, e.g. xb ID, LL type, assay info, target, assay concentration:",[[sg.Input(key="-assaymetadata-")]])],
        [sg.Frame("Please select desired assay plate type:", [[sg.Radio('1536W Assay', "AssayType", default=False, key="-A1536-"), sg.Radio('384W Assay', "AssayType", default=False, key="-A384-")]])],
        [sg.Checkbox('Remove empty rows?', default=False, key="-aremoveempties-")],
        [sg.Button("Generate Assay Transfer Sheet"), sg.Button("Exit (Assay Tab)")]
    ]

    #Define output textbox for Chemical Transfer Tab    
    chemstatus = [
        [sg.Text('Status:', size=[20,1])],
        [sg.Multiline(key='chemstatus',autoscroll=True,size=(30,20))],
    ]

    #Define output textbox for Assay Transfer Tab
    assaystatus = [
        [sg.Text('Status:', size=[20,1])],
        [sg.Multiline(key='assaystatus',autoscroll=True,size=(30,20))],
    ]

    #chem Transfer tab Window Layout
    chemlayout = [
        [
        sg.Column(chemxfer),
        sg.Column(chemstatus)
        ]
    ]

    #Assay Transfer tab Window Layout
    assaylayout = [
        [
        sg.Column(assayxfer),
        sg.Column(assaystatus)
        ]
    ]

    #Define layout with tabs
    tabgrp = [[sg.TabGroup([[
        sg.Tab('Chemical Transfer', chemlayout, title_color='Black', border_width=10),
        sg.Tab('Assay Transfer', assaylayout, title_color='Black', border_width=10)
        ]],
        tab_location='centertop',title_color='Green',
        tab_background_color='Gray',
        selected_title_color='Black',
        selected_background_color='Green', border_width=5)
    ]]
    return tabgrp

tabgrp = swindow()

#Define window
window = sg.Window('Chemical Database Tool', tabgrp, no_titlebar=False, alpha_channel=.9, grab_anywhere=True)

#Create event loop to enable user inputs
while True:
    event, values = window.read()
    if event in (sg.WINDOW_CLOSED, "Exit (Chemical Tab)") or event in (sg.WINDOW_CLOSED, "Exit (Assay Tab)"):
        break
    
   
    #Code for Chemical Transfer Tab
    elif event == "Generate Chemical Transfer Sheet":
        try:
            filename = values['-mf-']
            chemmetadata = values['-chemmetadata-']
            oner = process_xl(27, filename)
            chemfullsheet = ctablist()
            chemslice = zformat(chemfullsheet)           
            chemslice = ctvol(chemslice)
            chemfinal, outfile = csheetclean(chemslice)      
            chemfinal.to_csv(outfile, index=False)
            window['chemstatus'].print("Chemical Transfer Sheet Generation Successful! \n Export to outfile %s completed." % outfile)
        except:
            window['chemstatus'].print("Chemical Transfer Sheet did not generate.")

    #Code for Assay Transfer Tab
    elif event == "Generate Assay Transfer Sheet":
        try:
            assayfilename = values['-file-']
            assayxvol = values['-avol-']
            selxb = values['-xbp-']
            selbb = values['-bbp-']
            selhi = values['-hip-']
            sello = values['-lop-']
            ametadata = values['-assaymetadata-']
            oner = process_xl(27, assayfilename)
            assayfullsheet = atablist()
                 
            
            #Next sets of code decides whether to keep xb, bb8900, ec98, and ec57 wells fixed based on user selection, or to use entire column
            alphabetcheck, numcheck, assayslice = aformat(assayfullsheet)

            #Add transfer volume column for samples, xb, bb8900, ec98, ec57 based on user-input
            assayslice = atvol(assayslice, assayxvol)
            
            #Logic to make sure that users enter valid plate positions for xb, bb8900, ec98, and ec57 standards/controls.
            #Will notify user if whole "default" column will be used:
                #xb = Column 45
                #BB8900 = Column 46
                #EC98 Assay control = Column 47
                #EC57 Assay control = Column 48
            assayslice = actrl(selxb, selbb, selhi, sello, alphabetcheck, numcheck, assayslice)
               

            assayfinal, outfile = asheetclean(assayslice)
            assayfinal.to_csv(outfile, index=False)
            window['assaystatus'].print("Assay Transfer Sheet Generation Successful! \n Export to outfile %s completed." % outfile)
        except:
            window['assaystatus'].print("Assay Transfer Sheet did not generate.")
    elif event == sg.WIN_CLOSED:
        break
    elif event == "Exit":
        break
window.close()

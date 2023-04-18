# -*- coding: utf-8 -*-
"""
"""




import win32com.client
import pandas as pd
from datetime import datetime
import subprocess
import os

#os.chdir("")

SapGuiAuto = win32com.client.GetObject('SAPGUI')
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)
now = datetime.now()
dt_string = now.strftime("%Y%m%d %H%M")
filename = "FlightSalesReport " + dt_string + ".XLSX"
folderdir = "C:\\Users\\Weiche.Fu\\Downloads\\"
folderdir+filename
#startdate = input("Please input your start date in this format (MM/DD/YYYY)")
#enddate= input("Please input your ending date in this format(MM/DD/YYYY)")



###put the SAPscript into here/Remember to change the code into python

session.findById("wnd[0]").maximize()
session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "****"
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "*******"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "********"


session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "en"
session.findById("wnd[0]/usr/txtRSYST-MANDT").setFocus()
session.findById("wnd[0]/usr/txtRSYST-LANGU").caretPosition = 2
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[11]").press()
session.findById("wnd[0]/tbar[0]/okcd").text = "*****" 
session.findById("wnd[0]").sendVKey(0) 

session.findById("wnd[0]/usr/tabsTABSTR/tabpPRD/ssubSUBSCREENAREA:SAPMZVP13:0130/ctxtWA_SEL_0130-MATL").text = "*"
session.findById("wnd[0]/usr/tabsTABSTR/tabpPRD/ssubSUBSCREENAREA:SAPMZVP13:0130/chkWA_0130D-PREIS").setFocus()
session.findById("wnd[0]/usr/tabsTABSTR/tabpPRD/ssubSUBSCREENAREA:SAPMZVP13:0130/chkWA_0130D-PREIS").selected = True
session.findById("wnd[0]/usr/tabsTABSTR/tabpPRD/ssubSUBSCREENAREA:SAPMZVP13:0130/chkWA_0130D-MIN").setFocus()
session.findById("wnd[0]/usr/tabsTABSTR/tabpPRD/ssubSUBSCREENAREA:SAPMZVP13:0130/chkWA_0130D-MIN").selected = True
session.findById("wnd[0]/usr/tabsTABSTR/tabpPRD/ssubSUBSCREENAREA:SAPMZVP13:0130/ctxtWA_SEL_0130-VTWL").text = "Q1"
# session.findById("wnd[0]/usr/tabsTABSTR/tabpPRD/ssubSUBSCREENAREA:SAPMZVP13:0130/ctxtWA_SEL_0130-VTWH").text = "Q7"
session.findById("wnd[0]/usr/tabsTABSTR/tabpPRD/ssubSUBSCREENAREA:SAPMZVP13:0130/ctxtWA_SEL_0130-VTWL").setFocus()
session.findById("wnd[0]/usr/tabsTABSTR/tabpPRD/ssubSUBSCREENAREA:SAPMZVP13:0130/ctxtWA_SEL_0130-VTWL").caretPosition = 1
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/tabsTABSTR/tabpPRD/ssubSUBSCREENAREA:SAPMZVP13:0130/cntlCC0130/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
session.findById("wnd[0]/usr/tabsTABSTR/tabpPRD/ssubSUBSCREENAREA:SAPMZVP13:0130/cntlCC0130/shellcont/shell").selectContextMenuItem("&XXL")
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\Weiche.Fu\\Downloads\\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 33
session.findById("wnd[1]/tbar[0]/btn[0]").press()



connection= None
application = None
SapGuiAuto = None 

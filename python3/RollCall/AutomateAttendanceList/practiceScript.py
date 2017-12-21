import os
import win32com.client

if os.path.exists("C:\\Users\powel\Desktop\Roll Call\RollCall-FirstStop.xlsm") & os.path.exists("C:\\Users\powel\Desktop\Roll Call\Fall 2017 Criner Oscar CS 124 01.xlsm"):
    xl=win32com.client.Dispatch("Excel.Application")
    xl.Workbooks.Open("C:\\Users\powel\Desktop\Roll Call\RollCall-FirstStop.xlsm")
    xl.Workbooks.Open("C:\\Users\powel\Desktop\Roll Call\Fall 2017 Criner Oscar CS 124 01.xlsm")
    xl.Application.Run("RollCall-FirstStop.xlsm!RollCall.begin")
    ##    xl.Application.Save() # if you want to save then uncomment this line and change delete the ", ReadOnly=1" part from the open function.
    xl.Application.Quit() # Comment this out if your excel script closes
    del xl

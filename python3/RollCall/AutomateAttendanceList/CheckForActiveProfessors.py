import os
import win32com.client

macroBookName = "RollCall-FirstStop.xlsm!"
getSMSMacroName = "RollCall.begin"
attendanceCompareMacroName = "UpdateRoll.UpdateRoll"
WbName = "ClassRollSheet"
ModuleName = "!RollCall"
MacroName = "begin"
fullMacroPath = "{}{}"


def RunExcelMacro(macroName):
    xl = win32com.client.Dispatch('Excel.Application')
    xl.Workbooks.Open("C:\\Users\powel\Desktop\Roll Call\Fall 2017 Criner Oscar CS 124 01.xlsm")
    xl.Workbooks.Open("C:\\Users\powel\Desktop\Roll Call\RollCall-FirstStop.xlsm")
    xl.Visible = 1
    whatWasReturned = xl.Application.Run("'RollCall-FirstStop.xlsm'!begin()")
    print(whatWasReturned)
    # xl.Application.Run("UpdateRoll.UpdateRoll")
    xl.Application.Quit()


RunExcelMacro(ModuleName + MacroName)

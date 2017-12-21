import win32com.client as win32

macroBookName = "RollCall-FirstStop.xlsm!"
getSMSMacroName = "RollCall.begin"
attendanceCompareMacroName = "UpdateRoll.UpdateRoll"


def openWorkbook(xlapp, xlfile):
    try:
        xlwb = xlapp.Workbooks(xlfile)
    except Exception as e:
        try:
            xlwb = xlapp.Workbooks.Open(xlfile)
        except Exception as e:
            print(e)
            xlwb = None
    return (xlwb)


try:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    # professorsListwb = openWorkbook(excel,
    #                   'C:\\Users\powel\Python3\RollCall\AutomateAttendanceList\ProfessorsList.xlsx')
    # professorsListws = professorsListwb.Worksheets('Professors Directory List')

    rollCallwb = openWorkbook(excel,
                              "C:\\Users\powel\Desktop\Roll Call\RollCall-FirstStop.xlsm")
    rollCallws = rollCallwb.Worksheets('ClassRollSheet')
    excel.Application.Run(macroBookName+getSMSMacroName)
    excel.Visible = True

except Exception as e:
    print(e)

finally:
    # RELEASES RESOURCES
    ws = None
    wb = None
    excel = None

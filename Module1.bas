Attribute VB_Name = "Module1"
Sub AutomatedAttendenceUpdate()
    Dim startTime, endTime, elapsedTime As String
    Dim this As Workbook
    Set this = Workbooks("ProfessorsList.xlsm")
    Dim thisSheet As Worksheet
    Set thisSheet = this.Worksheets(1)
    Lastrow = thisSheet.Cells(thisSheet.Rows.Count, "A").End(xlUp).Row
    startTime = Now()
    For i = 1 To Lastrow
        If (IsEmpty(Cells(i, 3)) & i < Lastrow) Then
            i = i + 1
        Else
            If InStr(thisSheet.Cells(i, 3).Value, ".xlsm") > 0 Then
                If InStr(thisSheet.Cells(i, 3), "~$") > 0 Then
                Else
                    Application.DisplayAlerts = False
                    Set wb = Workbooks.Open(thisSheet.Cells(i, 3).Value)
                    wb.Close
                End If
            Else
                i = i + 1
                'MsgBox ("Not Valid Path")
            End If
        End If
    Next i
    endTime = Now()
    elapsedTime = endTime - startTime
    MsgBox ("EOF: " & Format(elapsedTime, "mm:ss") & " and " & i & " Records")
              
End Sub

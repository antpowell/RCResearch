VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Grade()
    Dim numberMissed As Integer, studentCol As Integer, studentRow As Integer, totalQuestions As Integer
    
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    studentCol = 5
    studentRow = 3
    totalQuestions = 37
    
    For i = studentRow To lastRow
        ActiveSheet.Cells(i, studentCol) = ((totalQuestions - ActiveSheet.Cells(i, studentCol).Value) / totalQuestions) * 100
    Next i
    
    
End Sub

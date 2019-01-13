Attribute VB_Name = "Module1"
Option Explicit

Sub Copy_Paste_Macros()
Attribute Copy_Paste_Macros.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' Copy_Paste_Macros Macro
'
' Keyboard Shortcut: Ctrl+q
'
    Dim lastRowFrom As Long
    Dim lastColFrom As Long
    Dim lastRowTo As Long
    Dim lastColTo As Long
    Dim fromBook As String
    Dim toBook As String
    Dim book As Workbook
    Dim q As Integer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    fromBook = ActiveWorkbook.Name
    
    For Each book In Workbooks
        If ActiveWorkbook.Name <> book.Name Then
            toBook = book.Name
        End If
    Next
    
    For q = 1 To Workbooks(fromBook).Sheets.Count
        ActiveWorkbook.Sheets(q).Activate
        
        lastRowFrom = Cells(Rows.Count, 1).End(xlUp).Row
        lastColFrom = Cells(1, Columns.Count).End(xlToLeft).Column
        
        Range(Cells(1, 1), Cells(lastRowFrom, lastColFrom)).Select
        Selection.Copy
        
        Windows(toBook).Activate
        
        ActiveWorkbook.Sheets(q).Activate
        
        lastRowTo = Cells(Rows.Count, 1).End(xlUp).Row
        lastColTo = Cells(1, Columns.Count).End(xlToLeft).Column
        
        ActiveSheet.Cells(lastRowTo + 1, 1).Select
        
        ActiveSheet.Paste
        
        Windows(fromBook).Activate
    Next q
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

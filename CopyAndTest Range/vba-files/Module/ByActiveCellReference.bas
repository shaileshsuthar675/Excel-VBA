Attribute VB_Name = "ByActiveCellReference"
Option Explicit


Sub CopyAndPasteAsValueOnWorksheetByActiveCell()
    Dim rng As String, WorksheetExist As Boolean, Sheet As Worksheet, tWsName As String, ws As Worksheet
    Dim ActiveCellAddress As String, xl_down_cell As String, xl_Down_Right_cell As String
    
    ActiveCellAddress = ActiveCell.Address
    xl_down_cell = Range(ActiveCellAddress).End(xlDown).Address
    xl_Down_Right_cell = Range(xl_down_cell).End(xlToRight).Address
    rng = ActiveCellAddress & ":" & xl_Down_Right_cell
    tWsName = ActiveSheet.Name
    WorksheetExist = False
    
    For Each Sheet In Worksheets
        If (Sheet.Name = "Output") Then
            WorksheetExist = True
            Worksheets("Output").Activate
            Cells.Clear
        End If
        Next
        If (WorksheetExist = False) Then
            Set ws = Worksheets.Add
            ws.Name = "Output"
        End If
        Worksheets(tWsName).Activate
        Range(rng).Copy
        Worksheets("Output").Range("A1").PasteSpecial xlPasteValues
        Worksheets("Output").Range("A1").PasteSpecial xlPasteFormats
        Worksheets("Output").Activate
        Application.CutCopyMode = False
End Sub


Sub CopyAndPasteAsValueOnWorkbookByActiveCell()
    Application.ScreenUpdating = False
    Dim ThisWB As Workbook, rng As String, tWsName As String, Newbook As Workbook
    Dim fname As String
    Set ThisWB = ActiveWorkbook
    Dim ActiveCellAddress As String, xl_down_cell As String, xl_Down_Right_cell As String
    
    ActiveCellAddress = ActiveCell.Address
    xl_down_cell = Range(ActiveCellAddress).End(xlDown).Address
    xl_Down_Right_cell = Range(xl_down_cell).End(xlToRight).Address
    rng = ActiveCellAddress & ":" & xl_Down_Right_cell
    tWsName = ThisWB.ActiveSheet.Name
    Range(rng).Copy
    Set Newbook = Workbooks.Add

    Newbook.Worksheets("Sheet1").Range("A1").PasteSpecial (xlPasteValues)
    Newbook.Worksheets("Sheet1").Range("A1").PasteSpecial (xlPasteFormats)
    Newbook.Worksheets("Sheet1").Range("A:J").Columns.AutoFit

    fname = ThisWB.Path & "\" & "Output.xlsx"
    If Dir(fname) <> "" Then
        If MsgBox("Output already exists, are you sure you want To overwrite?", vbOKCancel) = vbCancel Then Newbook.Close False: Application.CutCopyMode = False: Exit Sub
        End If

        Application.DisplayAlerts = False
        Newbook.SaveAs Filename:=fname
        Application.DisplayAlerts = True
        ThisWB.Activate
        ActiveWorkbook.Worksheets(tWsName).Range(rng).Select
        Newbook.Activate
        ActiveWorkbook.ActiveSheet.Range("A1").Select
        Application.CutCopyMode = False
        Application.ScreenUpdating = True

End Sub




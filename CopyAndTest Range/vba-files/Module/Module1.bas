Option Explicit

Private Sub ComboBox1_Change()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub PasteValues_Click()
    Dim rng As String, WorksheetExist As Boolean, Sheet As Worksheet, tWsName As String, ws As Worksheet
    Dim ActiveCellAddress As String, fname As String, book_sheet_name As String
    book_sheet_name = LCase(WSname)
    If (UserForm1.ComboBox1.Text = "Worksheet") Then
        ActiveCellAddress = ActiveCell.Address
        rng = ActiveCellAddress & ":" & xlCell
        tWsName = ActiveSheet.Name
        WorksheetExist = False
        For Each Sheet In Worksheets
            If (Sheet.Name = book_sheet_name) Then
                WorksheetExist = True
                Worksheets(book_sheet_name).Activate
                Cells.Clear
            End If
            Next
            If (WorksheetExist = False) Then
                Set ws = Worksheets.Add
                ws.Name = book_sheet_name
            End If
            Worksheets(tWsName).Activate
            Range(rng).Copy
            Worksheets(book_sheet_name).Range("A1").PasteSpecial xlPasteValues
            Worksheets(book_sheet_name).Range("A1").PasteSpecial xlPasteFormats
            Worksheets(book_sheet_name).Activate
            Application.CutCopyMode = False
        Else
            Application.ScreenUpdating = False
            Dim ThisWB As Workbook, Newbook As Workbook
            Set ThisWB = ActiveWorkbook

            ActiveCellAddress = ActiveCell.Address
            rng = ActiveCellAddress & ":" & xlCell
            tWsName = ThisWB.ActiveSheet.Name
            Range(rng).Copy
            Set Newbook = Workbooks.Add

            Newbook.Worksheets("Sheet1").Range("A1").PasteSpecial (xlPasteValues)
            Newbook.Worksheets("Sheet1").Range("A1").PasteSpecial (xlPasteFormats)
            Newbook.Worksheets("Sheet1").Range("A:J").Columns.AutoFit

            fname = ThisWB.Path & "\" & book_sheet_name & ".xlsx"
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
            End If
End Sub

Private Sub Quit_Click()
    Unload UserForm1
End Sub


Attribute VB_Name = "WBS"
Sub WBS()

    Dim wsData As Worksheet
    Dim lr As Integer ' Declare a variable to hold the last row number

    ' Ensure the active workbook is selected
    ActiveWorkbook.Activate

    ' Make the "WBS working" sheet visible and activate the "WBS Raw" sheet
    Worksheets("WBS working").Visible = xlSheetVisible
    Worksheets("WBS Raw").Activate

    ' Find the last row with data in the "WBS Raw" sheet
    lr = Cells.Find("*", Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious, False).Row

    ' Activate the "WBS working" sheet and autofill formulas for the data range
    Worksheets("WBS working").Activate
    Range("A3:AK3").Select
    Selection.AutoFill Destination:=Range("A3:AK" & lr), Type:=xlFillDefault

    ' Apply a filter to show only rows with REL or TECO in column I
    Worksheets("WBS working").Range("$A$2:$AK$" & lr).AutoFilter Field:=9, Criteria1:="=REL", _
    Operator:=xlOr, Criteria2:="=TECO"

    ' Activate the "WBS Data" sheet and clear its contents
    Worksheets("WBS Data").Activate
    Range("A3:AK3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

    ' Copy visible filtered data from "WBS Working" and paste it as values in "WBS Data"
    Worksheets("WBS Working").Activate
    ActiveSheet.Range("$A$2:$AK$" & lr).SpecialCells(xlCellTypeVisible).Copy

    Worksheets("WBS Data").Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    ' Remove the copy mode
    Application.CutCopyMode = False

    ' Remove the filter from "WBS working"
    Worksheets("WBS working").AutoFilterMode = False

    ' Activate the "WBS Data" sheet
    Worksheets("WBS Data").Activate

    ' Declare variables to hold the last row and use for looping
    Dim lastRow As Long
    Dim i As Long

    ' Set wsData as the "WBS Data" worksheet
    Set wsData = Worksheets("WBS Data")

    ' Find the last row with data in the "WBS Data" sheet
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    ' Loop through each row in the "WBS Data" sheet, starting from row 2
    For i = 2 To lastRow
        ' If column I has "TECO", set the values in columns Y, Z, AA, and AB to 0
        If wsData.Cells(i, "I").Value = "TECO" Then
            wsData.Cells(i, "Y").Value = 0
            wsData.Cells(i, "Z").Value = 0
            wsData.Cells(i, "AA").Value = 0
            wsData.Cells(i, "AB").Value = 0
        End If
    Next i

    ' Refresh all data connections in the workbook
    ActiveWorkbook.RefreshAll

    ' Display a message box when the process is complete
    MsgBox "Done"

End Sub



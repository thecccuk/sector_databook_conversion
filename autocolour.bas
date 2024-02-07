Sub AutoColour()
    Dim cell As Range
    Dim rng As Range
    Dim usedRng As Range
    Dim linkedDataColor As Long
    Dim linkedWorkbookColor As Long ' Color for links to another workbook
    Dim formulaColor As Long
    Dim valueColor As Long
    
    ' Disable Excel updates to speed up the macro
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    ' Check if the selection is a range
    If Not TypeName(Selection) = "Range" Then
        MsgBox "Please select a range of cells first.", vbExclamation
        GoTo CleanUp
    End If
    
    ' Use the intersection of the currently selected range and the used range of the sheet
    Set usedRng = Selection.Worksheet.UsedRange
    Set rng = Intersect(Selection, usedRng)
    
    ' Exit if there's no intersection with the used range
    If rng Is Nothing Then
        MsgBox "The selection does not intersect with the used range of the sheet.", vbInformation
        GoTo CleanUp
    End If
    
    ' Define colors
    linkedDataColor = RGB(255, 204, 102) ' yellow for linked data within the same workbook
    linkedWorkbookColor = RGB(255, 199, 206) ' red for links to another workbook
    formulaColor = RGB(204, 236, 255) ' blue for formulas
    numericColor = RGB(153, 255, 204) ' green for values
    
    ' Loop through each cell in the intersected range
    For Each cell In rng.Cells
        ' Check if cell contains a formula
        If cell.HasFormula Then
            ' Check for links to another workbook
            If InStr(1, cell.Formula, "[") > 0 And InStr(1, cell.Formula, "]") > 0 Then
                cell.Interior.Color = linkedWorkbookColor ' Color for links to another workbook
            ' Check for links within the same workbook
            ElseIf InStr(1, cell.Formula, "!") > 0 Then
                cell.Interior.Color = linkedDataColor ' Color for linked data within the same workbook
            Else
                cell.Interior.Color = formulaColor ' Color for formulas
            End If
        ' Skip coloring for hardcoded string values
        ElseIf Not IsEmpty(cell.Value) And Not cell.HasFormula Then
            If IsNumeric(cell.Value) Then
                cell.Interior.Color = numericColor
            Else
                cell.Interior.ColorIndex = xlNone
            End If
        ' Empty cells are changed to be uncolored
        Else
            cell.Interior.ColorIndex = xlNone
        End If
    Next cell

CleanUp:
    ' Re-enable Excel updates
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub

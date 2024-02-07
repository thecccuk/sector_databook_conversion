' Macro Name: AutoColour
' Purpose: This set of macros automatically colors cells in a selected Excel range based on specific criteria.
'          Colors are coloured differently based on whether they contain:
'          - Formulas linking to other workbooks (colored red)
'          - Formulas linking within the same workbook (colored yellow)
'          - Formulas not linking to other cells or workbooks (colored blue)
'          - Numeric values (colored green)
'          - Leaves cells with string values and unfulfilled conditions uncolored (white)
'
' How to Use:
' 1. Open Excel and navigate to the workbook where you want to add and use the macro.
' 2. Press Alt + F11 to open the Visual Basic for Applications (VBA) editor.
' 3. In the Project Explorer window (usually on the left side), right-click on any of the items under "VBAProject (YourWorkbookName.xlsm)" and select Insert > Module. This creates a new module.
' 4. Copy the entire code for both `AutoColour` and `MarkColumnsWithMissingValues` macros (including this docstring) and paste it into the newly created module.
' 5. Press Ctrl + S to save your workbook. If it's not already a .xlsm (macro-enabled workbook), you will be prompted to save it as such.
' 6. Close the VBA editor and return to your Excel workbook.
' 7. Before running the macro, select the cells in your Excel sheet that you want to color.
' 8. To run the `AutoColour` macro, you can do the following: Press Alt + F8, select `AutoColour` from the list, and click "Run."
'
' Notes:
' - Ensure your workbook is saved before running the macro to prevent any loss of data.
' - Macros may be disabled in Excel by default. If the macro doesn't run, you might need to enable macros in Excel's Trust Center settings (File > Options > Trust Center > Trust Center Settings > Macro Settings).
'

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

    ' After processing individual cells, check for and mark columns with missing values
    'MarkColumnsWithMissingValues rng

CleanUp:
    ' Re-enable Excel updates
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub


Sub MarkColumnsWithMissingValues(rng As Range)
    Dim col As Range
    Dim cell As Range
    Dim i As Long
    Dim missingValuesColor As Long
    
    ' Define color for missing values (empty cells)
    missingValuesColor = RGB(255, 235, 156) ' Light yellow, adjust as needed
    
    ' Iterate through each column in the range
    For Each col In rng.Columns
        Dim hasFoundFilledCell As Boolean
        hasFoundFilledCell = False ' Reset for each new column
        
        ' Check each cell in the column
        For i = 1 To col.Rows.Count
            Set cell = col.Cells(i, 1)
            If Not IsEmpty(cell.Value) Then
                hasFoundFilledCell = True ' Mark that we've found a filled cell
            ElseIf hasFoundFilledCell And IsEmpty(cell.Value) Then
                ' If we've previously found a filled cell and the current cell is empty, color it
                cell.Interior.Color = missingValuesColor
            End If
        Next i
    Next col
End Sub

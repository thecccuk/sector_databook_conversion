' Macro to convert data from the "By measure" sheets to the CB7 sector databook format
' Author: Sam Van Stroud

' Configuration
' -----------------------------------------------------------------------------------------------
' Define configurable constants for easy modification and readability
Const SRC_TITLE_ROW As Integer = 1       ' Row number where titles are located in source sheet
Const DST_TITLE_ROW As Integer = 1       ' Row number where titles will be placed in destination sheet
Const START_YEAR As Long = 2015          ' Starting year for the data series
Const END_YEAR As Integer = 2050         ' Ending year for the data series
Const NUM_YEARS As Integer = END_YEAR - START_YEAR + 1  ' Total number of years in the data series

' Constants for specific sheet names and sector
Const SOURCE_SHEET_NAME As String = "By measure"  ' Name of the source sheet
Const SECTOR_NAME As String = "Waste"             ' Name of the sector being processed

' Define pathway constants for categorizing data
Const BASELINE As String = "Baseline"
Const BALANCED As String = "Balanced Pathway"
Const ADDITIONAL_ACTION As String = "Additional Action Pathway"
' -----------------------------------------------------------------------------------------------
' Compute column indices in the source sheet
Private SRC_COL_COUNTRY As Integer
Private SRC_COL_SUBSECTOR As Integer
Private SRC_COL_MEASURE_NAME As Integer
Private SRC_COL_MEASURE_VARIABLE As Integer
Private SRC_COL_VARIABLE_UNIT As Integer

' Initialize the above column indices
Private Sub InitializeColumnIndices()
    Dim src_ws As Worksheet
    Set src_ws = Worksheets(SOURCE_SHEET_NAME)
    SRC_COL_COUNTRY = get_index(src_ws, "Country")
    SRC_COL_SUBSECTOR = get_index(src_ws, "Subsector")
    SRC_COL_MEASURE_NAME = get_index(src_ws, "Measure Name")
    SRC_COL_MEASURE_VARIABLE = get_index(src_ws, "Measure Variable")
    SRC_COL_VARIABLE_UNIT = get_index(src_ws, "Variable Unit")
End Sub

' ------------------------------------------------------------
' Main subroutine to initiate the data conversion process
Sub Main()

    ' Print a start message to the immediate window
    Debug.Print (vbNewLine & "START CONVERSION...")
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    InitializeColumnIndices

    ' Retrieve a reference to the source worksheet containing the data
    Dim src_ws As Worksheet
    Set src_ws = Worksheets(SOURCE_SHEET_NAME)

    ' Check if the source sheet has all the required columns
    If Not check_source_sheet(src_ws) Then
        Debug.Print ("ERROR: Source sheet does not contain all required columns")
        Exit Sub
    End If

    ' Create a collection to hold references to the output sheets for each pathway
    Dim dst_wss As Collection
    Set dst_wss = New Collection
    ' Add new worksheets for each pathway to the collection
    dst_wss.Add create_sd_sheet(START_YEAR, END_YEAR, "Baseline data"), BASELINE
    dst_wss.Add create_sd_sheet(START_YEAR, END_YEAR, "BP Measure level data"), BALANCED
    dst_wss.Add create_sd_sheet(START_YEAR, END_YEAR, "AAP Measure level data"), ADDITIONAL_ACTION
    dst_wss.Add create_sd_sheet(START_YEAR, END_YEAR, "Aggregate data"), "Aggregate"

    ' Initialize a collection to track the current row for data entry in each output sheet
    Dim dst_row As Collection
    Set dst_row = New Collection
    dst_row.Add DST_TITLE_ROW + 1, BASELINE
    dst_row.Add DST_TITLE_ROW + 1, BALANCED
    dst_row.Add DST_TITLE_ROW + 1, ADDITIONAL_ACTION
    dst_row.Add DST_TITLE_ROW + 1, "Aggregate"

    ' Iterate through all the rows in the source sheet and copy the data to the output sheets
    Dim src_row As Range
    For Each src_row In src_ws.Rows(SRC_TITLE_ROW + 1).Resize(src_ws.UsedRange.Rows.Count - SRC_TITLE_ROW, 1).Cells
        Set dst_row = copy_row(src_row, src_ws, dst_wss, dst_row)
    Next src_row

    ' Special handling for the baseline data: remove the "Measure ID" and "Measure Name" columns
    RemoveColumnsFromSheet dst_wss(BASELINE)
    RemoveColumnsFromSheet dst_wss("Aggregate")

    ' Rename the "Measure Variable" column
    dst_wss(BASELINE).Cells(DST_TITLE_ROW, find_col(dst_wss(BASELINE).Rows(DST_TITLE_ROW), "Measure Variable").Column).Value = "Baseline Variable"
    dst_wss("Aggregate").Cells(DST_TITLE_ROW, find_col(dst_wss("Aggregate").Rows(DST_TITLE_ROW), "Measure Variable").Column).Value = "Aggregate Variable"

    ' Autofit the columns in each output sheet for better presentation
    Dim ws As Worksheet
    For Each ws In dst_wss
        ws.Cells.EntireColumn.AutoFit
    Next ws

    ' Print a completion message to the immediate window
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Debug.Print ("DONE")
End Sub
' ------------------------------------------------------------


' Copy a single row of data from the source sheet to the relevant output sheet
Function copy_row(src_row As Range, src_ws As Worksheet, dst_wss As Collection, dst_row As Collection) As Collection
    Debug.Print ("Copying row " & src_row.Row)
    ' get the index of this row
    Dim src_row_index As Integer: src_row_index = src_row.Row

    ' Determine the pathway and corresponding destination worksheet for each row
    Dim pathway As String: pathway = lookup(src_ws, "Pathway", src_row_index)
    Dim dst_ws As Worksheet: Set dst_ws = dst_wss(pathway)
    dst_row_idx = dst_row(pathway)

    ' Copy data from the source to the destination worksheet
    ' This includes subsector, measure name, variable, and unit
    dst_ws.Cells(dst_row_idx, "C").Value = SECTOR_NAME
    dst_ws.Cells(dst_row_idx, "B").Value = src_row.Cells(1, SRC_COL_COUNTRY).Value
    dst_ws.Cells(dst_row_idx, "D").Value = src_row.Cells(1, SRC_COL_SUBSECTOR).Value
    dst_ws.Cells(dst_row_idx, "E").Value = src_row.Cells(1, SRC_COL_MEASURE_NAME).Value
    dst_ws.Cells(dst_row_idx, "F").Value = src_row.Cells(1, SRC_COL_MEASURE_VARIABLE).Value
    dst_ws.Cells(dst_row_idx, "G").Value = src_row.Cells(1, SRC_COL_VARIABLE_UNIT).Value

    ' Copy the actual time series data for the current row
    Dim src_ts_start As Range: Set src_ts_start = src_ws.Rows(SRC_TITLE_ROW).Find(What:=START_YEAR, LookIn:=xlValues, LookAt:=xlWhole)
    Dim dst_ts_start As Range: Set dst_ts_start = dst_ws.Rows(DST_TITLE_ROW).Find(What:=START_YEAR, LookIn:=xlValues, LookAt:=xlWhole)
    Dim src_range As Range: Set src_range = src_ws.Range(src_ws.Cells(src_row_index, src_ts_start.Column), src_ws.Cells(src_row_index, src_ts_start.Column).Offset(0, NUM_YEARS - 1))
    Dim dst_range As Range: Set dst_range = dst_ws.Range(dst_ws.Cells(dst_row_idx, dst_ts_start.Column), dst_ws.Cells(dst_row_idx, dst_ts_start.Column).Offset(0, NUM_YEARS - 1))
    dst_range.Value = src_range.Value

    ' Update the row index for the next data entry in the destination sheet
    dst_row.Remove pathway
    dst_row.Add dst_row_idx + 1, pathway

    Set copy_row = dst_row
End Function

Function get_index(ws As Worksheet, columnName As String) As Integer
    Dim foundRange As Range
    Set foundRange = ws.Rows(SRC_TITLE_ROW).Find(What:=columnName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundRange Is Nothing Then
        get_index = foundRange.Column
    Else
        Debug.Print "Error: Column '" & columnName & "' not found in worksheet " & ws.Name
        get_index = -1 ' Return -1 or another appropriate value to indicate the column wasn't found
    End If
End Function


Function lookup(ws As Worksheet, columnName As String, rowIndex As Integer) As Variant
    Dim targetColumn As Range
    Set targetColumn = ws.Rows(SRC_TITLE_ROW).Find(What:=columnName, LookIn:=xlValues, LookAt:=xlWhole)
    lookup = ws.Cells(rowIndex, targetColumn.Column).Value
End Function


' Check the source sheet has all the required columns
Function check_source_sheet(src_ws As Worksheet) As Boolean
    ' Define the required column headers for the source sheet
    Dim required_columns() As Variant
    required_columns = Array("Pathway", "Country", "Subsector", "Measure Name", "Measure Variable", "Variable Unit")

    ' Check if all the required columns are present in the source sheet
    Dim col As Variant
    For Each col In required_columns
        If IsError(Application.Match(col, src_ws.Rows(SRC_TITLE_ROW).Value, 0)) Then
            Debug.Print ("Missing column: " & col)
            check_source_sheet = False
            Exit Function
        End If
    Next col

    ' Check all the years in the time series are present
    Dim year As Integer
    For year = START_YEAR To END_YEAR
        If IsError(Application.Match(year, src_ws.Rows(SRC_TITLE_ROW).Value, 0)) Then
            Debug.Print ("Missing year: " & year)
            check_source_sheet = False
            Exit Function
        End If
    Next year

    ' Return True if all required columns are present
    check_source_sheet = True
End Function



' Function to create a new worksheet with a specified name
' If a sheet with the same name exists, it returns that sheet instead
Function create_new_sheet(name As String) As Worksheet
    Dim ws As Worksheet

    ' Loop through existing worksheets to check if sheet already exists
    For Each ws In Worksheets
        If ws.Name = name Then
            Set create_new_sheet = ws
            Exit Function
        End If
    Next ws

    ' Create and return a new sheet with the specified name if it doesn't exist
    Set ws = Worksheets.Add
    ws.Name = name
    Set create_new_sheet = ws
End Function

' Function to create a new worksheet for sector databook
' Sets up the worksheet with predefined column headers and formatting
Function create_sd_sheet(startDate As Integer, endDate As Integer, name As String) As Worksheet
    ' Create a new worksheet with the specified name
    Dim ws As Worksheet: Set ws = create_new_sheet(name)

    ' Define the headers for the new sheet's columns
    Dim columnHeaders() As Variant
    columnHeaders = Array("Measure ID", "Country", "Sector", "Subsector", "Measure Name", "Measure Variable", "Variable Unit")

    ' Write the column headers to the designated title row in the worksheet
    Dim headerRange As Range: Set headerRange = ws.Range("A" & DST_TITLE_ROW).Resize(1, UBound(columnHeaders) + 1)
    headerRange.Value = columnHeaders

    ' Initialize an offset for yearly time series data columns
    Dim columnOffset As Integer: columnOffset = UBound(columnHeaders) + 2

    ' Add a column for each year in the range from startDate to endDate
    Dim currentYear As Integer
    For currentYear = startDate To endDate
        ws.Cells(DST_TITLE_ROW, columnOffset).Value = currentYear
        columnOffset = columnOffset + 1
    Next currentYear

    ' Get the index of the last used column for formatting purposes
    Dim lastColumn As Integer: lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Set font style and size for the entire worksheet
    With ws.Cells.Font
        .Name = "Century Gothic"
        .Size = 10
    End With

    ' Format the header row with bold font and a light blue background
    With ws.Rows(DST_TITLE_ROW).Resize(1, lastColumn)
        .Font.Bold = True
        .Interior.Color = RGB(173, 216, 230) ' Light blue color
    End With

    ' Autofit the columns for better readability
    ws.Cells.EntireColumn.AutoFit

    ' Return the newly created and formatted worksheet
    Set create_sd_sheet = ws
End Function

' Function to find a specific column by name within a given range
' Raises an error if the column is not found
Function find_col(cols As Range, col_name As String) As Range
    Dim col As Range

    ' Loop through each cell in the range to find the matching column name
    For Each col In cols.Cells
        If col.Value = col_name Then
            Set find_col = col
            Exit Function
        End If
    Next col

    ' Raise a custom error if the column is not found
    Err.Raise 1000, Description:="Could not find column " & col_name
End Function

' Remove the "Measure ID" and "Measure Name" columns from the given worksheet
Function RemoveColumnsFromSheet(ws As Worksheet)
    ' Find and delete the "Measure ID" column
    Dim measure_id_col As Range
    Set measure_id_col = find_col(ws.Rows(DST_TITLE_ROW), "Measure ID")
    If Not measure_id_col Is Nothing Then
        measure_id_col.EntireColumn.Delete
    End If

    ' Find and delete the "Measure Name" column
    Dim measure_name_col As Range
    Set measure_name_col = find_col(ws.Rows(DST_TITLE_ROW), "Measure Name")
    If Not measure_name_col Is Nothing Then
        measure_name_col.EntireColumn.Delete
    End If
End Function
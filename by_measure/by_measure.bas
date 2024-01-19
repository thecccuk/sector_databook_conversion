' Macro to convert data from the "By measure" sheets to the CB7 sector databook format
' Author: Sam Van Stroud

' Define configurable constants for easy modification and readability
Const SRC_TITLE_ROW As Integer = 2       ' Row number where titles are located in source sheet
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

' ------------------------------------------------------------
' Function to check if a cell marks the start of a time series
' Time series is expected to start with START_YEAR and increment yearly
Function is_time_series_start(c As Range) As Boolean
    ' Check if the cell is numeric and matches the START_YEAR
    If Not IsNumeric(c.Value) Or c.Value <> CStr(START_YEAR) Then
        is_time_series_start = False
        Exit Function
    End If

    ' Verify that each subsequent cell increments by one year up to END_YEAR
    Dim i As Integer
    For i = 1 To END_YEAR - START_YEAR
        If c.Offset(0, i).Value <> START_YEAR + i Then
            is_time_series_start = False
            Exit Function
        End If
    Next i

    is_time_series_start = True
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

' Function to copy a time series from a source sheet to a collection of destination sheets
' Each destination sheet corresponds to a different data pathway
Function copy_time_series(c As Range, src_ws As Worksheet, dst_wss As Collection, dst_row As Collection) As Collection
    ' Retrieve the country name from the source sheet
    Dim country As String
    country = c.Offset(-1, 0).Value

    ' Find relevant columns in both source and destination sheets for data copying
    ' The following variables store references to these columns
    Dim dst_ws As Worksheet: Set dst_ws = dst_wss(BALANCED)
    Dim dst_country_col as Range: Set dst_country_col = find_col(dst_ws.Rows(DST_TITLE_ROW), "Country")
    Dim dst_sector_col as Range: Set dst_sector_col = find_col(dst_ws.Rows(DST_TITLE_ROW), "Sector")
    Dim src_subsector_col as Range: Set src_subsector_col = find_col(c.Parent.UsedRange.Rows(SRC_TITLE_ROW), "Subsector")
    Dim dst_subsector_col as Range: Set dst_subsector_col = find_col(dst_ws.Rows(DST_TITLE_ROW), "Subsector")
    Dim src_measure_name_col as Range: Set src_measure_name_col = find_col(c.Parent.UsedRange.Rows(SRC_TITLE_ROW), "Measure Name")
    Dim dst_measure_name_col as Range: Set dst_measure_name_col = find_col(dst_ws.Rows(DST_TITLE_ROW), "Measure Name")
    Dim src_measure_variable_col as Range: Set src_measure_variable_col = find_col(c.Parent.UsedRange.Rows(SRC_TITLE_ROW), "Measure Variable")
    Dim dst_measure_variable_col as Range: Set dst_measure_variable_col = find_col(dst_ws.Rows(DST_TITLE_ROW), "Measure Variable")
    Dim src_variable_unit_col as Range: Set src_variable_unit_col = find_col(c.Parent.UsedRange.Rows(SRC_TITLE_ROW), "Variable Unit")
    Dim dst_variable_unit_col as Range: Set dst_variable_unit_col = find_col(dst_ws.Rows(DST_TITLE_ROW), "Variable Unit")

    Dim dst_time_series_start_col as Range: Set dst_time_series_start_col = find_col(dst_ws.Rows(DST_TITLE_ROW), CStr(START_YEAR))
    Dim src_pathway_col as Range: Set src_pathway_col = find_col(c.Parent.UsedRange.Rows(SRC_TITLE_ROW), "Pathway")
    Dim pathway As String
    Dim row_idx As Integer

    ' Loop through each row in the source sheet's time series data
    Dim src_range As Range
    Dim dst_range As Range
    Dim i As Integer
    For i = SRC_TITLE_ROW+1 To c.Parent.UsedRange.Rows.Count
        ' Determine the pathway and corresponding destination worksheet for each row
        pathway = src_ws.Cells(i, src_pathway_col.Column).Value
        Set dst_ws = dst_wss(pathway)
        row_idx = dst_row(pathway)
  
        ' Copy data from the source to the destination worksheet
        ' This includes subsector, measure name, variable, and unit
        dst_ws.Cells(row_idx, dst_country_col.Column).Value = country
        dst_ws.Cells(row_idx, dst_sector_col.Column).Value = SECTOR_NAME
        dst_ws.Cells(row_idx, dst_subsector_col.Column).Value = src_ws.Cells(i, src_subsector_col.Column).Value
        dst_ws.Cells(row_idx, dst_measure_name_col.Column).Value = src_ws.Cells(i, src_measure_name_col.Column).Value
        dst_ws.Cells(row_idx, dst_measure_variable_col.Column).Value = src_ws.Cells(i, src_measure_variable_col.Column).Value
        dst_ws.Cells(row_idx, dst_variable_unit_col.Column).Value = src_ws.Cells(i, src_variable_unit_col.Column).Value

        ' Copy the actual time series data for the current row
        Set src_range = src_ws.Range(src_ws.Cells(i, c.Column), src_ws.Cells(i, c.Column).Offset(0, NUM_YEARS - 1))
        Set dst_range = dst_ws.Range(dst_ws.Cells(row_idx, dst_time_series_start_col.Column), dst_ws.Cells(row_idx, dst_time_series_start_col.Column).Offset(0, NUM_YEARS - 1))
        dst_range.Value = src_range.Value

        ' Update the row index for the next data entry in the destination sheet
        dst_row.Remove pathway
        dst_row.Add row_idx + 1, pathway
    Next i

    ' Return the updated collection of destination row indices
    Set copy_time_series = dst_row
End Function

' ------------------------------------------------------------
' Main subroutine to initiate the data conversion process
Sub Main()
    ' Print a start message to the immediate window
    Debug.Print (vbNewLine & "START CONVERSION...")

    ' Retrieve a reference to the source worksheet containing the data
    Dim src_ws As Worksheet
    Set src_ws = Worksheets(SOURCE_SHEET_NAME)

    ' Create a collection to hold references to the output sheets for each pathway
    Dim dst_wss As Collection
    Set dst_wss = New Collection
    ' Add new worksheets for each pathway to the collection
    dst_wss.Add create_sd_sheet(START_YEAR, END_YEAR, "Baseline data"), BASELINE
    dst_wss.Add create_sd_sheet(START_YEAR, END_YEAR, "BP Measure level data"), BALANCED
    dst_wss.Add create_sd_sheet(START_YEAR, END_YEAR, "AAP Measure level data"), ADDITIONAL_ACTION

    ' Initialize a collection to track the current row for data entry in each output sheet
    Dim dst_row As Collection
    Set dst_row = New Collection
    dst_row.Add DST_TITLE_ROW + 1, BASELINE
    dst_row.Add DST_TITLE_ROW + 1, BALANCED
    dst_row.Add DST_TITLE_ROW + 1, ADDITIONAL_ACTION

    ' Iterate through each cell in the title row of the source sheet to identify time series
    Dim c As Range
    For Each c In src_ws.Rows(SRC_TITLE_ROW).Cells
        ' Check if the current cell is the start of a time series
        If is_time_series_start(c) Then
            Debug.Print ("Found time series at " & c.Address & " For Country " & c.Offset(-1, 0).Value)
            ' Copy the identified time series to the corresponding output sheet
            Set dst_row = copy_time_series(c, src_ws, dst_wss, dst_row)
        End If
        ' Stop the loop if the last column is reached
        If c.Column = Columns("ZZ").Column Then Exit For
    Next c

    ' Special handling for the baseline data: remove the "Measure Name" column
    Dim blws As Worksheet: Set blws = dst_wss(BASELINE)
    Dim measure_name_col As Range: Set measure_name_col = find_col(blws.Rows(DST_TITLE_ROW), "Measure Name")
    measure_name_col.EntireColumn.Delete

    ' Autofit the columns in each output sheet for better presentation
    Dim ws As Worksheet
    For Each ws In dst_wss
        ws.Cells.EntireColumn.AutoFit
    Next ws

    ' Print a completion message to the immediate window
    Debug.Print ("DONE")
End Sub
' ------------------------------------------------------------

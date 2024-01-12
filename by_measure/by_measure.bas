
' Configurable constants
Const SRC_TITLE_ROW As Integer = 2
Const DST_TITLE_ROW As Integer = 1
Const START_YEAR As Long = 2015
Const END_YEAR As Integer = 2050
Const NUM_YEARS As Integer = END_YEAR - START_YEAR + 1

Const SOURCE_SHEET_NAME As String = "By Measure V2"
Const SECTOR_NAME As String = "Waste"

' Pathways
Const BASELINE As String = "Baseline"
Const BALANCED As String = "Balanced Pathway"
Const ADDITIONAL_ACTION As String = "Additional Action Pathway"


' ------------------------------------------------------------
Function is_time_series_start(c As Range) As Boolean
  '
  ' Check if a cell is the start of a time series
  '
  If Not IsNumeric(c.Value) Then
    is_time_series_start = False
    Exit Function
  ElseIf c.Value <> CStr(START_YEAR) Then
    is_time_series_start = False
    Exit Function
  End If

  ' this cell has the right start value, now check the rest of the row
  For i = 1 To END_YEAR - START_YEAR
    If c.Offset(0, i).Value <> START_YEAR + i Then
      is_time_series_start = False
      Exit Function
    End If
  Next i
  is_time_series_start = True
End Function




Function create_new_sheet(name As String) As Worksheet
  '
  ' Create a new sheet with the given name
  '
  Dim ws As Worksheet

  ' Check if a sheet with the specified name already exists
  For Each ws In Worksheets
    If ws.name = name Then
      Set create_new_sheet = ws
      Exit Function
     End If
  Next ws
  
  ' If it doesn't exist already, create a new sheet
  Set ws = Worksheets.Add
  ws.name = name
  Set create_new_sheet = ws
End Function


Function create_sd_sheet(startDate As Integer, endDate As Integer, name As String) As Worksheet
    '
    ' Create a new sheet for the sector databook
    '
    
    ' Create a new sheet
    Dim ws As Worksheet: Set ws = create_new_sheet(name)
    
    ' Define column headers
    Dim columnHeaders() As Variant
    columnHeaders = Array("Measure ID", "Country", "Sector", "Subsector", "Measure Name", "Measure Variable", "Variable Unit")
    
    ' Write column headers to the worksheet on DST_TITLE_ROW
    Dim headerRange As Range: Set headerRange = ws.Range("A" & DST_TITLE_ROW).Resize(1, UBound(columnHeaders) + 1)
    headerRange.Value = columnHeaders

    ' Set initial column offset for yearly time series
    Dim columnOffset As Integer: columnOffset = UBound(columnHeaders) + 2
    
    ' Loop through years and add columns
    Dim currentYear As Integer
    For currentYear = startDate To endDate
        ws.Cells(DST_TITLE_ROW, columnOffset).Value = currentYear
        columnOffset = columnOffset + 1
    Next currentYear
    
    ' Get the last column index
    Dim lastColumn As Integer: lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' set font and font size
    With ws.Cells.Font
        .Name = "Century Gothic"
        .Size = 10
    End With

    ' Format the font and background color for all time series columns
    with ws.Rows(DST_TITLE_ROW).Resize(1, lastColumn)
        .Font.Bold = True
        .Interior.Color = RGB(173, 216, 230) ' Light blue color
    End With

    ' Autofit columns for better visibility
    ws.Cells.EntireColumn.AutoFit

    ' Return the new worksheet
    Set create_sd_sheet = ws
End Function

Function find_col(cols As Range, col_name As String) As Range
  '
  ' Find the column index of a column with the given name
  '
  Dim col As Range
  For Each col In cols.Cells
    If col.Value = col_name Then
      Set find_col = col
      Exit Function
    End If
  Next col
  ' error if we didn't find the column
  Debug.Print ("ERROR: Could not find column " & col_name)
End Function

Function copy_time_series(c As Range, src_ws As Worksheet, dst_wss As Collection, dst_row As Collection) As Collection
  '
  ' Copy a time series from a "By Measure" sheet to the sector databook
  '
  ' get country as string
  Dim country As String
  country = c.Offset(-1, 0).Value

  ' automatically find relevant source and destination columns
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

  ' loop over rows in this time series
  Dim src_range As Range
  Dim dst_range As Range
  Dim i As Integer
  For i = SRC_TITLE_ROW+1 To c.CurrentRegion.Rows.Count
    ' get the target worksheet based on the Pathway column in the source sheet
    pathway = src_ws.Cells(i, src_pathway_col.Column).Value
    Set dst_ws = dst_wss(pathway)
    row_idx = dst_row(pathway)
  
    ' copy subsector, measure name, measure variable, and variable unit
    dst_ws.Cells(row_idx, dst_country_col.Column).Value = country
    dst_ws.Cells(row_idx, dst_sector_col.Column).Value = SECTOR_NAME
    dst_ws.Cells(row_idx, dst_subsector_col.Column).Value = src_ws.Cells(i, src_subsector_col.Column).Value
    dst_ws.Cells(row_idx, dst_measure_name_col.Column).Value = src_ws.Cells(i, src_measure_name_col.Column).Value
    dst_ws.Cells(row_idx, dst_measure_variable_col.Column).Value = src_ws.Cells(i, src_measure_variable_col.Column).Value
    dst_ws.Cells(row_idx, dst_variable_unit_col.Column).Value = src_ws.Cells(i, src_variable_unit_col.Column).Value

    ' copy the time series data in this row
    Set src_range = src_ws.Range(src_ws.Cells(i, c.Column), src_ws.Cells(i, c.Column).Offset(0, NUM_YEARS - 1))
    Set dst_range = dst_ws.Range(dst_ws.Cells(row_idx, dst_time_series_start_col.Column), dst_ws.Cells(row_idx, dst_time_series_start_col.Column).Offset(0, NUM_YEARS - 1))
    dst_range.Value = src_range.Value

    ' increment the current dst row index
    dst_row.Remove pathway
    dst_row.Add row_idx+1, pathway
  Next i
  Set copy_time_series = dst_row
End Function



' ------------------------------------------------------------
Sub Main()
  '
  ' Main Macro
  '
  Debug.Print (vbNewLine & "NEW RUN2")

  ' save a reference to the source sheet
  Dim src_ws As Worksheet
  Set src_ws = Worksheets(SOURCE_SHEET_NAME)

  ' create the output sheets in the sector databook, one for each pathway
  Dim dst_wss As Collection
  Set dst_wss = New Collection
  dst_wss.Add create_sd_sheet(START_YEAR, END_YEAR, "Baseline data"), BASELINE
  dst_wss.Add create_sd_sheet(START_YEAR, END_YEAR, "BP Measure level data"), BALANCED
  dst_wss.Add create_sd_sheet(START_YEAR, END_YEAR, "AAP Measure level data"), ADDITIONAL_ACTION

  ' initialize the index of the current row to be written to for each output sheet
  Dim dst_row As Collection
  Set dst_row = New Collection
  dst_row.Add DST_TITLE_ROW+1, BASELINE
  dst_row.Add DST_TITLE_ROW+1, BALANCED
  dst_row.Add DST_TITLE_ROW+1, ADDITIONAL_ACTION

  ' loop through cells in the title row and look for a time series
  Dim c As Range
  For Each c In src_ws.Rows(SRC_TITLE_ROW).Cells
    ' check if this is the start of a time series
    If is_time_series_start(c) Then
      Debug.Print ("Found time series at " & c.Address & " For Country " & c.Offset(-1, 0).Value)
      ' copy the time series to the new sheet
      Set dst_row = copy_time_series(c, src_ws, dst_wss, dst_row)
    End If
    If c.Column = Columns("ZZ").Column Then Exit For
  Next c

  ' for the baseline, delete the "Measure Name" column
  Dim blws As Worksheet: Set blws = dst_wss(BASELINE)
  Dim measure_name_col As Range: Set measure_name_col = find_col(blws.Rows(DST_TITLE_ROW), "Measure Name")
  measure_name_col.EntireColumn.Delete

  ' Autofit each output sheet
  Dim ws As Worksheet
  For Each ws In dst_wss
    ws.Cells.EntireColumn.AutoFit
  Next ws

  Debug.Print ("DONE")
End Sub
' ------------------------------------------------------------

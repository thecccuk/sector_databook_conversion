Const TYPE_ROW_OFFSET As Integer = -3
Const VARIABLE_ROW_OFFSET As Integer = -2
Const UNITS_ROW_OFFSET As Integer = -1


Function get_time_series_info(c As Range) As Variant
  '
  ' Get the time series info for a cell
  '
  Dim info(1 To 3) As String
  info(1) = c.Offset(TYPE_ROW_OFFSET, 0).Value
  info(2) = c.Offset(VARIABLE_ROW_OFFSET, 0).Value
  info(3) = c.Offset(UNITS_ROW_OFFSET, 0).Value
  get_time_series_info = info
End Function


Function is_abatement(c As Range) As Boolean
  '
  ' Check if a cell is an abatement time series
  '
  Dim info As Variant
  info = get_time_series_info(c)
  If info(1) = ABATEMENT_STRING Then
    is_abatement = True
  Else
    is_abatement = False
  End If
End Function

Function copy_row(src As Worksheet, dst As Worksheet, src_row As Integer, dst_row As Integer, src_start_col As Integer, dst_start_col As String) As Boolean
    ' Function to copy a row from one sheet to another as values

    Dim src_range As Range
    Dim dst_range As Range

    ' Set reference to the source and dest range
    Set src_start = src.Cells(src_row, src_start_col)
    Set src_end = src_start.Offset(0, NUM_YEARS - 1)
    Set dst_start = dst.Cells(dst_row, dst_start_col)
    Set dst_end = dst_start.Offset(0, NUM_YEARS - 1)
    Set src_range = src.Range(src_start, src_end)
    Set dst_range = dst.Range(dst_start, dst_end)

    ' Copy the source range to the destination range as values
    src_range.Copy
    dst_range.PasteSpecial Paste:=xlPasteValues

    ' Clear the clipboard (optional)
    Application.CutCopyMode = False

    ' Indicate successful copy
    copy_row = True
End Function


Function copy_time_series(c As Range, ws As Worksheet, nextRow As Integer, country As String) As Integer
  '
  ' Copy a time series to a worksheet
  '
  Dim info As Variant
  info = get_time_series_info(c)

  ' loop over rows in this time series
  Dim i As Integer
  For i = 1 To 100 ' For now just hard code the rows
      ' check if column A in the src sheet contains the required pathway string as a substring
      If InStr(1, Cells(c.Offset(i, 0).Row, "A").Value, "Headwinds") <> 0 Then
        ' first copy the sector name and info
        ws.Cells(nextRow, "B").Value = country
        ws.Cells(nextRow, "C").Value = SECTOR_NAME
        ws.Cells(nextRow, "D").Value = Cells(c.Offset(i, 0).Row, "A").Value
        ws.Cells(nextRow, "E").Value = Cells(c.Offset(i, 0).Row, "B").Value
        ws.Cells(nextRow, "F").Value = info(2)
        ws.Cells(nextRow, "G").Value = info(3)

        ' copy the time series data
        copy_row ActiveSheet, ws, c.Offset(i, 0).Row, nextRow, c.Column, "H"

        ' increment the next row
        nextRow = nextRow + 1
      End If
  Next i
  copy_time_series = nextRow
End Function


Function process_country(country As String, ws As Worksheet, nextRow As Integer)
  '
  ' Process a single country sheet as specified by name
  '
  Debug.Print ("Processing " & country)

  ' find the country sheet
  Dim countrySheet As Worksheet
  For Each countrySheet In Worksheets
    If countrySheet.name = country Then
      Exit For
    End If
  Next countrySheet

  ' error if we can't find the country sheet
  If countrySheet Is Nothing Then
    Debug.Print ("ERROR: Could not find sheet for " & country)
    Exit Function
  End If

  ' activate the country sheet
  countrySheet.Activate

  ' loop through columns and look for a time series
  Dim c As Range
  For Each c In Rows(TITLE_ROW).Cells
    ' check if this is the start of a time series
    If is_time_series_start(c) And is_abatement(c) Then
      Debug.Print ("Found abatement time series at " & c.Address)
      ' copy the time series to the new sheet
      nextRow = copy_time_series(c, ws, nextRow, country)
    End If
    If c.Column = Columns("ZZ").Column Then Exit For
  Next
End Function



' ------------------------------------------------------------
Sub Main()
  '
  ' Main Macro
  '
  Debug.Print (vbNewLine & "NEW RUN")

  ' define array of contries
  Dim countries() As Variant
  countries = Array("UK", "Scotland", "Wales", "NI")

  ' set up the target worksheet
  Dim ws As Worksheet
  Set ws = create_new_sheet("SVS TEST")
  ' store the index of the next free row to copy to
  Dim nextRow As Integer
  nextRow = 2

  ' loop over countries
  Dim country As Variant
  For Each country In countries
    process_country CStr(country), ws, nextRow
  Next country
End Sub
' --------------
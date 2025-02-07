Sub FullAnalisisMacro()
    Dim csvFile As String
    Dim folderPath As String
    Dim wb As Workbook
    Dim fileName As String
    Dim newSheet As Worksheet
    Dim csvFullPath As String
    
    ActiveSheet.Name = "Analytics"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Analytics")
    
    ws.Range("U6").Value = "Rules"
    ws.Range("U7").Value = "Uhol"
    ws.Range("U8").Value = "Priemer"
    ws.Range("U10").Value = "Vzdialenost"
    
    ws.Range("V6").Value = "Min"
    ws.Range("V7").Value = 113
    ws.Range("V8").Value = 6.92
    ws.Range("V9").Value = 8.22
    ws.Range("V10").Value = 2.2
    
    ws.Range("W6").Value = "Max"
    ws.Range("W7").Value = 117
    ws.Range("W8").Value = 7.5
    ws.Range("W9").Value = 8.8
    ws.Range("W10").Value = 2.8
    
    ws.Range("X6").Value = "Alt"
    ws.Range("X7").Value = 0
    ws.Range("X10").Value = 0
    
    ws.Range("U6:X10").Font.Bold = True
    
    ' Add a button for "Color Rule" on cells O12:Q12
    Dim btn As Button
    Set btn = ws.Buttons.Add(Left:=ws.Range("U12").Left, Top:=ws.Range("U12").Top, Width:=ws.Range("U12:X12").Width, Height:=ws.Range("S12:V12").Height)
    btn.Caption = "Color Rule"
    btn.OnAction = "ClearAllColors"
    
    ' Prompt the user to select a folder
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    folderDialog.Title = "Select Folder Containing CSV Files"
    
    If folderDialog.Show = -1 Then
        ' User selected a folder
        folderPath = folderDialog.SelectedItems(1) & "\"
    Else
        ' User cancelled the dialog
        MsgBox "No folder selected. Exiting macro.", vbExclamation
        Exit Sub
    End If

    ' Find the first CSV file in the folder
    csvFile = Dir(folderPath & "*.csv")
    
    ' Loop through all CSV files in the folder
    Do While csvFile <> ""
        ' Get the full path of the CSV file
        csvFullPath = folderPath & csvFile

        ' Get the filename without the extension
        fileName = Left(csvFile, InStrRev(csvFile, ".") - 1)

        ' Create a new sheet in the current workbook with the CSV filename
        Set wb = ThisWorkbook
        Set newSheet = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        newSheet.Name = fileName

        ' Open the CSV file and import it with semicolon delimiter
        With newSheet.QueryTables.Add(Connection:="TEXT;" & csvFullPath, Destination:=newSheet.Range("A1"))
            .TextFileParseType = xlDelimited
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = True
            .TextFileCommaDelimiter = False
            .TextFileOtherDelimiter = False
            .TextFileColumnDataTypes = Array(1) ' Treat all columns as text
            .Refresh BackgroundQuery:=False
        End With

        ' Format column B as a normal number
        newSheet.Columns("B").NumberFormat = "0"

        ' Run the other macros after importing the data
        UpdateZmenaTable newSheet
        CreatePivotTable newSheet
        
        ' Move to the next CSV file
        csvFile = Dir
    Loop

    ' MsgBox "All CSV files have been processed."
    
    MainAnalytics newSheet
    
End Sub

Sub ClearAllColors()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Sheets
        ws.Cells.Interior.ColorIndex = xlNone ' Clear all cell colors
        If ws.Index > 1 Then Call ColorOutOfRange(ws)
    Next ws
       
End Sub

Sub UpdateZmenaTable(ws As Worksheet)
    Dim lastRow As Long
    Dim newRow As Long

    ' Delete columns
    ws.Columns("C:F").Delete

    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Insert new column B
    ws.Columns("B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ' Name column B as "Difference"
    ws.Range("B1").Value = "Difference"

    ' Add an empty row at row 2
    ws.Rows(2).Insert Shift:=xlDown

    ' Copy A3 to A2 and adjust A2 to the beginning of the hour
    ws.Range("A2").Value = Int(ws.Range("A3").Value) + TimeSerial(Hour(ws.Range("A3").Value), 0, 0)

    ' Add a new row at the end
    newRow = lastRow + 2
    ws.Rows(newRow).Insert Shift:=xlDown
    
    ' Set the date and time in the last row (A) as A2 + 8 hours
    ws.Range("A" & newRow).Value = ws.Range("A2").Value + TimeValue("08:00:00")

    ' Update lastRow after adding the new row
    lastRow = lastRow + 1

    ' Apply the formula in column B from B3 to the last row
    If Not IsEmpty(ws.Range("A2")) And Not IsEmpty(ws.Range("A1")) Then
        ws.Range("B2:B" & lastRow + 1).Formula = "=IFERROR(ROUND((A2-A1)*86400, 0), 0)"
    End If
    ' ws.Range("B2:B" & lastRow + 1).Formula = "=IFERROR(ROUND((A2-A1)*86400, 0), 0)"
    
    ' Format column B as Text
    ws.Columns("B").NumberFormat = "@"
    
    ' Name columns
    ws.Range("R1").Value = "Uhol A"
    ws.Range("S1").Value = "Uhol B"
    ws.Range("T1").Value = "Priemer A"
    ws.Range("U1").Value = "Priemer B"
    ws.Range("V1").Value = "Vzdielanost A"
    ws.Range("W1").Value = "Vzdielanost B"
    
    ' Apply the formula to find state
    ws.Range("R3:R" & lastRow).Formula2 = "=IF(OR(AND(D3=Analytics!$V$7, E3=Analytics!$V$7), SUM((D3:E3<Analytics!$T$7)+(D3:E3>Analytics!$U$7))<>COLUMNS(D3:E3)), ""OK"", ""NOK"")"
    ws.Range("S3:S" & lastRow).Formula2 = "=IF(SUM((F3:G3<Analytics!$T$7)+(F3:G3>Analytics!$U$7))<>COLUMNS(F3:G3), ""OK"", ""NOK"")"
    ws.Range("T3:T" & lastRow).Formula2 = "=IF(AND(SUM((H3:I3<Analytics!$T$8)+(H3:I3>Analytics!$U$8))=COLUMNS(H3:I3), SUM((H3:I3<Analytics!$T$9)+(H3:I3>Analytics!$U$9))=COLUMNS(H3:I3)), ""NOK"", ""OK"")"
    ws.Range("U3:U" & lastRow).Formula2 = "=IF(AND(SUM((J3:K3<Analytics!$T$8)+(J3:K3>Analytics!$U$8))=COLUMNS(J3:K3), SUM((J3:K3<Analytics!$T$9)+(J3:K3>Analytics!$U$9))=COLUMNS(J3:K3)), ""NOK"", ""OK"")"
    ws.Range("V3:V" & lastRow).Formula2 = "=IF(OR(AND(L3=Analytics!$V$10, M3=Analytics!$V$10), SUM((L3:M3<Analytics!$T$10)+(L3:M3>Analytics!$U$10))<>COLUMNS(L3:M3)), ""OK"", ""NOK"")"
    ws.Range("W3:W" & lastRow).Formula2 = "=IF(SUM((N3:O3<Analytics!$T$10)+(N3:O3>Analytics!$U$10))<>COLUMNS(N3:O3), ""OK"", ""NOK"")"

End Sub

Sub CreatePivotTable(ws As Worksheet)
    Dim lastRowT As Long
    Dim lastRowP As Long
    Dim pivotTable As pivotTable
    Dim pivotCache As pivotCache
    Dim pivotRange As Range

    ' Find the last row with data in column A and B
    lastRowT = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row - 1

    ' Define the range for the Pivot Table (Columns A and B, from 1st row to the last row)
    Set pivotRange = ws.Range("A1:B" & lastRowT)

    ' Create the Pivot Cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)

    ' Create the Pivot Table on the current sheet, starting at V2
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=ws.Range("Y2"), TableName:="DifferencePivotTable")

    ' Add "Difference" to Rows
    With pivotTable.PivotFields("Difference")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' Add "Sum of Difference" to Values
    With pivotTable.PivotFields("Difference")
        .Orientation = xlDataField
        .Function = xlSum
        .NumberFormat = "#,##0"
        .Name = "Sum of Difference"
    End With

    ' Auto-fit the columns of the Pivot Table for better visibility
    ws.Columns.AutoFit
    
    ' Find the last row of the pivot table in column Y
    lastRowP = ws.Cells(ws.Rows.Count, "Y").End(xlUp).Row
    
    ' Add the new table for production time, short downtime, and long downtime in D3:E5
    ws.Range("AB3").Value = "Production Time"
    ws.Range("AB4").Value = "Short Down Time 11-60s"
    ws.Range("AB5").Value = "Long Down Time 60s and more"
    ws.Range("AC2").Value = "Seconds"
    ws.Range("AC8").Value = "Count"
    
    ws.Range("AB9").Value = "OK State"
    ws.Range("AB10").Value = "NOK State"
    ws.Range("AB11").Value = "NOK State Uhol A"
    ws.Range("AB12").Value = "NOK State Uhol B"
    ws.Range("AB13").Value = "NOK State Priemer A"
    ws.Range("AB14").Value = "NOK State Priemer B"
    ws.Range("AB15").Value = "NOK State Vzdielanost A"
    ws.Range("AB16").Value = "NOK State Vzdielanost B"
    ws.Range("AB17").Value = "NOK State 3D A"
    ws.Range("AB18").Value = "NOK State 3D B"
    
    
    ' Apply bold formatting to headers
    ws.Range("AB3:AB5").Font.Bold = True
    ws.Range("AC2").Font.Bold = True
    ws.Range("AB9:AB16").Font.Bold = True
    ws.Range("AC8").Font.Bold = True

    ' Add the corresponding formulas for E3 to E5
    ws.Range("AC3").Formula = "=SUMIF(Y2:Y" & lastRowP & ", ""<=11"", Z2:Z" & lastRowP & ")"
    ws.Range("AC4").Formula = "=SUMIFS(Z2:Z" & lastRowP & ", Y2:Y" & lastRowP & ", "">11"", Y2:Y" & lastRowP & ", ""<=60"")"
    ws.Range("AC5").Formula = "=SUMIF(Y2:Y" & lastRowP & ", "">60"", Z2:Z" & lastRowP & ")"
     
    ws.Range("AC11").Formula = "=COUNTIF('" & ws.Name & "'!R3:R" & lastRowT & ", ""NOK"")"
    ws.Range("AC12").Formula = "=COUNTIF('" & ws.Name & "'!S3:S" & lastRowT & ", ""NOK"")"
    ws.Range("AC13").Formula = "=COUNTIF('" & ws.Name & "'!T3:T" & lastRowT & ", ""NOK"")"
    ws.Range("AC14").Formula = "=COUNTIF('" & ws.Name & "'!U3:U" & lastRowT & ", ""NOK"")"
    ws.Range("AC15").Formula = "=COUNTIF('" & ws.Name & "'!V3:V" & lastRowT & ", ""NOK"")"
    ws.Range("AC16").Formula = "=COUNTIF('" & ws.Name & "'!W3:W" & lastRowT & ", ""NOK"")"
    ws.Range("AC17").Formula = "=COUNTIF('" & ws.Name & "'!P3:P" & lastRowT & ", ""NOK"")"
    ws.Range("AC18").Formula = "=COUNTIF('" & ws.Name & "'!Q3:Q" & lastRowT & ", ""NOK"")"
    
    ws.Range("AC10").Formula = "=AC11 + AC12 + AC13 + AC14 + AC15 + AC16 + AC17 + AC18"
    ws.Range("AC9").Formula = "=ROWS('" & ws.Name & "'!A3:A" & lastRowT & ")-AC10"
    
    ' Auto-fit the columns for better visibility
    ws.Columns("AB:AC").AutoFit
    
End Sub

Sub ColorOutOfRange(ws As Worksheet)
    Dim lastRow As Long
    Dim cell As Range
    Dim rngD_G As Range, rngH_K As Range, rngL_O As Range
    
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row - 1

    ' Define the range for D3:G to the last row
    Set rngD_G = ws.Range("D3:G" & lastRow)
    
    lowerBoundD_G = ThisWorkbook.Sheets("Analytics").Range("T7").Value
    upperBoundD_G = ThisWorkbook.Sheets("Analytics").Range("U7").Value
    altBoundD_G = ThisWorkbook.Sheets("Analytics").Range("V7").Value
    
    ' Loop through each cell in D3:G
    For Each cell In rngD_G
        If (cell.Value < lowerBoundD_G Or cell.Value > upperBoundD_G) And cell.Value <> altBoundD_G Then
            cell.Interior.Color = RGB(255, 0, 0) ' Red color for cells out of range
        Else
            cell.Interior.ColorIndex = xlNone ' No fill if within range
        End If
    Next cell

    ' Define the range for H3:K to the last row
    Set rngH_K = ws.Range("H3:K" & lastRow)
    
    lowerBoundAH_K = ThisWorkbook.Sheets("Analytics").Range("T8").Value
    upperBoundAH_K = ThisWorkbook.Sheets("Analytics").Range("U8").Value
    lowerBoundBH_K = ThisWorkbook.Sheets("Analytics").Range("T9").Value
    upperBoundBH_K = ThisWorkbook.Sheets("Analytics").Range("U9").Value
    
    ' Loop through each cell in H3:K
    For Each cell In rngH_K
        ' Check if the cell value is outside both ranges [6.92-7.5] or [8.22-8.8]
        If (cell.Value < lowerBoundAH_K Or cell.Value > upperBoundAH_K) And (cell.Value < lowerBoundBH_K Or cell.Value > upperBoundBH_K) Then
            ' If the cell value is out of range, color it red
            cell.Interior.Color = RGB(255, 0, 0) ' Red color for cells out of range
        Else
            ' Reset the cell color if it's within the valid ranges
            cell.Interior.ColorIndex = xlNone ' No fill if within range
        End If
    Next cell

    ' Define the range for L3:O to the last row
    Set rngL_O = ws.Range("L3:O" & lastRow)
    
    lowerBoundL_O = ThisWorkbook.Sheets("Analytics").Range("T10").Value
    upperBoundL_O = ThisWorkbook.Sheets("Analytics").Range("U10").Value
    altBoundL_O = ThisWorkbook.Sheets("Analytics").Range("V10").Value
    
    ' Loop through each cell in L3:O
    For Each cell In rngL_O
        If (cell.Value < lowerBoundL_O Or cell.Value > upperBoundL_O) And cell.Value <> altBoundL_O Then
            ' If the cell value is out of the range [2.2-2.8], color it red
            cell.Interior.Color = RGB(255, 0, 0) ' Red color for cells out of range
        Else
            ' Reset the cell color if it's within the range
            cell.Interior.ColorIndex = xlNone ' No fill if within range
        End If
    Next cell
    
    ' Define the range for R3:R to the last row
    Set rngR = ws.Range("P3:W" & lastRow)
    
    ' Loop through each cell in R3:T
    For Each cell In rngR
        If cell.Value = "NOK" Then
            cell.Interior.Color = RGB(255, 0, 0) ' Red color for NOK cells
        Else
            cell.Interior.ColorIndex = xlNone ' No fill if not NOK
        End If
    Next cell

End Sub

Sub MainAnalytics(ws As Worksheet)
    Dim i As Integer
    
    ' Activate the first sheet before the listing
    ThisWorkbook.Sheets(1).Activate
    
    ' Set headers in the first row
    With ThisWorkbook.ActiveSheet
        .Range("A1").Value = "Date&Zmena"
        .Range("B1").Value = "Production Time"
        .Range("C1").Value = "Short Down Time 11-60s"
        .Range("D1").Value = "Long Down Time 60s and more"
        .Range("E1").Value = "Short and Long Down Time"
        
        .Range("F1").Value = "NOK State Uhol A"
        .Range("G1").Value = "NOK State Uhol B"
        .Range("H1").Value = "NOK State Priemer A"
        .Range("I1").Value = "NOK State Priemer B"
        .Range("J1").Value = "NOK State Vzdielanost A"
        .Range("K1").Value = "NOK State Vzdielanost B"
        .Range("L1").Value = "NOK State 3D A"
        .Range("M1").Value = "NOK State 3D B"
        
        .Range("N1").Value = "OK State"
        .Range("O1").Value = "NOK State"
        .Range("Q1").Value = "All Produced"
        .Range("R1").Value = "NonProduced"
        .Range("S1").Value = "Efficienty"
        .Range("U2").Value = "Cielovy CT"
        .Range("U3").Value = "Target"
        .Range("V2").Value = "10"
        .Range("U16").Value = "OK State Sum"
        .Range("U17").Value = "NOK State Sum"
        .Range("V3").Value = "=60/V2*60*8"
        .Range("A1:S1").Font.Bold = True
        .Range("U2:X17").Font.Bold = True
    End With
    
    i = 2 ' Start writing in cells A2 and B2
    
    ' Set Excel to automatic calculation mode
    Application.Calculation = xlAutomatic
    
    ' Loop through each sheet starting from the second one
    For Each ws In ThisWorkbook.Sheets
        If ws.Index > 1 Then
            ' Write sheet name in column A with hyperlink
            ThisWorkbook.ActiveSheet.Hyperlinks.Add Anchor:=ThisWorkbook.ActiveSheet.Cells(i, 1), _
            Address:="", SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:=ws.Name
            
            ' Link cells dynamically
            ThisWorkbook.ActiveSheet.Cells(i, 2).Formula = "='" & ws.Name & "'!AC3"
            ThisWorkbook.ActiveSheet.Cells(i, 3).Formula = "='" & ws.Name & "'!AC4"
            ThisWorkbook.ActiveSheet.Cells(i, 4).Formula = "='" & ws.Name & "'!AC5"
            ThisWorkbook.ActiveSheet.Cells(i, 5).Formula = "=C" & i & "+D" & i
            
            ThisWorkbook.ActiveSheet.Cells(i, 6).Formula = "='" & ws.Name & "'!AC11"
            ThisWorkbook.ActiveSheet.Cells(i, 7).Formula = "='" & ws.Name & "'!AC12"
            ThisWorkbook.ActiveSheet.Cells(i, 8).Formula = "='" & ws.Name & "'!AC13"
            ThisWorkbook.ActiveSheet.Cells(i, 9).Formula = "='" & ws.Name & "'!AC14"
            ThisWorkbook.ActiveSheet.Cells(i, 10).Formula = "='" & ws.Name & "'!AC15"
            ThisWorkbook.ActiveSheet.Cells(i, 11).Formula = "='" & ws.Name & "'!AC16"
            ThisWorkbook.ActiveSheet.Cells(i, 12).Formula = "='" & ws.Name & "'!AC17"
            ThisWorkbook.ActiveSheet.Cells(i, 13).Formula = "='" & ws.Name & "'!AC18"
            
            ThisWorkbook.ActiveSheet.Cells(i, 14).Formula = "='" & ws.Name & "'!AC9"
            ThisWorkbook.ActiveSheet.Cells(i, 15).Formula = "='" & ws.Name & "'!AC10"
            
            ThisWorkbook.ActiveSheet.Cells(i, 17).Formula = "=N" & i & "+O" & i
            ThisWorkbook.ActiveSheet.Cells(i, 18).Formula = "=$V$3" & "-Q" & i
            ThisWorkbook.ActiveSheet.Cells(i, 19).Formula = "=Q" & i & "/$V$3*100"
        
            i = i + 1
        End If
    Next ws
    
    ' AutoFit the columns to adjust their width
    ThisWorkbook.ActiveSheet.Columns("A:X").AutoFit
    
    GraphCreate ThisWorkbook.ActiveSheet
    
    End Sub
    
Sub GraphCreate(ws As Worksheet)

    ' Activate the first sheet before the listing
    ThisWorkbook.Sheets(1).Activate
    
    ' Find the last row with data in column A
    lastRow = ThisWorkbook.ActiveSheet.Cells(ThisWorkbook.ActiveSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Set the range for the chart (columns A, B, and E)
    Set chartRange = Union(ThisWorkbook.ActiveSheet.Range("A1:A" & lastRow), _
                           ThisWorkbook.ActiveSheet.Range("B1:B" & lastRow), _
                           ThisWorkbook.ActiveSheet.Range("C1:C" & lastRow), _
                           ThisWorkbook.ActiveSheet.Range("D1:D" & lastRow))
    
    ' Add a new chart object to the worksheet
    Set chartObj = ThisWorkbook.ActiveSheet.ChartObjects.Add(Left:=400, Width:=1300, Top:=50, Height:=500)
    
    ' Set the chart's data source and chart type
    With chartObj.Chart
        .SetSourceData Source:=chartRange
        .ChartType = xlColumnStacked100
        .HasTitle = True
        .ChartTitle.Text = "Production Time Monitoring"
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 24
        .Legend.Format.TextFrame2.TextRange.Font.Size = 24
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 204, 51)
        .SeriesCollection(1).Format.Fill.Transparency = 0.3
        
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
        .SeriesCollection(2).Format.Fill.Transparency = 0.3
        
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(255, 51, 0)
        .SeriesCollection(3).Format.Fill.Transparency = 0.7
        .ChartGroups(1).GapWidth = 10 ' Set gap width to 10%
    End With
    
    Set efficientyRange = Union(ThisWorkbook.ActiveSheet.Range("A1:A" & lastRow), _
                                ThisWorkbook.ActiveSheet.Range("S1:S" & lastRow))
        
    Set chartObj = ThisWorkbook.ActiveSheet.ChartObjects.Add(Left:=400, Width:=700, Top:=400, Height:=300)
        
    With chartObj.Chart
        .SetSourceData Source:=efficientyRange
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Efficienty"
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 24
        .Legend.Format.TextFrame2.TextRange.Font.Size = 24
        .HasLegend = True  ' Set legend at the bottom
        .Legend.Position = xlLegendPositionBottom
    End With
    
    
    ' Calculate the sum of ranges F and G from the second row to the last row
    ThisWorkbook.Sheets("Analytics").Range("V16").Value = "=SUM(N2:N" & lastRow & ")"
    ThisWorkbook.Sheets("Analytics").Range("V17").Value = "=SUM(O2:O" & lastRow & ")"

    ' Create the pie chart
    Set pieChartObj = ThisWorkbook.ActiveSheet.ChartObjects.Add(Left:=400, Width:=300, Top:=750, Height:=300)

    With pieChartObj.Chart
        .ChartType = xlPie
        .SetSourceData Source:=ThisWorkbook.Sheets("Analytics").Range("U16:U17")
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Sum of Ranges"
        .SeriesCollection(1).XValues = Array("OK", "NOK")
        .SeriesCollection(1).Values = "=Analytics!V16:V17"
        .SeriesCollection(1).ApplyDataLabels xlDataLabelsShowPercent
        .HasTitle = True
        .ChartTitle.Text = "OK vs NOK State"
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 24
        .Legend.Format.TextFrame2.TextRange.Font.Size = 24
        .SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = RGB(51, 204, 51) ' Green for OK State
        .SeriesCollection(1).Points(1).Format.Fill.Transparency = 0.3
        .SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = RGB(255, 51, 0) ' Red for NOK State
        .SeriesCollection(1).Points(2).Format.Fill.Transparency = 0.7
    End With
    

    ' Set the source data for the chart
    Set AnalyticsRange = Union(ThisWorkbook.ActiveSheet.Range("A2:A" & lastRow), _
                                ThisWorkbook.ActiveSheet.Range("N2:N" & lastRow), _
                                ThisWorkbook.ActiveSheet.Range("O2:O" & lastRow), _
                                ThisWorkbook.ActiveSheet.Range("R2:R" & lastRow))
    ' Add a new chart object to the worksheet
    Set chartObj = ThisWorkbook.ActiveSheet.ChartObjects.Add(Left:=400, Width:=1300, Top:=50, Height:=500)
    
    ' Configure the chart
    With chartObj.Chart
        .HasTitle = True
        .ChartTitle.Text = "Production Monitoring"
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 24
        .ApplyLayout (1)
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Format.TextFrame2.TextRange.Font.Size = 24
        
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "OK"
        .SeriesCollection(1).XValues = ThisWorkbook.ActiveSheet.Range("A2:A" & lastRow)
        .SeriesCollection(1).Values = ThisWorkbook.ActiveSheet.Range("N2:N" & lastRow)
        .SeriesCollection(1).ChartType = xlColumnStacked
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 204, 51) ' Green for OK State
        .SeriesCollection(1).Format.Fill.Transparency = 0.3

        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "NOK"
        .SeriesCollection(2).XValues = ThisWorkbook.ActiveSheet.Range("A2:A" & lastRow)
        .SeriesCollection(2).Values = ThisWorkbook.ActiveSheet.Range("O2:O" & lastRow)
        .SeriesCollection(2).ChartType = xlColumnStacked
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 51, 0) ' Red for NOK State
        .SeriesCollection(2).Format.Fill.Transparency = 0.5

        .SeriesCollection.NewSeries
        .SeriesCollection(3).Name = "DIF TO TARGET"
        .SeriesCollection(3).XValues = ThisWorkbook.ActiveSheet.Range("A2:A" & lastRow)
        .SeriesCollection(3).Values = ThisWorkbook.ActiveSheet.Range("R2:R" & lastRow)
        .SeriesCollection(3).ChartType = xlColumnStacked
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(21, 96, 130) ' Blue for "Target"
        .SeriesCollection(3).Format.Fill.Transparency = 0.5
        
        .ChartGroups(1).GapWidth = 10 ' Set gap width to 10%
        
    
    End With

End Sub

Sub FullAnalisisMacro()
    Dim csvFile As String
    Dim folderPath As String
    Dim wb As Workbook
    Dim fileName As String
    Dim newSheet As Worksheet
    Dim csvFullPath As String
    
    ActiveSheet.Name = "Analitics"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Analitics")
    
    ws.Range("P6").Value = "Rules"
    ws.Range("Q6").Value = "Min"
    ws.Range("R6").Value = "Max"
    ws.Range("P7").Value = "Uhol"
    ws.Range("Q7").Value = 113
    ws.Range("R7").Value = 117
    ws.Range("P8").Value = "Priemer"
    ws.Range("Q8").Value = 6.92
    ws.Range("R8").Value = 7.5
    ws.Range("Q9").Value = 8.22
    ws.Range("R9").Value = 8.8
    ws.Range("P10").Value = "Vzdialenost"
    ws.Range("Q10").Value = 2.2
    ws.Range("R10").Value = 2.8
    ws.Range("S6").Value = "Alt"
    ws.Range("S7").Value = 0
    ws.Range("S10").Value = 0.001
    ws.Range("O6:S10").Font.Bold = True
    
    ' Add a button for "Color Rule" on cells O12:Q12
    Dim btn As Button
    Set btn = ws.Buttons.Add(Left:=ws.Range("P12").Left, Top:=ws.Range("P12").Top, Width:=ws.Range("P12:S12").Width, Height:=ws.Range("P12:S12").Height)
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
    
    MainAnalitics newSheet
    
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
    ws.Range("A2").Value = Int(ws.Range("A3").Value) + TimeValue(Hour(ws.Range("A3").Value) & ":00:00")
    
    ' Add a new row at the end
    newRow = lastRow + 2
    ws.Rows(newRow).Insert Shift:=xlDown
    
    ' Set the date and time in the last row (A) as A2 + 8 hours
    ws.Range("A" & newRow).Value = ws.Range("A2").Value + TimeValue("08:00:00")

    ' Update lastRow after adding the new row
    lastRow = lastRow + 1

    ' Apply the formula in column B from B3 to the last row
    ws.Range("B2:B" & lastRow + 1).Formula = "=IFERROR(ROUND((A2-A1)*86400, 0), 0)"
    
    ' Format column B as Text
    ws.Columns("B").NumberFormat = "@"
    
    ' Name columns
    ws.Range("R1").Value = "State"
    ws.Range("S1").Value = "State A"
    ws.Range("T1").Value = "State B"
    
    ' Apply the formula to find state
    ws.Range("R3:R" & lastRow).Formula = "=IF(AND(OR(AND(MIN(D3:G3)>=Analitics!$Q$7, MAX(D3:G3)<=Analitics!$R$7), OR(D3=Analitics!$S$7, E3=Analitics!$R$7)), OR(AND(MIN(H3:K3)>=Analitics!$Q$8, MAX(H3:K3)<=Analitics!$R$8), AND(MIN(H3:K3)>=Analitics!$Q$9, MAX(H3:K3)<=Analitics!$R$9)), OR(AND(MIN(L3:O3)>=Analitics!$Q$10, MAX(L3:O3)<=Analitics!$R$10), OR(L3=Analitics!$S$10, M3=Analitics!$S$10)), P3=""OK"", Q3=""OK""), ""OK"", ""NOK"")"
    ws.Range("S3:S" & lastRow).Formula = "=IF(AND(OR(AND(MIN(D3:E3)>=Analitics!$Q$7, MAX(D3:E3)<=Analitics!$R$7), OR(D3=Analitics!$S$7, E3=Analitics!$R$7)), OR(AND(MIN(H3:I3)>=Analitics!$Q$8, MAX(H3:I3)<=Analitics!$R$8), AND(MIN(H3:I3)>=Analitics!$Q$9, MAX(H3:I3)<=Analitics!$R$9)), OR(AND(MIN(L3:M3)>=Analitics!$Q$10, MAX(L3:M3)<=Analitics!$R$10), OR(L3=Analitics!$S$10, M3=Analitics!$S$10)), P3=""OK""), ""OK"", ""NOK"")"
    ws.Range("T3:T" & lastRow).Formula = "=IF(AND(MIN(F3:G3)>=Analitics!$Q$7, MAX(F3:G3)<=Analitics!$R$7, OR(AND(MIN(J3:K3)>=Analitics!$Q$8, MAX(J3:K3)<=Analitics!$R$8), AND(MIN(J3:K3)>=Analitics!$Q$9, MAX(J3:K3)<=Analitics!$R$9)), MIN(N3:O3)>=Analitics!$Q$10, MAX(N3:O3)<=Analitics!$R$10, Q3=""OK""), ""OK"", ""NOK"")"

End Sub

Sub CreatePivotTable(ws As Worksheet)
    Dim lastRowT As Long
    Dim lastRowP As Long
    Dim pivotTable As pivotTable
    Dim pivotCache As pivotCache
    Dim pivotRange As Range

    ' Find the last row with data in column A and B
    lastRowT = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Define the range for the Pivot Table (Columns A and B, from 1st row to the last row)
    Set pivotRange = ws.Range("A1:B" & lastRowT)

    ' Create the Pivot Cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)

    ' Create the Pivot Table on the current sheet, starting at V2
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=ws.Range("V2"), TableName:="DifferencePivotTable")

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
    
    ' Find the last row of the pivot table in column T
    lastRowP = ws.Cells(ws.Rows.Count, "V").End(xlUp).Row
    
    ' Add the new table for production time, short downtime, and long downtime in D3:E5
    ws.Range("Y3").Value = "Production Time"
    ws.Range("Y4").Value = "Short Down Time 11-60s"
    ws.Range("Y5").Value = "Long Down Time 60s and more"
    ws.Range("Z2").Value = "Seconds"
    ws.Range("Y9").Value = "OK State"
    ws.Range("Y10").Value = "NOK State"
    ws.Range("Y11").Value = "NOK State A"
    ws.Range("Y12").Value = "NOK State B"
    ws.Range("Y13").Value = "NOK State Priemer A"
    ws.Range("Y14").Value = "NOK State Priemer B"
    ws.Range("Z8").Value = "Count"
    
    ' Apply bold formatting to headers
    ws.Range("Y3:Y5").Font.Bold = True
    ws.Range("Z2").Font.Bold = True
    ws.Range("Y9:Y14").Font.Bold = True
    ws.Range("Z8").Font.Bold = True

    ' Add the corresponding formulas for E3 to E5
    ws.Range("Z3").Formula = "=SUMIF(V2:V" & lastRowP & ", ""<=11"", W2:W" & lastRowP & ")"
    ws.Range("Z4").Formula = "=SUMIFS(W2:W" & lastRowP & ", V2:V" & lastRowP & ", "">11"", V2:V" & lastRowP & ", ""<=60"")"
    ws.Range("Z5").Formula = "=SUMIF(V2:V" & lastRowP & ", "">60"", W2:W" & lastRowP & ")"
    ws.Range("Z9").Formula = "=COUNTIF('" & ws.Name & "'!R3:R" & lastRowT & ", ""OK"")"
    ws.Range("Z10").Formula = "=COUNTIF('" & ws.Name & "'!R3:R" & lastRowT & ", ""NOK"")"
    ws.Range("Z11").Formula = "=COUNTIF('" & ws.Name & "'!S3:S" & lastRowT & ", ""NOK"")"
    ws.Range("Z12").Formula = "=COUNTIF('" & ws.Name & "'!T3:T" & lastRowT & ", ""NOK"")"
    ws.Range("Z13").Formula = "=SUMPRODUCT((H3:I" & lastRowT & "<6.92) + ((H3:I" & lastRowT & ">7.5) * (H3:I" & lastRowT & "<8.22)) + (H3:I" & lastRowT & ">8.8))"
    ws.Range("Z14").Formula = "=SUMPRODUCT((J3:K" & lastRowT & "<6.92) + ((J3:K" & lastRowT & ">7.5) * (J3:K" & lastRowT & "<8.22)) + (J3:K" & lastRowT & ">8.8))"
    
    ' Auto-fit the columns for better visibility
    ws.Columns("Y:Z").AutoFit
    
End Sub

Sub ColorOutOfRange(ws As Worksheet)
    Dim lastRow As Long
    Dim cell As Range
    Dim rngD_G As Range, rngH_K As Range, rngL_O As Range
    
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row - 1

    ' Define the range for D3:G to the last row
    Set rngD_G = ws.Range("D3:G" & lastRow)
    
    lowerBoundD_G = ThisWorkbook.Sheets("Analitics").Range("Q7").Value
    upperBoundD_G = ThisWorkbook.Sheets("Analitics").Range("R7").Value
    
    ' Loop through each cell in D3:G
    For Each cell In rngD_G
        If (cell.Value < lowerBoundD_G Or cell.Value > upperBoundD_G) And cell.Value <> 0 Then
            cell.Interior.Color = RGB(255, 0, 0) ' Red color for cells out of range
        Else
            cell.Interior.ColorIndex = xlNone ' No fill if within range
        End If
    Next cell

    ' Define the range for H3:K to the last row
    Set rngH_K = ws.Range("H3:K" & lastRow)
    
    lowerBoundAH_K = ThisWorkbook.Sheets("Analitics").Range("Q8").Value
    upperBoundAH_K = ThisWorkbook.Sheets("Analitics").Range("R8").Value
    lowerBoundBH_K = ThisWorkbook.Sheets("Analitics").Range("Q9").Value
    upperBoundBH_K = ThisWorkbook.Sheets("Analitics").Range("R9").Value
    
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
    
    lowerBoundL_O = ThisWorkbook.Sheets("Analitics").Range("Q10").Value
    upperBoundL_O = ThisWorkbook.Sheets("Analitics").Range("R10").Value
    
    ' Loop through each cell in L3:O
    For Each cell In rngL_O
        If (cell.Value < lowerBoundL_O Or cell.Value > upperBoundL_O) And cell.Value <> 0.001 Then
            ' If the cell value is out of the range [2.2-2.8], color it red
            cell.Interior.Color = RGB(255, 0, 0) ' Red color for cells out of range
        Else
            ' Reset the cell color if it's within the range
            cell.Interior.ColorIndex = xlNone ' No fill if within range
        End If
    Next cell
    
    ' Define the range for R3:R to the last row
    Set rngR = ws.Range("R3:T" & lastRow)
    
    ' Loop through each cell in R3:T
    For Each cell In rngR
        If cell.Value = "NOK" Then
            cell.Interior.Color = RGB(255, 0, 0) ' Red color for NOK cells
        Else
            cell.Interior.ColorIndex = xlNone ' No fill if not NOK
        End If
    Next cell

End Sub

Sub MainAnalitics(ws As Worksheet)
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
        .Range("F1").Value = "NOK State A"
        .Range("G1").Value = "NOK State B"
        .Range("H1").Value = "NOK State Priemer A"
        .Range("I1").Value = "NOK State Priemer B"
        .Range("J1").Value = "OK State"
        .Range("K1").Value = "NOK State"
        .Range("L1").Value = "All Produced"
        .Range("M1").Value = "NonProduced"
        .Range("N1").Value = "Efficienty"
        .Range("P2").Value = "Cielovy CT"
        .Range("P3").Value = "Target"
        .Range("Q2").Value = "10"
        .Range("P16").Value = "OK State Sum"
        .Range("P17").Value = "NOK State Sum"
        .Range("Q3").Value = "=60/$Q$2*60*8"
        .Range("A1:N1").Font.Bold = True
        .Range("P2:S10").Font.Bold = True
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
            ThisWorkbook.ActiveSheet.Cells(i, 2).Formula = "='" & ws.Name & "'!Z3"
            ThisWorkbook.ActiveSheet.Cells(i, 3).Formula = "='" & ws.Name & "'!Z4"
            ThisWorkbook.ActiveSheet.Cells(i, 4).Formula = "='" & ws.Name & "'!Z5"
            ThisWorkbook.ActiveSheet.Cells(i, 5).Formula = "=C" & i & "+D" & i
            ThisWorkbook.ActiveSheet.Cells(i, 6).Formula = "='" & ws.Name & "'!Z11"
            ThisWorkbook.ActiveSheet.Cells(i, 7).Formula = "='" & ws.Name & "'!Z12"
            ThisWorkbook.ActiveSheet.Cells(i, 8).Formula = "='" & ws.Name & "'!Z13"
            ThisWorkbook.ActiveSheet.Cells(i, 9).Formula = "='" & ws.Name & "'!Z14"
            ThisWorkbook.ActiveSheet.Cells(i, 10).Formula = "='" & ws.Name & "'!Z9"
            ThisWorkbook.ActiveSheet.Cells(i, 11).Formula = "='" & ws.Name & "'!Z10"
            ThisWorkbook.ActiveSheet.Cells(i, 12).Formula = "=J" & i & "+K" & i
            ThisWorkbook.ActiveSheet.Cells(i, 13).Formula = "=$Q$3" & "-L" & i
            ThisWorkbook.ActiveSheet.Cells(i, 14).Formula = "=L" & i & "/$Q$3*100"
        
            i = i + 1
        End If
    Next ws
    
    ' AutoFit the columns to adjust their width
    ThisWorkbook.ActiveSheet.Columns("A:S").AutoFit
    
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
                                ThisWorkbook.ActiveSheet.Range("N1:N" & lastRow))
        
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
    ThisWorkbook.Sheets("Analitics").Range("Q16").Value = "=SUM(F2:F" & lastRow & ")"
    ThisWorkbook.Sheets("Analitics").Range("Q17").Value = "=SUM(G2:G" & lastRow & ")"

    ' Create the pie chart
    Set pieChartObj = ThisWorkbook.ActiveSheet.ChartObjects.Add(Left:=400, Width:=300, Top:=750, Height:=300)

    With pieChartObj.Chart
        .ChartType = xlPie
        .SetSourceData Source:=ThisWorkbook.Sheets("Analitics").Range("P16:P17")
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Sum of Ranges"
        .SeriesCollection(1).XValues = Array("OK", "NOK")
        .SeriesCollection(1).Values = "=Analitics!Q16:Q17"
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
    Set analiticsRange = Union(ThisWorkbook.ActiveSheet.Range("A2:A" & lastRow), _
                                ThisWorkbook.ActiveSheet.Range("J2:J" & lastRow), _
                                ThisWorkbook.ActiveSheet.Range("K2:K" & lastRow), _
                                ThisWorkbook.ActiveSheet.Range("M2:M" & lastRow))
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
        .SeriesCollection(1).Values = ThisWorkbook.ActiveSheet.Range("J2:J" & lastRow)
        .SeriesCollection(1).ChartType = xlColumnStacked
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 204, 51) ' Green for OK State
        .SeriesCollection(1).Format.Fill.Transparency = 0.3

        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "NOK"
        .SeriesCollection(2).XValues = ThisWorkbook.ActiveSheet.Range("A2:A" & lastRow)
        .SeriesCollection(2).Values = ThisWorkbook.ActiveSheet.Range("K2:K" & lastRow)
        .SeriesCollection(2).ChartType = xlColumnStacked
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 51, 0) ' Red for NOK State
        .SeriesCollection(2).Format.Fill.Transparency = 0.5

        .SeriesCollection.NewSeries
        .SeriesCollection(3).Name = "DIF TO TARGET"
        .SeriesCollection(3).XValues = ThisWorkbook.ActiveSheet.Range("A2:A" & lastRow)
        .SeriesCollection(3).Values = ThisWorkbook.ActiveSheet.Range("M2:M" & lastRow)
        .SeriesCollection(3).ChartType = xlColumnStacked
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(21, 96, 130) ' Blue for "Target"
        .SeriesCollection(3).Format.Fill.Transparency = 0.5
        
        .ChartGroups(1).GapWidth = 10 ' Set gap width to 10%
        
    
    End With

End Sub



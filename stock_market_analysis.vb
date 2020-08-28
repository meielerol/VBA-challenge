Option Explicit

'stock market analysis
Sub stockAnalysis()
    'main data variables
    Dim ws As Worksheet
    Dim ticker As String
    Dim yearBeginPrice As Double
    Dim yearClosePrice As Double
    Dim yearDeltaPrice As Double
    Dim yearPercentChange As Double
    Dim yearVolume As Double
    
    'counters
    Dim i, j As Long
    Dim lastRow, lastCol As Long
    Dim summaryTableRow As Long
    Dim flag As Boolean
    
    'conversion variables
    Dim numDate As Long
    Dim stringDate As String
    Dim shortDate As Date
    
    For Each ws In Worksheets
    '--------CONVERT/CLEAN THE DATA PER WS------------
        'initialize lastRow and lastCol for new ws
        lastRow = 0
        lastCol = 0
        'count the rows and columns
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        'MsgBox ("# of rows: " & lastRow & vbNewLine & "# of cols: " & lastCol)
        
        'autofit all columns on ws
        ws.Cells.Columns.AutoFit
        'clean up the data: alphabetical, date
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Range("A1"), Order:=xlAscending
            .SortFields.Add Key:=ws.Range("B1"), Order:=xlAscending
            .SetRange ws.Range("A1:G" & lastRow)
            .Header = xlYes
            .Apply
        End With
        'convert string date to date
        ' For i = 2 To lastRow
            ' numDate = ws.Cells(i, 2).Value
            ' stringDate = CStr(numDate)
            ' shortDate = DateSerial(CInt(Left(stringDate, 4)), CInt(Mid(stringDate, 5, 2)), CInt(Right(stringDate, 2)))
            ' ws.Cells(i, 2).Value = shortDate
        ' Next i
    '----------------------------------------------------

    '--------COLLECT DATA FOR EACH TICKER AND RECORD-----
        'create columns for data to summarize
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        summaryTableRow = 2
        'loop through data for info
        For i = 2 To lastRow
            'set current ticker value
            ticker = ws.Cells(i, 1).Value
            
            'is ticker in the last of the current group? if so then...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                yearClosePrice = ws.Cells(i, 6).Value
                yearVolume = yearVolume + ws.Cells(i, 7).Value
                yearDeltaPrice = yearClosePrice - yearBeginPrice
                                
                'enter print statement
                'MsgBox ("Ticker: " & ticker & vbNewLine & "Year opening price: " & yearBeginPrice & vbNewLine & "Year closing price: " & yearClosePrice & vbNewLine & "% change: "&format(yearDeltaPrice/yearBeginPrice,"Percent")&vbNewLine & "Total stock volume: "&yearVolume)
                ws.Range("I" & summaryTableRow).Value = ticker
                ws.Range("J" & summaryTableRow).Value = FormatNumber(yearDeltaPrice, 2) 'not sure why this isn't working for format. also tried format(yearDeltaPrice,"#,##0.00#") with no luck among other variations
                    If yearDeltaPrice > 0 Then
                        ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4
                    ElseIf yearDeltaPrice < 0 Then
                        ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3
                    End If
                
                If yearBeginPrice = 0 Then
                    yearPercentChange = vbNull
                Else
                    yearPercentChange = yearDeltaPrice / yearBeginPrice
                End If
                
                ws.Range("K" & summaryTableRow).Value = Format(yearPercentChange, "Percent")
                ws.Range("L" & summaryTableRow).Value = yearVolume
                'increase summary row counter
                summaryTableRow = summaryTableRow + 1
                'flag for reset yearVolume total at end of for loop
                flag = 1
            'is the ticker the start of a new group? if so then...
            ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                yearBeginPrice = ws.Cells(i, 3).Value
                yearVolume = ws.Cells(i, 7).Value
            Else
                yearVolume = yearVolume + ws.Cells(i, 7).Value
                
            End If
            
            'reset the variables for a new ticker
            If flag = 1 Then
                yearVolume = 0
                yearBeginPrice = ""
                yearClosePrice = ""
                yearDeltaPrice = ""
            End If
        Next i
    '-----------------------------------------------------------------

    '----------SUMMARIZE/ANALYZE THE SUMMARY OF EACH TICKER-----------
        'count last row and column of the summary table
        lastRow = ws.Range("I1", ws.Range("I1").End(xlDown)).Rows.Count
        lastCol = ws.Range("L1", ws.Range("L1").End(xlToLeft)).Columns.Count
        'MsgBox ("# of rows: " & lastRow & vbNewLine & "# of cols: " & lastCol)
        
        'create columns for analysis to show
        'ws.Range("N1").Value = "Analysis"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        'create row labels
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        'for loop the summary info
        For i = 2 To lastRow
            'compare greatest percent increase
            If ws.Cells(i, 11).Value > 0 And ws.Cells(i, 11).Value > ws.Cells(2, 16).Value Then
                ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 16).Value = Format(ws.Cells(i, 11).Value, "Percent")
            'compare greatest percent decrease
            ElseIf ws.Cells(i, 11).Value < 0 And ws.Cells(i, 11).Value < ws.Cells(3, 16).Value Then
                ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 16).Value = Format(ws.Cells(i, 11).Value, "Percent")
            End If
            'compare greatest total volume
            If ws.Cells(i, 12).Value > ws.Cells(3, 16).Value Then
                ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 12).Value
            End If
        Next i
    '-----------------------------------------------------------------
     
        'autofit all columns on ws
        ws.Cells.Columns.AutoFit
    Next ws
End Sub
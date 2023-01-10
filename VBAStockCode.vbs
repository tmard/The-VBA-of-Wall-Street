Sub StockData()
' Loop code through every worksheet in workbook
Dim ws As Worksheet
Dim k As Integer

For k = 1 To Sheets.Count
    Set ws = Sheets(k)
    ws.Select

' Starting of code
' Title and format for titles
    Range("I1").Value = "Ticker"
    Range("I1").Font.Bold = True

    Range("J1").Value = "Yearly Change"
    Range("J1").Font.Bold = True

    Range("K1").Value = "Percent Change"
    Range("K1").Font.Bold = True

    Range("L1").Value = "Total Volume"
    Range("L1").Font.Bold = True
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(2, 15).Font.Bold = True
    
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(3, 15).Font.Bold = True
    
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(4, 15).Font.Bold = True
    
    Range("P1").Value = "Ticker"
    Range("P1").Font.Bold = True
    
    Range("Q1").Value = "Value"
    Range("Q1").Font.Bold = True


' Declare variables
    Dim tickername As String
    Dim i As Long
    Dim j As Integer
    Dim tickercounter As Integer
    Dim LastRowA As Long
    Dim LastRowJ As Long
    Dim StartDateValue As Variant
    Dim EndDateValue As Variant
    Dim YearlyChange As Variant
    Dim PercentChange As Variant
    Dim VolumeTotal As Variant


' Define and set variables
    LastRowA = Cells(Rows.Count, 1).End(xlUp).Row
    LastRowJ = Cells(Rows.Count, 10).End(xlUp).Row

    tickercounter = 2
    YearlyChange = 0
    VolumeTotal = 0
    

' Loop through each row
For i = 2 To LastRowA
    tickername = Cells(i, 1).Value
    
    ' Check if within the same ticker, if it is not, then the following will occur
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            ' Find unique variable for each ticker and place in Column I
            Range("I" & tickercounter).Value = tickername
         
            ' Setting EndDateValue for each ticker
            EndDateValue = Cells(i, 6).Value
        
            ' Define YearlyChange and place in Column J
            YearlyChange = EndDateValue - StartDateValue
            Range("J" & tickercounter).Value = YearlyChange
                If Range("J" & tickercounter).Value > 0 Then
                    Range("J" & tickercounter).Interior.ColorIndex = 4
                    
                ElseIf Range("J" & tickercounter).Value < 0 Then
                    Range("J" & tickercounter).Interior.ColorIndex = 3
                        
                Else
                    Range("J" & tickercounter).Interior.ColorIndex = 6
                        
                End If

            ' Find YearlyPercent with conditional formatting and place in Column K
            YearlyPercent = ((EndDateValue / StartDateValue) * 100) - 100
                YearlyPercent = Format(YearlyPercent, "0.00") + "%"
            Range("K" & tickercounter).Value = YearlyPercent
  
            ' Find total stock volume and place in Column L
            VolumeTotal = VolumeTotal + Cells(i, 7).Value
            Range("L" & tickercounter).Value = VolumeTotal
        
            ' Add one to the ticker each time, ensure VolumeTotal starts at 0 each time
            tickercounter = tickercounter + 1
            VolumeTotal = 0
   
            ' Find StartDateValue for each ticker
            ElseIf Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            StartDateValue = Cells(i, 3).Value

            Else
            VolumeTotal = VolumeTotal + Cells(i, 7).Value
        
            End If

    Next i
            
' Greatest % increase - Find maximum value in yearly change with associated ticker name
    Cells(2, 17).Value = WorksheetFunction.Max(Range("J2:J91"))
    Cells(2, 17).Value = Format(Cells(2, 17), "0.00") + "%"
    Cells(2, 16).Value = Evaluate("Index(I2:I91,Match(Max(J2:J91), J2:J91,0))")
    
' Greatest % decrease - Find minimum value in yearly change with associated ticker name
    Cells(3, 17).Value = WorksheetFunction.Min(Range("J2:J91"))
    Cells(3, 17).Value = Format(Cells(3, 17), "0.00") + "%"
    Cells(3, 16).Value = Evaluate("Index(I2:I91,Match(Min(J2:J91), J2:J91,0))")
    
' Greatest total volume - Find maximum value in total stock volume with associated ticker name
    Cells(4, 17).Value = WorksheetFunction.Max(Range("L2:L91"))
    Cells(4, 17).Value = Round(Cells(4, 17), 2)
    Cells(4, 16).Value = Evaluate("Index(I2:I91,Match(Max(L2:L91), L2:L91,0))")


Next k


End Sub

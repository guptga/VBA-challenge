Sub alphabetical_testing()

Dim Ticker As String
Dim MinDate As Date
Dim MaxDate As Date
Dim OpenValue As Double
Dim CloseValue As Double
Dim TotalVolume As Double
Dim OutputRowCount As Long
Dim RowCount As Long
Dim Greatestincreaseticker As String
Dim Greatestdecreaseticker As String
Dim Greatestincreasevalue As Double
Dim Greatestdecreasevalue As Double
Dim Greatesttotalvolumeticker As String
Dim Greatesttotalvolumevalue As Double


'Activate sheet
For Each ws In ActiveWorkbook.Sheets
'Sheets("B").Activate
    ws.Activate

    'Fix Date format
    Dim c As Range
    Application.ScreenUpdating = False
    For Each c In Range("B2:B" & Cells(Rows.Count, "B").End(xlUp).Row)
        If c.NumberFormat <> "mm/dd/yyyy" Then
            c.Value = DateSerial(Left(c.Value, 4), Mid(c.Value, 5, 2), Right(c.Value, 2))
            c.NumberFormat = "mm/dd/yyyy"
        End If
          
    Next

    'Add Headers

    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Stock Volume"
    Cells(1, 17).Value = "Ticker"
    Cells(1, 18).Value = "Value"
    Cells(2, 16).Value = "Greatest % increase"
    Cells(3, 16).Value = "Greatest % decrease"
    Cells(4, 16).Value = "Greatest total volume"

    'Initialize
    OutputRowCount = 2
    RowCount = 3

    Ticker = Cells(2, 1).Value
    MinDate = Cells(2, 2).Value
    MaxDate = Cells(2, 2).Value
    OpenValue = Cells(2, 3).Value
    CloseValue = Cells(2, 6).Value
    TotalVolume = Cells(2, 7).Value

    'Loop
    Do While Len(Range("A" & CStr(RowCount)).Value) > 0

    If Cells(RowCount, 1) = Ticker Then

        If Cells(RowCount, 2).Value > MaxDate Then
            MaxDate = Cells(RowCount, 2).Value
            CloseValue = Cells(RowCount, 6).Value
            TotalVolume = TotalVolume + Cells(RowCount, 7).Value
        ElseIf Cells(RowCount, 2).Value < MinDate Then
            MinDate = Cells(RowCount, 2).Value
            OpenValue = Cells(RowCount, 3).Value
            TotalVolume = TotalVolume + Cells(RowCount, 7).Value
        Else
            TotalVolume = TotalVolume + Cells(RowCount, 7).Value
        End If
      
    Else
        Cells(OutputRowCount, 10).Value = Ticker
        Cells(OutputRowCount, 11).Value = CloseValue - OpenValue
        Cells(OutputRowCount, 12).Value = (CloseValue - OpenValue) / OpenValue
        Cells(OutputRowCount, 12).NumberFormat = "0.00%"
        Cells(OutputRowCount, 13).Value = TotalVolume
        
        Cells(OutputRowCount, 11).FormatConditions.Delete
        Cells(OutputRowCount, 11).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
        Cells(OutputRowCount, 11).FormatConditions(1).Interior.Color = vbRed
        Cells(OutputRowCount, 11).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
        Cells(OutputRowCount, 11).FormatConditions(2).Interior.Color = vbGreen
        
        OutputRowCount = OutputRowCount + 1
        
        Ticker = Cells(RowCount, 1).Value
        MinDate = Cells(RowCount, 2).Value
        MaxDate = Cells(RowCount, 2).Value
        OpenValue = Cells(RowCount, 3).Value
        CloseValue = Cells(RowCount, 6).Value
        TotalVolume = Cells(RowCount, 7).Value
        
    End If

    RowCount = RowCount + 1

    Loop

    Cells(OutputRowCount, 10).Value = Ticker
    Cells(OutputRowCount, 11).Value = CloseValue - OpenValue
    Cells(OutputRowCount, 12).Value = (CloseValue - OpenValue) / OpenValue
    Cells(OutputRowCount, 12).NumberFormat = "0.00%"
    Cells(OutputRowCount, 13).Value = TotalVolume
    
    Cells(OutputRowCount, 11).FormatConditions.Delete
    Cells(OutputRowCount, 11).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
    Formula1:="=0"
    Cells(OutputRowCount, 11).FormatConditions(1).Interior.Color = vbRed
    Cells(OutputRowCount, 11).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
    Formula1:="=0"
    Cells(OutputRowCount, 11).FormatConditions(2).Interior.Color = vbGreen

'Bonus question



'initialize

Greatestincreaseticker = Cells(2, 10).Value
Greatestdecreaseticker = Cells(2, 10).Value
Greatestincreasevalue = Cells(2, 12).Value
Greatestdecreasevalue = Cells(2, 12).Value
Greatesttotalvolumeticker = Cells(2, 10).Value
Greatesttotalvolumevalue = Cells(2, 13).Value

'forloop

For x = 3 To OutputRowCount
If Cells(x, 12).Value > Greatestincreasevalue Then
    Greatestincreasevalue = Cells(x, 12).Value
    Greatestincreaseticker = Cells(x, 10).Value
End If
    
    
If Cells(x, 12).Value < Greatestdecreasevalue Then
    Greatestdecreasevalue = Cells(x, 12).Value
    Greatestdecreaseticker = Cells(x, 10).Value
End If
    
If Cells(x, 13).Value > Greatesttotalvolumevalue Then
    Greatesttotalvolumevalue = Cells(x, 13).Value
    Greatesttotalvolumeticker = Cells(x, 10).Value
End If

Next

'enter final result

Cells(2, 17).Value = Greatestincreaseticker
Cells(3, 17).Value = Greatestdecreaseticker
Cells(2, 18).Value = Greatestincreasevalue
Cells(2, 18).NumberFormat = "0.00%"
Cells(3, 18).Value = Greatestdecreasevalue
Cells(3, 18).NumberFormat = "0.00%"
Cells(4, 17).Value = Greatesttotalvolumeticker
Cells(4, 18).Value = Greatesttotalvolumevalue


    
Next ws

End Sub




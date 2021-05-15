Attribute VB_Name = "Module1"
'Test Code
Sub Homework2()

    ' Define
    Dim Ticker As String
    Dim Vol As Integer
    Dim Number_Ticker As Integer
    Dim Year_Open As Double
    Dim Year_Close As Double
    Dim Annual_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Dim LastRowState As Long
    Dim Summary_Table_Row As Integer
    
    'Loop All Sheets
    For Each ws In Worksheets
        ws.Activate
        
    LastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Set Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Annual Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' Set Variables
    Number_Ticker = 0
    Ticker = ""
    Annual_Change = 0
    Year_Open = 0
    Percent_Change = 0
    Total_Stock_Volume = 0
    
    'Loop All Rows
    For i = 2 To LastRowState
    
    Ticker = Cells(i, 1).Value
        
        ' Set Open Price
        If Year_Open = 0 Then
            Year_Open = Cells(i, 3).Value
        End If
        
        ' Add Total Stock Volume
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
        ' Different Ticker
        If Cells(i + 1, 1).Value <> Ticker Then
        
            Number_Ticker = Number_Ticker + 1
            Cells(Number_Ticker + 1, 9).Value = Ticker
            Year_Close = Cells(i, 6).Value
            
            ' Annual Change
            Annual_Change = Year_Close - Year_Open
            Cells(Number_Ticker + 1, 10).Value = Annual_Change
            
            'Colors
            If Annual_Change > 0 Then
                ws.Cells(Number_Ticker + 1, 10).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                ws.Cells(Number_Ticker + 1, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(Number_Ticker + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            ' Percent change
            If Year_Open = 0 Then
                Percent_Change = 0
            Else
                Percent_Change = (Annual_Change / Year_Open)
            End If
            
            Cells(Number_Ticker + 1, 11).Value = Format(Percent_Change, "Percent")
            
            If Percent_Change > 0 Then
                ws.Cells(Number_Ticker + 1, 11).Interior.ColorIndex = 4
            ElseIf Percent_Change < 0 Then
                ws.Cells(Number_Tickers + 1, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(Number_Ticker + 1, 11).Interior.ColorIndex = 6
            End If
            
            ' Open Price
            Year_Open = 0
            
            ' Total Stock Volume
            Cells(Number_Ticker + 1, 12).Value = Total_Stock_Volume
            
            ' Set total stock volume back to 0 when we get to a different ticker in the list.
            Total_Stock_Volume = 0
        End If
        
    Next i
    Next
    

End Sub

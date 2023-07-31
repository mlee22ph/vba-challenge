Attribute VB_Name = "Module1"
Sub Main():
    Dim LastRow As Long
    Dim i As Long
    Dim OuputIndex As Integer
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As LongLong

    

    ''Loop thru sheets
    For Each ws In Worksheets
    
        ''-------------------------------------
        ''Display Ticker Summary
        ''-------------------------------------
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ''Initiate columns
        OuputIndex = 2
        TotalVolume = 0
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ''Loop thru all data to create output summary table
        For i = 2 To LastRow
        
            '' Save open price at start for 1st Ticker
            If i = 2 Then
                OpenPrice = ws.Cells(i, "C").Value
            End If
            
            ''Add volume
            TotalVolume = TotalVolume + ws.Cells(i, "G").Value
            
            If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then
                ''Ticker has changed
                ''Save close price and compute yearly change
                ClosePrice = ws.Cells(i, "F").Value
                YearlyChange = ClosePrice - OpenPrice
                
                ''Compute PercentChange
                PercentChange = YearlyChange / OpenPrice
                
                ''Update output table
                ws.Cells(OuputIndex, "I").Value = ws.Cells(i, "A").Value
                ws.Cells(OuputIndex, "J").Value = YearlyChange
                ws.Cells(OuputIndex, "K").Value = PercentChange
                ws.Cells(OuputIndex, "L").Value = TotalVolume
                
                
                ''Apply conditional formatting
                If YearlyChange < 0 Then
                    ''Set negative value to red
                    ws.Cells(OuputIndex, "J").Interior.ColorIndex = 3
                Else
                    ''Set positive value to green
                    ws.Cells(OuputIndex, "J").Interior.ColorIndex = 4
                End If
                
                ''Apply style formatting
                ws.Cells(OuputIndex, "J").NumberFormat = "0.00"
                ws.Cells(OuputIndex, "K").NumberFormat = "0.00%"
                
                '' Re-Initialise variables
                OuputIndex = OuputIndex + 1
                TotalVolume = 0
                
                ''Save new OpenPrice
                OpenPrice = ws.Cells(i + 1, "C").Value
                
            End If
            
        Next i
    
    
        ''-------------------------------------
        ''Display Percdentage Summary
        ''-------------------------------------
        '' Declare variables
        Dim GreatIncreaseTicker As String
        Dim GreatDecreaseTicker As String
        Dim GreatTotalVolTicker As String
        Dim GreatTotalVolume As LongLong
        Dim GreatIncrease As Double
        Dim GreatDecrease As Double
        
        LastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        ''Initiate all values to 1st row
        GreatIncreaseTicker = ws.Cells(2, "I").Value
        GreatDecreaseTicker = ws.Cells(2, "I").Value
        GreatTotalVolTicker = ws.Cells(2, "I").Value
        GreatIncrease = ws.Cells(2, "K").Value
        GreatDecrease = ws.Cells(2, "K").Value
        GreatTotalVolume = ws.Cells(2, "L").Value
    
        ''Loop thru summary table to get greatest increase/decrease/volume summary
        For i = 3 To LastRow
        
            ''Check for greatest increase
            If ws.Cells(i, "K").Value > GreatIncrease Then
                GreatIncreaseTicker = ws.Cells(i, "I").Value
                GreatIncrease = ws.Cells(i, "K").Value
            End If
        
            ''Check for greatest decrease
            If ws.Cells(i, "K").Value < GreatDecrease Then
                GreatDecreaseTicker = ws.Cells(i, "I").Value
                GreatDecrease = ws.Cells(i, "K").Value
            End If
        
            ''Check for greatest total volume
            If ws.Cells(i, "L").Value > GreatTotalVolume Then
                GreatTotalVolTicker = ws.Cells(i, "I").Value
                GreatTotalVolume = ws.Cells(i, "L").Value
            End If
        Next i
        
        ws.Cells(1, "P").Value = "Ticker"
        ws.Cells(1, "Q").Value = "Value"
        
        ws.Cells(2, "O").Value = "Greatest % increase"
        ws.Cells(2, "P").Value = GreatIncreaseTicker
        ws.Cells(2, "Q").Value = GreatIncrease
        ''Apply style formatting
        ws.Cells(2, "Q").NumberFormat = "0.00%"
    
        ws.Cells(3, "O").Value = "Greatest % decrease"
        ws.Cells(3, "P").Value = GreatDecreaseTicker
        ws.Cells(3, "Q").Value = GreatDecrease
        ''Apply style formatting
        ws.Cells(3, "Q").NumberFormat = "0.00%"
        
        ws.Cells(4, "O").Value = "Greatest Total Volume"
        ws.Cells(4, "P").Value = GreatTotalVolTicker
        ws.Cells(4, "Q").Value = GreatTotalVolume
    
        ''Adjust columns to view
        ws.Columns("A:Q").AutoFit
    Next ws
    MsgBox ("COMPLETE!")

End Sub






Sub Price()

    Dim lastRowData As Long
    Dim tickerSearch As String
    Dim dateEnd As String
    Dim dateStart As String
    Dim priceOpen As Variant
    Dim priceClose As Variant
    Dim priceOpenLocated As Boolean
    Dim priceCloseLocated As Boolean
    Dim Ticker As String
    Dim percentageChange As Double
    Dim yearlyChange As Double
    Dim tickerCount As Integer
    Dim encontrado As Boolean
    Dim totalStockVolume As Variant
    Dim maxPercentageChange As Variant
    Dim maxTicker As String
    Dim minPercentageChange As Variant
    Dim minTicker As String
    Dim maxVolume As Variant
    Dim maxVolumeTicker As String
    
    'Timer
    Debug.Print ("Time_start:" & Time)
    
    For Each ws In Worksheets
        'initialize sheet
        lastRowData = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        ' Add columns for the Ticker
        With ws.Range("I1:L1")
            .Value = Array("Ticker", "Yearly Change", "Percentage change", "Total Stock Volume")
            .Font.Bold = True
            .Font.Color = vbBlack
        End With
    
        ' Add columns for resume
        With ws.Range("P1:Q1")
            .Value = Array("Ticker", "Value")
            .Font.Bold = True
            .Font.Color = vbBlack
        End With
        
        With ws.Range("O2")
            .Value = Array("Greatest % increase")
            .Font.Bold = True
            .Font.Color = vbBlack
        End With
        
        With ws.Range("O3")
            .Value = Array("Greatest % decrease")
            .Font.Bold = True
            .Font.Color = vbBlack
        End With
         
        With ws.Range("O4")
            .Value = Array("Greatest total volume")
            .Font.Bold = True
            .Font.Color = vbBlack
        End With
        
        'Create dateEnd using Sheet Year
        dateStart = ws.Name + "0102"
        dateEnd = ws.Name + "1231"
        
        tickerCount = 1
        'Add ticker to ticker column
        For i = 2 To lastRowData
            If (ws.Range("A" & i).Value <> ws.Range("A" & i - 1).Value) Then
                'Increment tickercount
                tickerCount = tickerCount + 1
                
                'Write new ticker to list position tickerCount
                tickerSearch = ws.Range("A" & i).Value
                ws.Range("I" & tickerCount).Value = tickerSearch
                
                'Clean values for ticker
                priceClose = ""
                priceOpen = ""
                totalStockVolume = 0
                priceOpenLocated = False
                priceCloseLocated = False
                totalStockVolume = ws.Range("G" & i).Value
        
                'Check dateStart to obtain priceOpen
                If ws.Range("B" & i).Value = dateStart Then
                    priceOpen = ws.Range("C" & i).Value
                    priceOpenLocated = True
                End If
                'Check dateEnd to obtain priceClose
                If ws.Range("B" & i).Value = dateEnd Then
                    priceClose = ws.Range("F" & i).Value
                    priceCloseLocated = True
                End If
        
            Else
                'Check dateStart to obtain priceOpen
                If ws.Range("B" & i).Value = dateStart Then
                    priceOpen = ws.Range("C" & i).Value
                    priceOpenLocated = True
                End If
                
                'Check dateEnd to obtain priceClose
                If ws.Range("B" & i).Value = dateEnd Then
                    priceClose = ws.Range("F" & i).Value
                    priceCloseLocated = True
                End If
                
                'Add Volume
                totalStockVolume = totalStockVolume + ws.Range("G" & i).Value
                
                'Calc yearlyChange and Percentage Change
                If priceOpenLocated And priceCloseLocated Then
                    yearlyChange = priceClose - priceOpen
                    
                    'write yearly change
                    ws.Range("J" & tickerCount).Value = yearlyChange
                    ws.Range("J" & tickerCount).NumberFormat = "0.00"

                    'add color format to Yearly Change
                    If (yearlyChange >= 0) Then
                        ws.Range("J" & tickerCount).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & tickerCount).Interior.ColorIndex = 3
                    End If
                    
                    If (priceOpen <> 0) Then
                        percentageChange = (priceClose - priceOpen) / priceOpen
                    Else
                        percentageChange = 0
                    End If
                    
                    'write percentage change
                    ws.Range("K" & tickerCount).Value = percentageChange
                    ws.Range("K" & tickerCount).NumberFormat = "0.00%"
                    
                    'add color format to percentage Change
                    If (percentageChange >= 0) Then
                        ws.Range("K" & tickerCount).Interior.ColorIndex = 4
                    Else
                        ws.Range("K" & tickerCount).Interior.ColorIndex = 3
                    End If
                End If
                                
                'Write TotalStockVolume
                ws.Range("L" & tickerCount).Value = totalStockVolume
                ws.Range("L" & tickerCount).NumberFormat = "#,##0"
            
            End If
        Next i
         
        maxPercentageChange = ws.Range("K2").Value
        maxTicker = ws.Range("I" & i).Value
        minPercentageChange = ws.Range("K2").Value
        minTicker = ws.Range("I" & i).Value
        maxVolume = ws.Range("L2").Value
        maxVolumeTicker = ws.Range("I" & i).Value
            
        For i = 3 To tickerCount - 1
            'max percentage change
            If ws.Range("K" & i).Value > maxPercentageChange Then
                maxPercentageChange = ws.Range("K" & i).Value
                maxTicker = ws.Range("I" & i).Value
            End If
            'min percentage change
            If ws.Range("K" & i).Value < minPercentageChange Then
                minPercentageChange = ws.Range("K" & i).Value
                minTicker = ws.Range("I" & i).Value
            End If
            'max volume
            If ws.Range("L" & i).Value > maxVolume Then
                maxVolume = ws.Range("L" & i).Value
                maxVolumeTicker = ws.Range("I" & i).Value
            End If
        Next i
        'Write Greastest % increase
        ws.Range("P2").Value = maxTicker
        ws.Range("Q2").Value = maxPercentageChange
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P3").Value = minTicker
        ws.Range("Q3").Value = minPercentageChange
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P4").Value = maxVolumeTicker
        ws.Range("Q4").Value = maxVolume
        ws.Range("Q4").NumberFormat = "#,##0"
        
        'Adjust columns width
        ws.Columns("I:Q").WrapText = False
        ws.Columns("I:Q").AutoFit
        
        
    Next ws
    
    'Timer
    Debug.Print ("Time_end:" & Time)
        
End Sub

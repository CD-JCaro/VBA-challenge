Sub Process():

    'variable declaration
    Dim strCurrTick As String
    Dim dCurrVolume As Double
    Dim dCurrStart As Double
    Dim lCurrRow As Long
    Dim lCurrPrintRow As Long
    Dim dCurrEnd As Double
    Dim dMaxPercent As Double
    Dim dMinPercent As Double
    Dim dMaxVolume As Double
    Dim nNumSheets As Integer
    Dim dChange As Double
    
    nNumSheets = ActiveWorkbook.Worksheets.Count

    'looping through our sheets
    For i = 1 To nNumSheets

        're-init our variables for our new sheet
        lCurrRow = 2
        lCurrPrintRow = 2
        dCurrVolume = 0
        dMaxVolume = 0
        dMaxPercent = 0
        dMinPercent = 0

        'saving our first ticker symbol and saving its start value
        strCurrTick = ActiveWorkbook.Worksheets(i).Cells(2, "A").Value
        dCurrStart = ActiveWorkbook.Worksheets(i).Cells(2, "C").Value

        'adding our labels for the table
        ActiveWorkbook.Worksheets(i).Cells(1, "I").Value = "Ticker"
        ActiveWorkbook.Worksheets(i).Cells(1, "J").Value = "Yearly Change"
        ActiveWorkbook.Worksheets(i).Cells(1, "K").Value = "Percent Change"
        ActiveWorkbook.Worksheets(i).Cells(1, "L").Value = "Total Stock Volume"

        'loop through our rows until we  find one with no ticker symbol
        For lCurrRow = 2 To ActiveWorkbook.Worksheets(i).Cells(Rows.Count, 1).End(xlUp).Row

            ' if we dont have a start value already... set our start value
            If dCurrStart = 0 Then
                dCurrStart = ActiveWorkbook.Worksheets(i).Cells(lCurrRow, "C").Value
            End If
            
            'adding up our total volume for current ticker as well as keeping track of the current end
            'as we wont know if this is the last data set for this symbol until our next loop
            dCurrVolume = dCurrVolume + ActiveWorkbook.Worksheets(i).Cells(lCurrRow, "G").Value
                    
            'if our new ticker symbol does not match our next ticker symbol then we need to do some work
            If ActiveWorkbook.Worksheets(i).Cells(lCurrRow, "A").Value <> ActiveWorkbook.Worksheets(i).Cells(lCurrRow + 1, "A").Value Then
                
                dCurrEnd = ActiveWorkbook.Worksheets(i).Cells(lCurrRow, "F").Value
                'time to calculate and print stats
                
                ActiveWorkbook.Worksheets(i).Cells(lCurrPrintRow, "I").Value = strCurrTick
                dChange = dCurrEnd - dCurrStart
    
                'printing change and color coding it based on +-
                ActiveWorkbook.Worksheets(i).Cells(lCurrPrintRow, "J").Value = dChange
                
                If dChange > 0 Then
                    ActiveWorkbook.Worksheets(i).Cells(lCurrPrintRow, "J").Interior.ColorIndex = 10
                ElseIf dChange < 0 Then
                    ActiveWorkbook.Worksheets(i).Cells(lCurrPrintRow, "J").Interior.ColorIndex = 3
                End If
        
                If dCurrStart = 0 Then
                    ActiveWorkbook.Worksheets(i).Cells(lCurrPrintRow, "K").Value = 0
                Else
                    ActiveWorkbook.Worksheets(i).Cells(lCurrPrintRow, "K").Value = dChange / dCurrStart
                    
                    'we see if the change is > our current highest % gainer or < our current highest % loser
                    If dChange / dCurrStart > dMaxPercent Then
                        ActiveWorkbook.Worksheets(i).Cells(2, "Q").Value = dChange / dCurrStart
                        ActiveWorkbook.Worksheets(i).Cells(2, "Q").NumberFormat = "0.00%"
                        ActiveWorkbook.Worksheets(i).Cells(2, "P").Value = strCurrTick
                        dMaxPercent = dChange / dCurrStart
        
                    'less than min?
                    ElseIf dChange / dCurrStart < dMinPercent Then
                        ActiveWorkbook.Worksheets(i).Cells(3, "Q").Value = dChange / dCurrStart
                        ActiveWorkbook.Worksheets(i).Cells(3, "Q").NumberFormat = "0.00%"
                        ActiveWorkbook.Worksheets(i).Cells(3, "P").Value = strCurrTick
                        dMinPercent = dChange / dCurrStart
                    End If
        
            
                   
                End If
                
                'formatting our %change column to represent percentages
                ActiveWorkbook.Worksheets(i).Cells(lCurrPrintRow, "K").NumberFormat = "0.00%"
                
                'print out our total volume for our old ticker symbol
                ActiveWorkbook.Worksheets(i).Cells(lCurrPrintRow, "L").Value = dCurrVolume
        
                'checking if our current symbol takes the lead for highest volume
                If dCurrVolume > dMaxVolume Then
                    dMaxVolume = dCurrVolume
                    ActiveWorkbook.Worksheets(i).Cells(4, "Q").Value = dCurrVolume
                    ActiveWorkbook.Worksheets(i).Cells(4, "P").Value = strCurrTick
                End If
                
                'once printed we need to reset values for the next ticker symbol
                dCurrStart = 0
                lCurrPrintRow = lCurrPrintRow + 1
                dCurrVolume = 0
                strCurrTick = ActiveWorkbook.Worksheets(i).Cells(lCurrRow + 1, "A").Value
                
            End If
            
        Next lCurrRow 'off to the next row

        'printing out our headers for the new table we created
        ActiveWorkbook.Worksheets(i).Cells(1, "P").Value = "Ticker"
        ActiveWorkbook.Worksheets(i).Cells(1, "Q").Value = "Value"

        ActiveWorkbook.Worksheets(i).Cells(2, "O").Value = "Greatest % Increase"
        ActiveWorkbook.Worksheets(i).Cells(3, "O").Value = "Greatest % Decrease"
        ActiveWorkbook.Worksheets(i).Cells(4, "O").Value = "Greatest Total Volume"

        'make sure everything is nice and cozy in their new cells
        ActiveWorkbook.Worksheets(i).Range("I:Q").Columns.AutoFit

    Next i 'time to go to next worksheet
    
End Sub

Sub clear():
    
    Dim nNumSheets As Integer
    
    nNumSheets = ActiveWorkbook.Worksheets.Count

    For i = 1 To nNumSheets
        ActiveWorkbook.Worksheets(i).Range("I:Q").clear
    Next i
        
End Sub


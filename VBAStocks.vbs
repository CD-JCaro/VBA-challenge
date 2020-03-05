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
        While ActiveWorkbook.Worksheets(i).Cells(lCurrRow, "A").Value <> ""

            'if our new ticker symbol does not match our current ticker symbol then we need to do some work
            If ActiveWorkbook.Worksheets(i).Cells(lCurrRow, "A").Value <> strCurrTick Then
                
                'time to calculate and print stats
                Call PrintValue(i, lCurrPrintRow, strCurrTick, dCurrEnd, dCurrStart, dCurrVolume, dMaxPercent, dMinPercent, dMaxVolume)

                'once printed we need to reset values for the next ticker symbol
                dCurrStart = 0
                lCurrPrintRow = lCurrPrintRow + 1
                dCurrVolume = 0
                strCurrTick = ActiveWorkbook.Worksheets(i).Cells(lCurrRow, "A").Value
                
            End If

            'adding up our total volume for current ticker as well as keeping track of the current end
            'as we wont know if this is the last data set for this symbol until our next loop
            dCurrVolume = dCurrVolume + ActiveWorkbook.Worksheets(i).Cells(lCurrRow, "G").Value

            ' if we dont have a start value already... set our start value
            If dCurrStart = 0 Then
                dCurrStart = ActiveWorkbook.Worksheets(i).Cells(lCurrRow, "C").Value
            End If

            dCurrEnd = ActiveWorkbook.Worksheets(i).Cells(lCurrRow, "F").Value
            lCurrRow = lCurrRow + 1
        Wend 'off to the next row

        'well... i forgot to print the last guy... oops
        'super lazy copy pasta

        'TODO: Make this crap a function
        Call PrintValue(i, lCurrPrintRow, strCurrTick, dCurrEnd, dCurrStart, dCurrVolume, dMaxPercent, dMinPercent, dMaxVolume)
        

        'END TODO

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

Sub PrintValue(ByVal i As Integer, ByVal lPrintrow As Long, ByVal strTick As String, ByVal currEnd As Double, ByVal currStart As Double, ByVal currVolume As Double, ByRef maxPercent As Double, ByRef minPercent As Double, ByRef maxVolume As Double)

    Dim dChange As Double

    ActiveWorkbook.Worksheets(i).Cells(lPrintrow, "I").Value = strTick
        dChange = currEnd - currStart

        'printing change and color coding it based on +-
        ActiveWorkbook.Worksheets(i).Cells(lPrintrow, "J").Value = dChange
        If dChange > 0 Then
            ActiveWorkbook.Worksheets(i).Cells(lPrintrow, "J").Interior.ColorIndex = 10
        ElseIf dChange < 0 Then
            ActiveWorkbook.Worksheets(i).Cells(lPrintrow, "J").Interior.ColorIndex = 3
        End If

        If dChange = 0 Then
            ActiveWorkbook.Worksheets(i).Cells(lPrintrow, "K").Value = 0
        Else
            ActiveWorkbook.Worksheets(i).Cells(lPrintrow, "K").Value = dChange / currStart
            
            'we see if the change is > our current highest % gainer or < our current highest % loser
            If dChange / currStart > maxPercent Then
                ActiveWorkbook.Worksheets(i).Cells(2, "Q").Value = dChange / currStart
                ActiveWorkbook.Worksheets(i).Cells(2, "Q").NumberFormat = "0.00%"
                ActiveWorkbook.Worksheets(i).Cells(2, "P").Value = strTick
                maxPercent = dChange / currStart

            'less than min?
            ElseIf dChange / currStart < minPercent Then
                ActiveWorkbook.Worksheets(i).Cells(3, "Q").Value = dChange / currStart
                ActiveWorkbook.Worksheets(i).Cells(3, "Q").NumberFormat = "0.00%"
                ActiveWorkbook.Worksheets(i).Cells(3, "P").Value = strTick
                minPercent = dChange / currStart
            End If
        End If

        'formatting our %change column to represent percentages
        ActiveWorkbook.Worksheets(i).Cells(lPrintrow, "K").NumberFormat = "0.00%"

        'print out our total volume for our old ticker symbol
        ActiveWorkbook.Worksheets(i).Cells(lPrintrow, "L").Value = currVolume

        'checking if our current symbol takes the lead for highest volume
        If currVolume > maxVolume Then
            maxVolume = currVolume
            ActiveWorkbook.Worksheets(i).Cells(4, "Q").Value = currVolume
            ActiveWorkbook.Worksheets(i).Cells(4, "P").Value = strTick
        End If

End Sub

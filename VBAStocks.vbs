Sub Process():
    'variable declaration
    Dim strCurrTick As String
    Dim dCurrVolume As Double
    Dim dCurrStart As Double
    Dim lCurrRow As Long
    dim lCurrPrintRow as Long
    dim dCurrEnd as double
    dim dChange as double
    dim dMaxPercent as double
    dim dMinPercent as Double
    dim dMaxVolume as double
    dim nNumSheets as Integer
    
    nNumSheets = activeworkbook.worksheets.count

    'looping through our sheets
    for i = 1 to nNumSheets
        're-init our variables for our new sheet
        lCurrRow = 2
        lCurrPrintRow = 2
        dCurrVolume = 0
        dMaxVolume = 0
        dMaxPercent = 0
        dMinPercent = 0

        'saving our first ticker symbol and saving its start value
        strCurrTick = activeworkbook.worksheets(i).Cells(2, "A").Value
        dCurrStart = activeworkbook.worksheets(i).cells(2, "C").Value
    
        'loop through our rows until we  find one with no ticker symbol
        While activeworkbook.worksheets(i).Cells(lCurrRow, "A").Value <> ""

            'if our new ticker symbol does not match our current ticker symbol then we need to do some work
            if activeworkbook.worksheets(i).Cells(lCurrRow,"A").Value <> strCurrTick then
                
                'time to calculate and print stats
                
                'printing out our old ticker symbol to our results table
                activeworkbook.worksheets(i).cells(lCurrPrintRow, "I").value = strCurrTick
                dChange = dCurrEnd - dCurrStart

                'printing change and color coding it based on +-
                activeworkbook.worksheets(i).cells(lCurrPrintRow, "J").Value = dChange
                if dchange > 0   then
                    activeworkbook.worksheets(i).cells(lCurrPrintRow, "J").Interior.ColorIndex = 10
                elseif dchange < 0 then
                    activeworkbook.worksheets(i).cells(lCurrPrintRow, "J").Interior.ColorIndex = 3
                end if
                if dChange = 0 then
                    activeworkbook.worksheets(i).cells(lCurrPrintRow, "K").value = 0
                else    
                    activeworkbook.worksheets(i).cells(lCurrPrintRow, "K").value = dChange / dCurrStart
                    
                    'we see if the change is > our current highest % gainer or < our current highest % loser
                    if dchange/dCurrStart > dMaxPercent then
                        activeworkbook.worksheets(i).cells(2, "Q").value = dChange/dCurrStart
                        activeworkbook.worksheets(i).cells(2, "Q").NumberFormat = "0.00%"
                        activeworkbook.worksheets(i).cells(2, "P").value = strCurrTick
                        dMaxPercent = dchange/dCurrStart

                    elseif dchange/dCurrStart < dMinPercent then
                        activeworkbook.worksheets(i).cells(3, "Q").value = dChange/dCurrStart
                        activeworkbook.worksheets(i).cells(3, "Q").NumberFormat = "0.00%"
                        activeworkbook.worksheets(i).cells(3, "P").value = strCurrTick
                        dMinPercent = dchange/dCurrStart
                    end if
                end if
                'formatting our %change column to represent percentages
                activeworkbook.worksheets(i).cells(lCurrPrintRow, "K").NumberFormat = "0.00%"
                'print out our total volume for our old ticker symbol
                activeworkbook.worksheets(i).cells(lCurrPrintRow, "L").value = dCurrVolume

                'checking if our current symbol takes the lead for highest volume
                if dCurrVolume > dMaxVolume then
                    dMaxVolume = dCurrVolume
                    activeworkbook.worksheets(i).cells(4, "Q").Value = dCurrVolume
                    activeworkbook.worksheets(i).cells(4, "P").Value = strCurrTick
                end if

                'once printed we need to reset values for the next ticker symbol
                dCurrStart = 0
                lCurrPrintRow = lCurrPrintRow + 1
                dCurrVolume = 0
                strCurrTick = activeworkbook.worksheets(i).cells(lCurrRow,"A").value
                
            end if
            'adding up our total volume for current ticker as well as keeping track of the current end
            'as we wont know if this is the last data set for this symbol until our next loop
            dCurrVolume = dCurrVolume + activeworkbook.worksheets(i).cells(lCurrRow,"G").value
            if dCurrStart = 0 then
                dCurrStart = activeworkbook.worksheets(i).cells(lCurrRow,"C").value
            end if

            dCurrEnd = activeworkbook.worksheets(i).cells(lCurrRow,"F").value
            lCurrRow = lCurrRow + 1
        Wend

        'printing out our headers for the new table we created
        activeworkbook.worksheets(i).cells(1, "P").Value = "Ticker"
        activeworkbook.worksheets(i).cells(1, "Q").Value = "Value"

        activeworkbook.worksheets(i).cells(2, "O").Value = "Greatest % Increase"
        activeworkbook.worksheets(i).cells(3, "O").Value = "Greatest % Decrease"
        activeworkbook.worksheets(i).cells(4, "O").Value = "Greatest Total Volume"

        'make sure everything is nice and cozy in their new cells
        activeworkbook.worksheets(i).range("I:Q").columns.autofit
    next i 'time to go to next worksheet
    
End Sub

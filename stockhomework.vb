Sub stocks()
    ' this sub is the main entry please execute me
    For Each ws In Worksheets
        Dim lastRowTicker As String
        Dim lastRowTickerRowIndex As Long

        ' defining cols where we are going to extract the data
        Dim VolCol As Integer
        VolCol = 7
        Dim HighCol As Integer
        HighCol = 4
        Dim LowCol As Integer
        LowCol = 5
        Dim CloseCol As Integer
        CloseCol = 6
        Dim StartCol As Integer
        StartCol = 3

        ' here we are going to add ticker data
        Dim tickerGroupRowNumber As Integer
        Dim tickerGroupColNumber As Integer
        tickerGroupColNumber = 9 ' we are going to set info in column I
        tickerGroupRowNumber = 2 ' we're going to start from here + 1 in order to skip header

        totalNumberOfRows = ws.Cells(Rows.Count, 1).End(xlUp).row() ' total number of rows with data in the sheet
        lastRowTickerRowIndex = 2 'starting from rwo number 2 so we can skip header

        '''' Value acumulators
            Dim acumClosePrice As Double
            Dim acumStartPrice As Double
            Dim acumVol As Variant
            Dim acumHigh As Double
            Dim acumLow As Double
            Dim acumDays As Integer

        initCols ws,1,tickerGroupColNumber, 1, 16
        For currentRow = lastRowTickerRowIndex To totalNumberOfRows
            currentRowTickerValue = ws.Cells(lastRowTickerRowIndex, 1).Value
            ' if current row is different from last row or we have reached last row
            ' then let's just print data and go next sheet or row (whatever it's true)
            If ((currentRowTickerValue <> lastRowTicker And lastRowTicker <> "") Or lastRowTickerRowIndex = totalNumberOfRows) Then

                processAcumValues ws, tickerGroupRowNumber, tickerGroupColNumber, lastRowTicker, acumClosePrice, acumStartPrice, acumVol, acumHigh, acumLow,acumDays
                tickerGroupRowNumber = tickerGroupRowNumber + 1 ' increase offset so we can write in the next row

                ' once data is written then we can reset our acums
                acumClosePrice = 0
                acumStartPrice = 0
                acumVol = 0
                acumHigh = 0
                acumLow = 0
                acumDays=0

            Else
                ' we can just acumulate values for this group, later we'll process it before writing it
                acumClosePrice = acumClosePrice + ws.Cells(lastRowTickerRowIndex, CloseCol).Value
                acumStartPrice = acumStartPrice + ws.Cells(lastRowTickerRowIndex, StartCol).Value
                acumVol = acumVol + ws.Cells(lastRowTickerRowIndex, VolCol).Value
                acumHigh = acumHigh + ws.Cells(lastRowTickerRowIndex, HighCol).Value
                acumLow = acumLow + ws.Cells(lastRowTickerRowIndex, LowCol).Value
                acumDays=acumDays+1
            End If
            lastRowTicker = currentRowTickerValue ' current ticker value is now saved
            lastRowTickerRowIndex = 1 + lastRowTickerRowIndex  ' keep row counter increasing by one
        Next currentRow
        writePctInc ws, 2, 15,9,11,12
    Next ws
    
    
End Sub

Sub initCols(ws, rowInformation1, colInformation1, rowInformation2, colInformation2)
    ws.Cells(rowInformation1, colInformation1).Value = "Ticker"
    ws.Cells(rowInformation1, colInformation1+1).Value = "Yearly Change"
    ws.Cells(rowInformation1, colInformation1+2).Value = "Percentage Change"
    ws.Cells(rowInformation1, colInformation1+3).Value = "Total Stock Volume"

    ws.Cells(rowInformation2, colInformation2).Value = "Ticker"
    ws.Cells(rowInformation2, colInformation2+1).Value = "Value"

End Sub

Sub writePctInc(ws, rowWhereToWrite, colWhereToWrite, tickerTextCol, percentChangeCol, totalStockVolumeCol)

        Dim rowGreatestPctInc AS Integer
        Dim rowGreatestPctDec AS Integer
        Dim rowGreatestVol AS Integer
        Dim lastRowWithResults AS Integer
        Dim lastPctRow AS Integer
        Dim OperationNameCol AS Integer
        Dim tickerNameCol AS Integer
        Dim valueCol AS Integer
        
        ' extract volume range
        lastRowWithResults = ws.Cells(Rows.Count, totalStockVolumeCol).End(xlUp).row() -1' total number of rows with data in the sheet
        set volumeRange = Range(Cells(2, totalStockVolumeCol), Cells(lastRowWithResults, totalStockVolumeCol))

        ' extract pct range 
        set pctRange = Range(Cells(2, percentChangeCol), Cells(lastRowWithResults, percentChangeCol))
 
        ' adjust cols where to write
        OperationNameCol = colWhereToWrite
        tickerNameCol = colWhereToWrite+1
        valueCol = colWhereToWrite + 2

        ' greatest % increase
        maxIncreaseValue=WorksheetFunction.Max(pctRange)
        maxIncreaseRow=WorksheetFunction.Match(maxIncreaseValue,pctRange)
        ' these lines will print titles and values
        rowGreatestPctInc=rowWhereToWrite
        ws.Cells(rowGreatestPctInc, OperationNameCol).Value = "Greatest % increase"
        ws.Cells(rowGreatestPctInc, tickerNameCol).Value = ws.Cells(maxIncreaseRow,tickerTextCol).Value
        ws.Cells(rowGreatestPctInc, valueCol).Value = maxIncreaseValue

        ' greatest  % decrease 
        ' not able to find the
       ' minDecValue=WorksheetFunction.Min(pctRange)
       ' minDecRow=WorksheetFunction.Match(minDecValue,pctRange)
        ' these lines will print titles and values
       ' rowGreatestPctDec=rowWhereToWrite+1
       ' ws.Cells(rowGreatestPctDec, OperationNameCol).Value = "Greatest % decrease"
       ' ws.Cells(rowGreatestPctDec, tickerNameCol).Value = ws.Cells(minDecRow,tickerTextCol).Value
       ' ws.Cells(rowGreatestPctDec, valueCol).Value = minDecValue


        ' greatest total volume
        maxVolumeValue=WorksheetFunction.Max(volumeRange)
        maxVolumeRow=WorksheetFunction.Match(maxVolumeValue,volumeRange)
        ' these lines will print titles and values
        rowGreatestVol=rowWhereToWrite+2
        ws.Cells(rowGreatestVol, OperationNameCol).Value =  "Greatest total volume"
        ws.Cells(rowGreatestVol, tickerNameCol).Value = ws.Cells(maxVolumeRow,tickerTextCol).Value
        ws.Cells(rowGreatestVol, valueCol).Value = maxVolumeValue
 



End Sub

Sub processAcumValues(ws, row, col, tickerText, acumClosePrice, acumStartPrice, acumVol, acumHigh, acumLow,acumDays)
    if(acumClosePrice > 0 ) then 
        yearlyChange = (acumClosePrice - acumStartPrice)-1
        opePriceAvg= acumStartPrice/acumDays
        closePriceAvg=acumClosePrice/acumDays
        percentChange = ((closePriceAvg*100)/opePriceAvg) - 100
        'Debug.Print opePriceAvg closePriceAvg acumStartPrice  acumClosePrice  acumDays
        addTickerGroup ws, row, col, tickerText, yearlyChange, percentChange, acumVol
    end if 
End Sub

Sub addTickerGroup(ws, row, col, tickerText, yearlyChange, percentChange, totalStockVolume)
    ws.Cells(row, col).Value = tickerText

     ws.Cells(row, col+1).Value = yearlyChange
    if(yearlyChange > 0 ) then
         ws.Cells(row, col+1).interior.ColorIndex=4 ' green
    Else 
         ws.Cells(row, col+1).interior.ColorIndex=3 ' red
    End If 
    ws.Cells(row, col+2).Value = percentChange
    ws.Cells(row, col+3).Value = totalStockVolume

End Sub


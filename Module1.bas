Attribute VB_Name = "Module1"
Sub Main()
    For Each ws In Worksheets
        Dim ticker As String
        Dim gIncreaseTicker As String       'Seperate tickers for the Greatest % Increase
        Dim gDecreaseTicker As String       'Greatest % Decrease
        Dim gVolumeTicker As String         'And Greatest Total Volume
        Dim yearOpen As Double
        Dim yearClose As Double
        Dim yearChange As Double
        Dim percentChange As Double
        Dim totalVolume As LongLong     'A LongLong is needed, as we are dealing with numbers over 2 billion.
        Dim gTotalVolume As LongLong
        gTotalVolume = 0            'Same as the seperate tickers above
        Dim gIncrease As Double     'These hold their respective values
        gIncrease = 0               'They are initially set to 0
        Dim gDecrease As Double
        gDecrease = 0
        Dim LastRow As Long         'We are dealing with WAY too many rows to use an Integer
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Cells(1, 9).Value = "Ticker"                 'These all need to be added
        ws.Cells(1, 10).Value = "Yearly Change"         'Before we loop
        ws.Columns(10).AutoFit
        ws.Cells(1, 11).Value = "Percent Change"        'And add values
        ws.Columns(11).AutoFit
        ws.Cells(1, 12).Value = "Total Stock Volume"    'Below them
        ws.Cells(1, 16).Value = "Ticker"                'BONUS
        ws.Cells(1, 17).Value = "Value"                 'BONUS
        ws.Cells(2, 15).Value = "Greatest % Increase"   'BONUS
        ws.Cells(3, 15).Value = "Greatest % Decrease"   'BONUS
        ws.Cells(4, 15).Value = "Greatest Total Volume" 'BONUS
        ws.Columns(15).AutoFit
        Dim firstRow As Boolean         'Needed in order to find the yearOpen when we loop below
        firstRow = True                 'We start with the first row of the first stock, so this starts as TRUE
        Dim displayCounter As Integer   'The display row for each stock will easily fit into an Integer
        displayCounter = 2
        Dim rowTick As String       'Initializing final veriables
        Dim dayOpen As Double       'Before looping
        Dim dayClose As Double
        Dim dayVolume As LongLong
        For i = 2 To LastRow    'Loop thru every row ignoring the header row
            ticker = ws.Cells(i, 1).Value           'Read and store the <ticker> for this row
            dayOpen = ws.Cells(i, 3).Value          'Read and store the <open price> for this row
            dayClose = ws.Cells(i, 6).Value         'Read and store the <close price> for this row
            dayVolume = ws.Cells(i, 7).Value        'Read and store the <volume of stock> for this row
            totalVolume = totalVolume + dayVolume   'Add dayVolume to the totalVolume for this stock
            If ws.Cells(i + 1, 1).Value <> ticker Then
                yearClose = dayClose                    'This row is the last day of the year, so the dayClose is the yearClose
                yearChange = yearClose - yearOpen
                percentChange = yearChange / yearOpen
                ws.Cells(displayCounter, 9).Value = ticker
                ws.Cells(displayCounter, 10).Value = yearChange
                ws.Cells(displayCounter, 10).NumberFormat = "###0.00"
                If yearChange < 0 Then
                    ws.Cells(displayCounter, 10).Interior.ColorIndex = 3    'Format RED for negative
                    ws.Cells(displayCounter, 11).Interior.ColorIndex = 3
                Else                                                        'Formatting BOTH YearlyChange and PercentChange for good mesure
                    ws.Cells(displayCounter, 10).Interior.ColorIndex = 4
                    ws.Cells(displayCounter, 11).Interior.ColorIndex = 4    'Format GREEN otherwise
                End If
                If percentChange > gIncrease Then           'If this stock has a higher % change then the current gIncrease value,
                    gIncreaseTicker = rowTick               'Then the gIncreaseTicker is updated with this stock's ticker
                    gIncrease = percentChange               'And gIncrease is updated with this stock's percentChange
                ElseIf percentChange < gDecrease Then
                    gDecreaseTicker = rowTick               'Same as above, but if it has a lower % change then
                    gDecrease = percentChange               'the current gDecrease value
                Else
                    'Nothing happens. Because this stock's percent change is neither the greatest increase nor decrease.
                End If
                If totalVolume > gTotalVolume Then          'Same as above, but checking for greatest total volume
                    gVolumeTicker = rowTick
                    gTotalVolume = totalVolume
                Else
                    'Nothing Happens (see notes above)
                End If
                ws.Cells(displayCounter, 11).Value = percentChange
                ws.Cells(displayCounter, 11).NumberFormat = "##0.00%"
                ws.Cells(displayCounter, 12).Value = totalVolume
                
                displayCounter = displayCounter + 1     'Displays the next stock on the next row
                totalVolume = 0                         'RESETS totalVolume for the next stock
                firstRow = True     'This part is important because it only triggers when the next row's ticker is the DIFFERENT than
                                    'THIS row's ticker.
                                    'This is the last code that runs before the next loop,
                                    'So firstRow needs to be set to TRUE so that it can set the yearOpen for the next stock!
            Else
            'This "Else" runs every time the stock ticker for the NEXT row is the SAME as the stock ticker in THIS row.
            'Given that we have hundreds of thousands of lines of code PER SHEET,
            'This "Else" runs WAYYY more that the "If" above.
            
            'SINCE firstRow is a Boolean,
            'The If statement below CAN look like it does! It needs no equal sign :D
            'The only requirement for an If statement is that the test equates to TRUE or FALSE, which it does!
            'I chose to do it this way because a counter would needlesly slow the code down,
            'By needlessly counting how many loops it went thru a given stock.
            'The only thing that matters is weither or not this is the First Row.
            'If it is, then the dayOpen for this row needs to be saved as the yearOpen
                If firstRow Then
                    yearOpen = dayOpen              'This row is the first day of the year, so the dayOpen is the yearOpen
                Else
                    'Nothing happens. This row is not the first day of the year
                End If
                firstRow = False    'This part is important, because it only triggers when the next row's ticker is the SAME
                                    'As THIS row's ticker, therefore, the next row WILL NOT be the first row for it's stock.
            End If
        Next i                                      '==========BONUS==========
        ws.Cells(2, 16).Value = gIncreaseTicker     'The Ticker for Greatest % Increase
        ws.Cells(2, 17).Value = gIncrease           'And the value
        ws.Cells(2, 17).NumberFormat = "##0.00%"
        ws.Cells(3, 16).Value = gDecreaseTicker     'The Ticker for Greatest % Decrease
        ws.Cells(3, 17).Value = gDecrease           'And the value
        ws.Cells(3, 17).NumberFormat = "##0.00%"
        ws.Cells(4, 16).Value = gVolumeTicker       'The Ticker for Greatest Total Volume
        ws.Cells(4, 17).Value = gTotalVolume        'And the value (note, this is not formatted because it isn't a percent)
        ws.Columns(12).AutoFit      'This autoFits the Total Stock Volume Column . . .
                                    'I included this last, so it isn't always re-formatting . . .
       'In the example, I notice that the picture for hard_solution.png has the Greatest Total Volume Showing as: "1.69E+12"
       'As opposed to the Total Stock Volume column which shows the whole number for each stock's volume . . .
       'I have designed my code to match this, so I am intentionally ommitting ws.Columns(17).AutoFit . . .
    Next ws
End Sub

Attribute VB_Name = "Module1"
Sub getInfo():
    ' tickerList, ArrayList to save the unique stock tickers seen in data for quarter.
    Dim tickerList As ArrayList
    Set tickerList = New ArrayList
    
    ' startDate, ArrayList to save the first date of stock sold during quarter.
    ' endDate, ArrayList to save the last date of stock sold during quarter.
    Dim startDate As ArrayList, endDate As ArrayList
    Set startDate = New ArrayList
    Set endDate = New ArrayList
    
    ' openList, ArrayList to save open price of stock at beginning of quarter.
    ' closeList, ArrayList to save close price of stock at end of quarter.
    Dim openList As ArrayList, closeList As ArrayList
    Set openList = New ArrayList
    Set closeList = New ArrayList
    
    ' volumeList, ArrayList to save volume of stock sold during quarter.
    Dim volumeList As ArrayList
    Set volumeList = New ArrayList
    
    ' index, an Integer representing the index of our stock ticker being looked at.
    '      ... This is useful for not repeating: 'tickerList.indexOf(Cells(cell.Row, 1).value, 0))
    '      ... which we would have to repeat quite a bit... and this makes our code more readable.
    Dim index As Integer
    
    ' Loop thru each non-blank cell in Column A.
    For Each cell In ActiveWorkbook.ActiveSheet.Range("A2:A" & Cells(Rows.Count, 2).End(xlUp).Row)
        If Trim(cell) <> "" Then
            ' See if Ticker exists inside of tickerList
            If tickerList.Contains(Cells(cell.Row, 1).value) Then
                ' It does exist in tickerList.
                ' Let's save that index from tickerList for ease of use.
                index = tickerList.indexOf(Cells(cell.Row, 1).value, 0)
                
                ' Is the date earlier than our saved start date for that ticker?
                If CDate(Cells(cell.Row, 2).value) < startDate(index) Then
                    ' It is earlier: save this date as our new start date.
                    startDate(index) = CDate(Cells(cell.Row, 2).value)
                    
                    ' Let's update the open price because this date is earlier.
                    openList(index) = Cells(cell.Row, 3).value
                End If
                
                ' Is the date later than our saved end date for that ticker?
                If CDate(Cells(cell.Row, 2).value) > endDate(index) Then
                    ' It is later: save this date as our new end date.
                    endDate(index) = CDate(Cells(cell.Row, 2).value)
                    
                    ' Let's update the close price because this date is later.
                    closeList(index) = Cells(cell.Row, 6).value
                End If
                
                ' We need to add this entries' volume to our volumeList
                volumeList(index) = volumeList(index) + Cells(cell.Row, 7).value
            
            Else
                ' Well, we didn't find that ticker in our tickerList.
                ' Let's add it in!
                ' Append the stock ticker to tickerList.
                tickerList.Insert tickerList.Count, (Cells(cell.Row, 1).value)
                
                ' Append the Date to both startList and endList
                ' We put it in both because it will simplify expanding upon it later.
                ' If we just put it as the start date, then there would be an error somewhere
                '   along the line after we try to initiate a new end date off of an empty space.
                startDate.Insert startDate.Count, (Cells(cell.Row, 2).value)
                endDate.Insert endDate.Count, (Cells(cell.Row, 2).value)
                
                ' Append the Open price and Close price to openList and closeList.
                ' Our motivations for this mirror those for the date above. Avoid the error.
                openList.Insert openList.Count, (Cells(cell.Row, 3).value)
                closeList.Insert closeList.Count, (Cells(cell.Row, 6).value)
                
                ' Append the Volume to our volumeList.
                volumeList.Insert volumeList.Count, (Cells(cell.Row, 7).value)
            End If
        End If
    Next

    ' Now, we are going to create and populate the following columns...
    '   Ticker
    '   Quarterly Change
    '   Percent Change
    '   Total Stock Volume

    ' Create the new columns for our synthesis of the stock sales.
    Cells(1, 9).value = "Ticker"
    Cells(1, 10).value = "Quarterly Change"
    Cells(1, 11).value = "Percent Change"
    Cells(1, 12).value = "Total Stock Volume"
    ' Format these real quick... Left-aligned, ...
    Range(Cells(1, 9), Cells(1, 12)).HorizontalAlignment = xlLeft
    
    ' We will need to loop thru our data to output our new columns.
    Dim i As Integer
    For i = 0 To tickerList.Count - 1
        
        ' Populate the Ticker column.
        Cells(i + 2, 9).value = tickerList(i)
        
        ' Populate the Quarterly Change column.
        Cells(i + 2, 10).value = closeList(i) - openList(i)
        
        ' Now, change the fill color of the cell to match the value of change.
        ' Positive value = green fill.
        ' Negative value = red fill.
        ' No change = no fill.
        If closeList(i) - openList(i) > 0 Then
            Cells(i + 2, 10).Interior.Color = RGB(150, 255, 150)
            
        ElseIf closeList(i) - openList(i) < 0 Then
            Cells(i + 2, 10).Interior.Color = RGB(255, 150, 150)
            ' Do smthng.
        End If
        
        ' Populate the Percent Change column.
        ' Note: the formula here is (V2 - V1) / |V1| x 100
        ' Where... V2 is the close price, V1 is the open price.
        ' But... we omit the x 100 since our cell formatting takes care of that for us.
        Cells(i + 2, 11).NumberFormat = "0.00%"
        Cells(i + 2, 11).value = (closeList(i) - openList(i)) / Abs(openList(i))
        
        ' Populate the Total Stock Volume column.
        Cells(i + 2, 12).value = volumeList(i)
    Next
    
    ' Now, we are going to display info on...
    '   Greatest % Increase
    '   Greatest % Decrease
    '   Greatest Total Volume
    
    ' Declare and assign the variables to store our highest/lowest values with accompanying Ticker.
    ' tickerHighPer = stock ticker for highest Percent change.
    ' tickerLowPer = stock ticker for lowest percent change.
    ' tickerHighVol = stock ticker for highest volume traded.
    ' numHighPer = value of highest percent change.
    ' numLowPer = value of lowest percent change.
    ' numHighVol = value of highest volume traded.
    Dim tickerHighPer As String, tickerLowPer As String, tickerHighVol As String
    Dim numHighPer As Double, numLowPer As Double, numHighVol As Double
    tickerHighPer = ""
    tickerLowPer = ""
    tickerHighVol = ""
    numHighPer = 0
    numLowPer = 0
    numHighVol = 0
    
    ' Create the columns and rows to display these values!
    Cells(1, 15).value = "Ticker"
    Cells(1, 16).value = "Value"
    
    Cells(2, 14).value = "Greatest % Increase"
    Cells(3, 14).value = "Greatest % Decrease"
    Cells(4, 14).value = "Greatest Total Volume"
    
    ' Let's format those now: left-aligned, ...
    Range(Cells(1, 15), Cells(1, 16)).HorizontalAlignment = xlLeft
    Range(Cells(2, 14), Cells(4, 14)).HorizontalAlignment = xlLeft
    
    ' Now, we loop thru the column "I" and grab relevant information for making our new columns.
    For Each cell In ActiveWorkbook.ActiveSheet.Range("I2:I" & Cells(Rows.Count, 2).End(xlUp).Row)
        
        ' Is this entry the greatest percent increase?
        If Cells(cell.Row, cell.Column + 2).value > numHighPer Then
            numHighPer = Cells(cell.Row, cell.Column + 2).value
            tickerHighPer = Cells(cell.Row, cell.Column).value
        End If
        
        ' Is this entry the greatest percent decrease?
        If Cells(cell.Row, cell.Column + 2).value < numLowPer Then
            numLowPer = Cells(cell.Row, cell.Column + 2).value
            tickerLowPer = Cells(cell.Row, cell.Column).value
        End If
        
        ' Is this entry the greatest volume traded?
        If Cells(cell.Row, cell.Column + 3).value > numHighVol Then
            numHighVol = Cells(cell.Row, cell.Column + 3).value
            tickerHighVol = Cells(cell.Row, cell.Column).value
        End If
    Next
    
    ' Input the values into the proper cells for the data retrieved above.
    ' Note: first, we format the percent cells (2, 16) and (3, 16).
    Range(Cells(2, 16), Cells(3, 16)).NumberFormat = "0.00%"
    
    Cells(2, 15).value = tickerHighPer
    Cells(3, 15).value = tickerLowPer
    Cells(4, 15).value = tickerHighVol
    
    Cells(2, 16).value = numHighPer
    Cells(3, 16).value = numLowPer
    Cells(4, 16).value = numHighVol
    
End Sub

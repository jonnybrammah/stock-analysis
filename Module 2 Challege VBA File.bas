Attribute VB_Name = "Module1"
Sub StockMarket()

Dim column As Integer
Dim lrow As Long
Dim Volume As Double
Dim TickerNum As Integer
Dim Op As Double
Dim Clos As Double
Dim Change As Double
Dim PercChange As Double

column = 1
TickerNum = 2
Volume = 0
lrow = Cells(Rows.Count, 2).End(xlUp).Row

'Initially define Opening Price as First Stock's first Open Value
Op = Cells(2, 3).Value

'Write the titles of the columns in the sheet
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'''Determines the total volume for each ticker symbol and writes the ticker symbol next to the total volume'''

For i = 2 To lrow

    If Cells(i + 1, column).Value = Cells(i, column).Value Then
        'Increase volume
        Volume = Volume + Cells(i, 7).Value
    
    ElseIf Cells(i + 1, column).Value <> Cells(i, column).Value Then
        
        'Increase volume one last time
         Volume = Volume + Cells(i, 7).Value
        
        'Record Ticker Name and Value
        Cells(TickerNum, 9).Value = Cells(i, column).Value
        Cells(TickerNum, 12).Value = Volume
        
        'Define Closing Price
        Clos = Cells(i, 6).Value
        
        'Calculate Absolute Change in Price
        Change = Clos - Op
        
        'Calculate Percentage Change in Price
        PercChange = (Change / Op) * 100
               
        'Record Absolute Change
        Cells(TickerNum, 10).Value = Change
        
        'Record Percentage Change
        Cells(TickerNum, 11).Value = PercChange
        
        'Now redefine Opening Price as first instance of Open of new stock
        Op = Cells(i + 1, 3).Value
        
        'Reset Other values to zero and increase TickerNum
        Volume = 0
        TickerNum = TickerNum + 1
        Clos = 0
   
    End If

Next i

'For loop to color Absolute and Percentage Change
For i = 2 To lrow

    If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
        Cells(i, 11).Interior.ColorIndex = 4
    ElseIf Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
        Cells(i, 11).Interior.ColorIndex = 3
    End If
    
Next i


'''''''''''''''''''''''''''
'''''''''' BONUS ''''''''''
'''''''''''''''''''''''''''

'Determine Greatest Percentage Increase
Dim GreatPerc As Double
'Initially set Greatest Percentage as percentage increase of first stock
GreatPerc = Cells(2, 11).Value

'For loop to determine if the next Percent Increase is larger than the current Greatest Percentage, and if so replace Greatest percentage with that value
For i = 2 To lrow

    If Cells(i, 11) > GreatPerc Then
        GreatPerc = Cells(i, 11).Value
        GreatPercTick = Cells(i, 9).Value
    End If
    
Next i

'Record Greatest Percentage Increase
Cells(2, 15).Value = "Greatest % Increase"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 16).Value = GreatPercTick
Cells(2, 17).Value = GreatPerc

'Determine Greatest Percentage Decrease
Dim GreatPercDe As Double
'Same idea here: Initially set the greatest decrease as percentage increase of first stock
GreatPercDe = Cells(2, 11).Value

'For loop to compare each percentage increase with current Greatest Decrease and if it is *smaller*, replace Greatest Decrease with that value
For i = 2 To lrow

    If Cells(i, 11) < GreatPercDe Then
        GreatPercDe = Cells(i, 11).Value
        GreatPercDeTick = Cells(i, 9).Value
    End If
    
Next i

'Record Greatest Percentage Decrease
Cells(3, 15).Value = "Greatest % Decrease"
Cells(3, 16).Value = GreatPercDeTick
Cells(3, 17).Value = GreatPercDe

'Determine Greatest Volume
Dim GreatVol As Double
'Same idea again, but this time compare volumes
GreatVol = Cells(2, 12).Value

For i = 2 To lrow

    If Cells(i, 12) > GreatVol Then
        GreatVol = Cells(i, 12).Value
        GreatVolTick = Cells(i, 9).Value
    End If
    
Next i

'Record Greatest Volume
Cells(4, 15).Value = "Greatest Volume"
Cells(4, 16).Value = GreatVolTick
Cells(4, 17).Value = GreatVol


End Sub


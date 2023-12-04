Sub Stock_Data()

'Declare a worksheet
Dim ws As Worksheet

'Loop through all stocks / years
For Each ws In Worksheets

'Declare variables and set initial values as 0
Dim Ticker As String
Ticker = " "
Dim Ticker_Volume As Double
Ticker_Volume = 0

Dim Open_Price As Double
Open_Price = 0
Dim Close_Price As Double
Close_Price = 0
Dim Yearly_Change As Double
Yearly_Change = 0
Dim Percent_Change As Double
Percent_Change = 0

Dim MAX_Ticker As String
MAX_Ticker = " "
Dim MIN_Ticker As String
MIN_Ticker = " "
Dim MAX_Volume_Ticker As String
MAX_Volume_Ticker = " "
Dim MAX_Percent As Double
MAX_Percent = 0
Dim MIN_Percent As Double
MIN_Percent = 0
Dim MAX_Ticker_Volume As Double
MAX_Ticker_Volume = 0

'Keep track of the location for each ticker name
Dim Summary_Table_Row As Long
Summary_Table_Row = 2

'Declare variable for rows and Lastrow
Dim Lastrow As Long
Dim i As Long
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Create headers for the output
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Set initial Open Stock Price
Open_Price = ws.Cells(2, 3).Value

'Loop through all stocks
For i = 2 To Lastrow

' Check if we are still within the stock, if it is not...
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

' Set the Ticker name
Ticker = ws.Cells(i, 1).Value

'Set Yearly price change
Close_Price = ws.Cells(i, 6).Value
Yearly_Change = Close_Price - Open_Price

If Open_Price <> 0 Then
Percent_Change = (Yearly_Change / Open_Price) * 100
Else: Percent_Change = 0
End If

' Add to the Total stock volume
Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value

' Print the Ticker in the Summary Table
ws.Range("I" & Summary_Table_Row).Value = Ticker

' Print the Yearly Change in the Summary Table
ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

'Conditional formatting
If (Yearly_Change > 0) Then
'Fill column with GREEN color - good
ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
ElseIf (Yearly_Change <= 0) Then
'Fill column with RED color - bad
ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
End If

' Print the Percent change in the Summary Table
ws.Range("K" & Summary_Table_Row).Value = Str(Percent_Change) & "%"

'Conditional formatting
If (Percent_Change > 0) Then
'Fill column with GREEN color - good
ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
ElseIf (Percent_Change <= 0) Then
'Fill column with RED color - bad
ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
End If


' Print the Total stock volume to the Summary Table
ws.Range("L" & Summary_Table_Row).Value = Ticker_Volume

'Reset variable value
Summary_Table_Row = Summary_Table_Row + 1
Open_Price = ws.Cells(i + 1, 3).Value
Close_Price = 0
Yearly_Change = 0

'Greatest value calculations
If (Percent_Change > MAX_Percent) Then
MAX_Percent = Percent_Change
MAX_Ticker = Ticker
ElseIf (Percent_Change < MIN_Percent) Then
MIN_Percent = Percent_Change
MIN_Ticker = Ticker
End If
                       
If (Ticker_Volume > MAX_Ticker_Volume) Then
MAX_Ticker_Volume = Ticker_Volume
MAX_Volume_Ticker = Ticker
End If

'Reset variables for Percent change & Ticker volume
Percent_Change = 0
Ticker_Volume = 0

'if the same Ticker name, keep adding up values
Else: Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value

End If

'Output for Greatest values
ws.Range("Q2").Value = Str(MAX_Percent) & "%"
ws.Range("Q3").Value = Str(MIN_Percent) & "%"
ws.Range("Q4").Value = MAX_Ticker_Volume
ws.Range("P2").Value = MAX_Ticker
ws.Range("P3").Value = MIN_Ticker
ws.Range("P4").Value = MAX_Volume_Ticker

Next i

Next ws

End Sub


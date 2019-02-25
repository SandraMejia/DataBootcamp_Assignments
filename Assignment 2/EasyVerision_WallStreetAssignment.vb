'The VBA of Wall Street
'Assignment 2
'Sandra Mejia Avendaño

'1. Solution for easy assignment

Sub Stockdata()

Dim ticker As String 
Dim vol As Single 'Sumation of total volume of each ticker
Dim i As Long 'Counter for rows
Dim t As Long 'Counter for tickers

Cells(1, 9) = "<ticker>"
Cells(1, 10) = "<vol>"

i = 2 'row counter
t = 2 'ticker counter

Do While Cells(i, 1) <> Empty
    ticker = Cells(i, 1)
    vol = 0
        Do While Cells(i, 1) = ticker
            vol = vol + Cells(i, 7)
            i = i + 1
        Loop
    Cells(t, 9) = ticker 'Display ticker symbol
    Cells(t, 10) = vol 'Display total volume of stock
    t = t + 1
Loop

End Sub
'The VBA of Wall Street
'Assignment 2
'Sandra Mejia Avendaño

'2. Solution for moderate assignment

Sub Stockdata()

Dim ticker As String
Dim vol As Single 'Sumation of total volume of each ticker
Dim i As Long 'Counter for rows
Dim t As Long 'Counter for tickers
Dim op As Long 'opening value
Dim cl As Long 'closing value


Cells(1, 9) = "<ticker>"
Cells(1, 10) = "<delta>"
Cells(1, 11) = "<%delta>"
Cells(1, 12) = "<vol>"

Columns(10).NumberFormat = "#0.00"
Columns(11).NumberFormat = "#0.00%"

i = 2 'row counter
t = 2 'ticker counter

Do While Cells(i, 1) <> Empty
    ticker = Cells(i, 1)
    op = Cells(i, 3) 'Determe opening value (First value)
    vol = 0
        Do While Cells(i, 1) = ticker
            vol = vol + Cells(i, 7)
            i = i + 1
        Loop
    cl = Cells(i, 6) 'Determine closing value (Last value)
    Cells(t, 9) = ticker 'Display ticker symbol
    Cells(t, 10) = cl - op 'Display yearly change
            If Cells(t, 10) < 0 Then 'Format column according to yearly change
                Cells(t, 10).Interior.ColorIndex = 3
                Else: Cells(t, 10).Interior.ColorIndex = 4
            End If
    'Display percent change over the year
        If op = 0 Then
            Cells(t, 11) = Empty 'Avoid division by zero
            Else: Cells(t, 11) = (cl - op) / op
        End If
    Cells(t, 12) = vol 'Display total volume of stock

    t = t + 1
Loop

End Sub

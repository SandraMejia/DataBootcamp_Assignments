'The VBA of Wall Street
'Assignment 2
'Sandra Mejia Avendaño

'3. Solution for hard assignment

Sub Stockdata()

Dim ticker As String
Dim vol As Single 'Sumation of total volume of each ticker
Dim i As Long 'Counter for rows
Dim t As Long 'Counter for tickers
Dim op As double 'opening value
Dim cl As double 'closing value
Dim maxincrease As Double 'value of maximum increase
Dim maxdecrease As Double 'value of maximum decrease
Dim maxvolume As Single 'value of maximum total volume
Dim maxincreaseticker As String 'ticker corresponding to maximum increase
Dim maxdecreaseticker As String 'ticker corresponding to maximum decrease
Dim maxvolumeticker As String 'ticker corresponding to maximum total volume

Cells(1, 9) = "<ticker>"
Cells(1, 10) = "<delta>"
Cells(1, 11) = "<%delta>"
Cells(1, 12) = "<vol>"

Cells(2, 14) = "Greatest % Increase"
Cells(3, 14) = "Greatest % Decrease"
Cells(4, 14) = "Greatest total volume"
Cells(1, 15) = "<ticker>"
Cells(1, 16) = "value"

Columns(10).NumberFormat = "#0.00"
Columns(11).NumberFormat = "#0.00%"
Range("P2:P3").NumberFormat = "#0.00%"


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

    If Cells(t, 11) > maxincrease Then 'look for maximum increase
            maxincrease = Cells(t, 11)
            maxincreaseticker = Cells(t, 9)
        ElseIf Cells(t, 11) < maxdecrease Then 'look for maximum decrease
            maxdecrease = Cells(t, 11)
            maxdecreaseticker = Cells(t, 9)
        End If
    If Cells(t, 12) > maxvolume Then 'look for maximum total volume
            maxvolume = Cells(t, 12)
            maxvolumeticker = Cells(t, 9)
        End If
    t = t + 1
Loop

'Create summary table
Cells(2, 15) = maxincreaseticker
Cells(3, 15) = maxdecreaseticker
Cells(4, 15) = maxvolumeticker
Cells(2, 16) = maxincrease
Cells(3, 16) = maxdecrease
Cells(4, 16) = maxvolume

End Sub
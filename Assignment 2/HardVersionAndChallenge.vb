'The VBA of Wall Street
'Assignment 2
'Sandra Mejia Avendaño

'3. Solution for hard assignment
'   Challenge: Run script on all sheets at once

Sub Stockdata()

Dim ticker As String
Dim vol As Single 'Sumation of total volume of each ticker
Dim i As Long 'Counter for rows
Dim t As Long 'Counter for tickers
Dim op As Double 'opening value
Dim cl As Double 'closing value
Dim maxincrease As Double 'value of maximum increase
Dim maxdecrease As Double 'value of maximum decrease
Dim maxvolume As Single 'value of maximum total volume
Dim maxincreaseticker As String 'ticker corresponding to maximum increase
Dim maxdecreaseticker As String 'ticker corresponding to maximum decrease
Dim maxvolumeticker As String 'ticker corresponding to maximum total volume

Dim ws As Variant

For Each ws In Worksheets

    ws.Cells(1, 9) = "<ticker>"
    ws.Cells(1, 10) = "<delta>"
    ws.Cells(1, 11) = "<%delta>"
    ws.Cells(1, 12) = "<vol>"

    ws.Cells(2, 14) = "Greatest % Increase"
    ws.Cells(3, 14) = "Greatest % Decrease"
    ws.Cells(4, 14) = "Greatest total volume"
    ws.Cells(1, 15) = "<ticker>"
    ws.Cells(1, 16) = "value"

    ws.Columns(10).NumberFormat = "#0.00"
    ws.Columns(11).NumberFormat = "#0.00%"
    ws.Range("P2:P3").NumberFormat = "#0.00%"

    maxincrease = 0
    maxdecrease = 0
    maxvolume = 0
        
    i = 2 'row counter
    t = 2 'ticker counter

    Do While ws.Cells(i, 1) <> Empty
        ticker = ws.Cells(i, 1)
        op = ws.Cells(i, 3) 'Determe opening value (First value)
        vol = 0
            Do While ws.Cells(i, 1) = ticker
                vol = vol + ws.Cells(i, 7)
                i = i + 1
            Loop
        cl = ws.Cells(i - 1, 6) 'Determine closing value (Last value)

        ws.Cells(t, 9) = ticker 'Display ticker symbol
        ws.Cells(t, 10) = cl - op 'Display yearly change
                If ws.Cells(t, 10) < 0 Then 'Format column according to yearly change
                    ws.Cells(t, 10).Interior.ColorIndex = 3
                    Else: ws.Cells(t, 10).Interior.ColorIndex = 4
                End If
        'Display percent change over the year
        If op = 0 Then
            ws.Cells(t, 11) = Empty 'Avoid division by zero
            Else: ws.Cells(t, 11) = (cl - op) / op
        End If
        ws.Cells(t, 12) = vol 'Display total volume of stock

        If ws.Cells(t, 11) > maxincrease Then 'look for maximum increase
                maxincrease = ws.Cells(t, 11)
                maxincreaseticker = ws.Cells(t, 9)
        ElseIf ws.Cells(t, 11) < maxdecrease Then 'look for maximum decrease
                maxdecrease = ws.Cells(t, 11)
                maxdecreaseticker = ws.Cells(t, 9)
        End If
        If ws.Cells(t, 12) > maxvolume Then 'look for maximum total volume
                maxvolume = ws.Cells(t, 12)
                maxvolumeticker = ws.Cells(t, 9)
            End If
        t = t + 1
    Loop

    'Create summary table
    ws.Cells(2, 15) = maxincreaseticker
    ws.Cells(3, 15) = maxdecreaseticker
    ws.Cells(4, 15) = maxvolumeticker
    ws.Cells(2, 16) = maxincrease
    ws.Cells(3, 16) = maxdecrease
    ws.Cells(4, 16) = maxvolume

Next ws
End Sub

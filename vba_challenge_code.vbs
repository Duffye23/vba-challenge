Sub stock()
    'looping through each sheet
    For Each ws In Worksheets
      ws.Activate
        Dim ticker As String
        Dim openprice As Double
        Dim closeprice As Double
        Dim vol As Double
        Dim table As Integer
        Dim change As Double
        Dim percent As Double
        Dim increase As Double
        Dim decrease As Double
        Dim highvol As Double
        Dim incsymb As String
        Dim decsymb As String
        Dim volsymb As String
        
        'finding the last row without counting'
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        'need a second last row counter for the second loop
        lstrw = Cells(Rows.Count, 11).End(xlUp).Row
        'and a third for the third loop
        ltr = Cells(Rows.Count, 12).End(xlUp).Row
        'setting the volume variable to 0
        vol = 0
        table = 2
        increase = 0
        decrease = 0
        highvol = 0
        'get the first value's info (ticker, opening price)
    
        For i = 2 To lastrow
            
            'check if we are in the same stock, if we are capture the required info'
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                'get the closing price
                closeprice = Cells(i, 6).Value
                'Get the ticker symbol
                ticker = Cells(i, 1).Value
                'add to the total volume
                vol = vol + Cells(i, 7).Value
                'Print each ticker symbol
                Range("I" & table).Value = ticker
                'Print the total volume of each symbol
                Range("L" & table).Value = vol
                Range("N" & table).Value = closeprice
                change = closeprice - openprice
                Range("J" & table).Value = change
                percent = change / openprice
                Range("K" & table).Value = percent
                'add a row to the table
                table = table + 1
                'reset the total volume so the next symbol starts at 0
                vol = 0
                change = 0
            
            ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                openprice = Cells(i, 3).Value
                Range("M" & table).Value = openprice
            
            Else
                'if the cell matches the previous, just add to the volume total
                vol = vol + Cells(i, 7).Value
            End If
            
            If Range("J" & table).Value > 0 Then
                    Range("J" & table).Interior.ColorIndex = 4
                  ElseIf Range("J" & table).Value < 0 Then
                    Range("J" & table).Interior.ColorIndex = 3
            End If
            
            If Range("K" & table).Value > 0 Then
                    Range("K" & table).Interior.ColorIndex = 4
                  ElseIf Range("J" & table).Value < 0 Then
                    Range("K" & table).Interior.ColorIndex = 3
            End If
                
        Next i
     
    'Building the second table of values (Greatest an Least %change and highest volume traded
        For j = 2 To lstrw
            If Cells(j, 11).Value > increase Then
                increase = Cells(j, 11).Value
                incsymb = Cells(j, 9).Value
                Cells(2, 17).Value = incsymb
            End If
            Cells(2, 18).Value = increase
        
        
            If Cells(j, 11).Value < decrease Then
                decrease = Cells(j, 11).Value
                decsymb = Cells(j, 9).Value
                Cells(3, 17).Value = decsymb
            End If
            Cells(3, 18).Value = decrease
        
        Next j
    
        For k = 2 To ltr
            If Cells(k, 12).Value > highvol Then
                highvol = Cells(k, 12).Value
                volsymb = Cells(k, 9).Value
                Cells(4, 17).Value = volsymb
            End If
            Cells(4, 18).Value = highvol
        
        Next k
        
    Next ws
    
End Sub


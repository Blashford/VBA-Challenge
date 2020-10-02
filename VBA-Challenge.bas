Attribute VB_Name = "Module1"
Sub stocks():
    
    Dim current As Worksheet
    
    For Each current In Worksheets
        current.Activate
    
        'just hardcoding these for each sheet
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        
        Dim count As Integer
        Dim row_count As LongLong
        'finding the end of the data so it knows where to stop
        row_count = Range("A2").End(xlDown).Row
        
        'starts on row 2 so we miss the headers
        count = 2
        
        Dim i As LongLong
        Dim current_cell As String
        Dim next_cell As String
        Dim last_cell As String
        Dim change1 As Double
        Dim change2 As Double
        Dim total_stock As LongLong
        
        total_stock = 0
        
        For i = 2 To row_count
            
            'looking ahead and behind
            current_cell = Cells(i, 1).Value
            next_cell = Cells(i + 1, 1).Value
            last_cell = Cells(i - 1, 1).Value
            
            'to save the opening price for the year
            If current_cell <> last_cell Then
                change1 = Cells(i, 3).Value
            End If
            
            'keeping track of the total stock
            
            total_stock = total_stock + Cells(i, 7).Value
            
            
            'moving and assigning data
            If current_cell <> next_cell Then
                
                change2 = Cells(i, 6).Value
                
                Cells(count, 9).Value = current_cell
                
                Cells(count, 10).Value = change2 - change1
                
                'PLNT really messed this up for a while
                If change1 = 0 Then
                    Cells(count, 11).Value = change2
                Else
                    Cells(count, 11).Value = Cells(count, 10).Value / change1
                End If
                 
                Cells(count, 11) = FormatPercent(Cells(count, 11), 2)
                
                Cells(count, 12).Value = total_stock
                
                If Cells(count, 10).Value > 0 Then
                    Cells(count, 10).Interior.ColorIndex = 4
                Else
                    Cells(count, 10).Interior.ColorIndex = 3
                End If
                
                'resetting the total stock and incrementing the counter
                total_stock = 0
                count = count + 1
            
            End If
        Next i
        
        'hardcoding again
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        
        Dim ticker_cell As String
        Dim volume_cell As LongLong
        Dim j As LongLong
        
        row_count = Range("k2").End(xlDown).Row
        
        Dim rng As Range
        Dim min As Double
        Dim max As Double
        
        'finding the min and max values and assigning them
        Set rng = Range("K:K")
        min = Application.WorksheetFunction.min(rng)
        Cells(3, 17).Value = min
        max = Application.WorksheetFunction.max(rng)
        Cells(2, 17).Value = max
        
        For j = 2 To row_count
            
            current_cell = Cells(j, 11).Value
            ticker_cell = Cells(j, 9).Value
            volume_cell = Cells(j, 12).Value
            
            'matching the max value to its ticker
            If current_cell = max Then
                Cells(2, 16).Value = ticker_cell
            End If
            
            'matching the min value to its ticker
            If current_cell = min Then
                Cells(3, 16).Value = ticker_cell
            End If
            
            'finding the biggest total volume
            If Cells(4, 17).Value < volume_cell Then
                Cells(4, 17).Value = volume_cell
                Cells(4, 16).Value = ticker_cell
            End If
        
        Next j
        
        
        
        'cells need to have a value in them to format percent I guess? or it overwrote the formatting
        'either way I put it down here
        Cells(2, 17) = FormatPercent(Cells(2, 17), 2)
        Cells(3, 17) = FormatPercent(Cells(3, 17), 2)
        
        current.Columns.AutoFit
    Next
    
    
    
End Sub


Attribute VB_Name = "Module5"
Sub stockcount2()

    'set headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'define variables
    Dim tickercheck As String
    Dim ticker As String
    Dim summaryrow As Integer
    Dim volume As Double
    Dim total_volume As Double
    Dim total_yearly As Double
    Dim opening As Double
    Dim closing As Double
    Dim percent_change As Double
    Dim closing_final As Double
    Dim opening_final As Double
    Dim tickerprecheck As String
    Dim lastrow As Long
    
    
    'set summary value
    summaryrow = 2
    
    'set last row value
    lastrow = ActiveSheet.Range("A1").CurrentRegion.Rows.Count
    
    
    'start for loop
    For i = 2 To lastrow
    
        'assign values for loop
        tickerprecheck = Cells(i - 1, 1)
        tickercheck = Cells(i + 1, 1)
        ticker = Cells(i, 1)
        volume = Cells(i, 7)
        opening = Cells(i, 3)
        closing = Cells(i, 6)

            'end of series calculations
            If tickercheck <> ticker Then
                closing_final = closing
                total_volume = total_volume + volume
                total_yearly = (closing_final - opening_final)
    
                'calculate percent change (prevent dividing by 0)
                If opening_final <> 0 Then
                    percent_change = (total_yearly / opening_final)
                    Else: percent_change = 0
                End If
    
                'display results
                Cells(summaryrow, 9).Value = ticker
                Cells(summaryrow, 12).Value = total_volume
                Cells(summaryrow, 10).Value = total_yearly
                Cells(summaryrow, 11).Value = percent_change
                Cells(summaryrow, 11).NumberFormat = "0.00%"
            
                'continue summary
                summaryrow = summaryrow + 1
        
                'reset variables
                total_volume = 0
                total_yearly = 0
                percent_change = 0
                closing_final = 0
                opening_final = 0
    
            'continue series calculations
            Else
            total_volume = total_volume + volume
        
            End If
        
        'obtain opening price
        If tickerprecheck <> ticker Then
            opening_final = opening
            
        End If
    
    Next i



'Color yearly change

    For i = 2 To summaryrow
        If Cells(i, 10) > 0 Then
            Cells(i, 10).Interior.Color = RGB(0, 255, 0)

        ElseIf Cells(i, 10) < 0 Then
            Cells(i, 10).Interior.Color = RGB(255, 0, 0)

        End If

    Next i



'find greatests
    
    'set headers
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'define variables
    Dim greatest As Double
    Dim decrease As Double
    Dim greatest_volume As Double
    Dim decrease_ticker As String
    Dim greatest_ticker As String
    Dim volume_ticker As String
    
    'format percentages
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"
    
    'find greatest % increase and display
    For i = 2 To summaryrow
        If Cells(i, 11) > greatest Then
            greatest = Cells(i, 11)
            greatest_ticker = Cells(i, 9)
            Range("Q2").Value = greatest
            Range("P2").Value = greatest_ticker
        
        'find greatest % decrease and display
        ElseIf Cells(i, 11) < decrease Then
            decrease = Cells(i, 11)
            decrease_ticker = Cells(i, 9)
            Range("Q3").Value = decrease
            Range("P3").Value = decrease_ticker
            
        'find greatest total volume and display
        ElseIf Cells(i, 12) > greatest_volume Then
            greatest_volume = Cells(i, 12)
            volume_ticker = Cells(i, 9)
            Range("Q4").Value = greatest_volume
            Range("P4").Value = volume_ticker
            
        End If
        
    Next i
    

End Sub

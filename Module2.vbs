Attribute VB_Name = "Module1"
Sub MultipleYearStockData():

    For Each ws In ThisWorkbook.Worksheets
    
        Dim ThisWorksheet As String
      
        Dim i, j, TickCount As Long
        Dim LastRowColumnA, LastRow1 As Long
        Dim PerChange, GreatIncrease, GreatDecrease, GreatVolume As Double
        
    ThisWorksheet = ws.Name
        
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        TickCount = 2
        
      
        j = 2
        
      
            LastRowColumnA = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
            'Loop through all rows
            For i = 2 To LastRowColumnA
            
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Write ticker in column I (#9)
               ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
                
                'Calculate and write Yearly Change in column J (#10)
               ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                   
                    If ws.Cells(TickCount, 10).Value < 0 Then
                
                    'Set cell background color to red
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                
                   Else
        
                 'Set cell background color to green
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                   'Calculate and write percent change in column K (#11)
                    If ws.Cells(j, 3).Value <> 0 Then
                    PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    'Percent formating
                    ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                    
                    Else
                    
                   ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                    
                   End If
                    
                'Calculate and write total volume in column L (#12)
                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                
                'Increase TickCount by 1
                TickCount = TickCount + 1
                
                'Set new start row of the ticker block
                j = i + 1
                
                End If
            
            Next i
            
     
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
       
   
        GreatVolume = ws.Cells(2, 12).Value
        GreatIncrease = ws.Cells(2, 11).Value
        GreatDecrease = ws.Cells(2, 11).Value
        

            For i = 2 To LastRowI
            
                If ws.Cells(i, 12).Value > GreatVolume Then
               GreatVolume = ws.Cells(i, 12).Value
               ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVolume = GreatVolume
            
                End If
                
                If ws.Cells(i, 11).Value > GreatIncrease Then
                GreatIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                 Else
                
                GreatIncrease = GreatIncrease
                
                End If
                
               If ws.Cells(i, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecrease = GreatDecrease
                
                End If
                
           ws.Cells(2, 17).Value = Format(GreatIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVolume, "Scientific")
            
           Next i
            

        ThisWorkbook.Worksheets(ThisWorksheet).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub

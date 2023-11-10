Attribute VB_Name = "HungryHungryHippo"
Sub RamMuncher()

    'add functionality for code to operate through each sheet
    For Each ws In Worksheets

        'establish values
        Dim Ticker As String
        'make a value for resetting values within loop for context
        Dim Clear As Integer
            Clear = 0
        Dim Volume As Long
            Volume = Clear
        
        Dim Opening As Double
            Opening = ws.Cells(2, 3).Value
        
        Dim Closing As Double
        
        Dim YearChange As Double
        
        Dim Percentage As Double
        'Row location for code to place information
        Dim TableRow As Integer
            TableRow = 2
        'Non static ending value for loops becuase each sheet has different amounts of rows
        EndRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'loop sequentially through ticker tags and place them in the new column
        For n = 2 To EndRow

            ' grab the data from all cells with the same ticker tag
            If ws.Cells(n + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              Ticker = ws.Cells(n, 1).Value

              Volume = Volume + ws.Cells(n, 7).Value

              ' adds the ticker tag to the summary column
              ws.Range("I" & TableRow).Value = Ticker

              ' adds the volume to the summary column
              ws.Range("L" & TableRow).Value = Volume

              Closing = ws.Cells(i, 6).Value

              YearChange = (Closing - Opening)
              
              ' adds the yearly change to the summary column
              ws.Range("J" & TableRow).Value = YearChange

              'just to stop any divide by zero oopsies
                If Opening = 0 Then
                    Percentage = 0
                
                Else
                    Percentage = YearChange / Opening
                
                End If

              ws.Range("K" & TableRow).Value = Percentage
              'format range to percentages with 2 decimals
              ws.Range("K" & TableRow).NumberFormat = "0.00%"
   
              TableRow = TableRow + 1

              Volume = Clear

              Opening = ws.Cells(n + 1, 3)
            
            Else
              
               Volume = Volume + ws.Cells(n, 7).Value

            End If
            
        ' continue to the next unique ticker tag and repeat
        Next n

     EndrowTable = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Conditionally format the yearly change column
        For n = 2 To EndrowTable
            
            If ws.Cells(n, 10).Value > 0 Then
                ws.Cells(n, 10).Interior.ColorIndex = 10
            
            Else
                ws.Cells(n, 10).Interior.ColorIndex = 3
            
            End If
        
        Next n

        For n = 2 To EndrowTable
        
            If ws.Cells(n, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & EndrowTable)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                'formats cells to a percentage with 2 decimals
                ws.Cells(2, 17).NumberFormat = "0.00%"

            ElseIf ws.Cells(n, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & EndrowTable)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            ElseIf ws.Cells(n, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & EndrowTable)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next n
    
    Next ws
        
End Sub

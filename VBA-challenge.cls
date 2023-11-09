VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stocks()

 For Each ws In Worksheets
 'Establish Variables
        Dim WorksheetName As String
        Dim i As Long
        
        Dim tickerCount As Long
        Dim openPrice As Double
        Dim closePrice As Double
        Dim yearlyChange As Double
        Dim totalVolume As Double
        
         WorksheetName = ws.Name
    'Establish Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly_Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total  stock volume"
         tickerCount = 2
         
          'locating the last row  for loop
          lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
         
    ' Start loop
          For i = 2 To lastRow
              'Checking if still within the same ticker, if not...
              If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                  'Write ticker in column I (#9)
                ws.Cells(tickerCount, 9).Value = ws.Cells(i, 1).Value
                openPrice = ws.Cells(2, 3).Value
                closePrice = ws.Cells(i, 6).Value
                yearlyChange = closePrice - openPrice
                'set value for yearly_change
               ws.Cells(tickerCount, 10).Value = yearlyChange
               'set value for percent_change
               If openPrice = 0 Then
                ws.Cells(tickerCount, 11).Value = 0
                Else
               ws.Cells(tickerCount, 11).Value = yearlyChange / openPrice
                ws.Cells(tickerCount, 11).NumberFormat = "0.00%"
               End If
               totalVolume = 0
               totalVolume = totalVolume + ws.Cells(i, 7).Value
               ws.Cells(tickerCount, 12).Value = totalVolume
               'Reset the opening price
              openPrice = ws.Cells(i + 1, 3)
                 tickerCount = tickerCount + 1
                 
               totalVolume = 0
            Else
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            ws.Cells(tickerCount, 12).Value = totalVolume
        End If
          Next i
          
            lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
             'yearly change color formatting depending on positive and negative numbers
        For j = 2 To lastrow_summary_table
            
            If ws.Cells(j, 10).Value > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 10
            
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            
            End If
        
        Next j
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
          For i = 2 To lastrow_summary_table
        
            'locating the maximum percent change
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            'locating the minimum percent change
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'locating the maximum volume
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary_table)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
          Worksheets(WorksheetName).Columns("A:Z").AutoFit
        
        Next ws
End Sub

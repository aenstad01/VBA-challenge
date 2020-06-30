Sub Stocks()

        
Dim ws As Worksheet
Dim Total_Stock_Volume As Double
Dim Ticker_Name As String
Dim Year_Change As Double
Dim Start As Long
Dim Percent_Change As Double
Dim New_Value As Long

    
    Total_Stock_Volume = 0



'1. Add in all of the new column headers in every worksheet
    For Each ws In Worksheets
        'Define the Last Row
            Dim LastRow As Long
            LastRow = ws.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    
                ws.Range("I1").Value = "Ticker"
                ws.Range("J1").Value = "Yearly Change"
                ws.Range("K1").Value = "Percent Change"
                ws.Range("L1").Value = "Total Stock Volume"
                
                ws.Range("P1").Value = "Ticker"
                ws.Range("Q1").Value = "Value"
                ws.Range("O2").Value = "Greatest % Increase"
                ws.Range("O3").Value = "Greatest % Decrease"
                ws.Range("O4").Value = "Greatest Total Volume"
    
    Next


'2. List each unique ticker name under the "ticker" column and find the closing price for each ticker
For Each ws In Worksheets

Dim Summary_row As Integer
Summary_row = 2

  
        ' Set a variable for the <ticker> column
          Dim column As Integer
          column = 1
        
        ' Start on the first row that isn't 0
        Start = 2
        
        
        ' Loop through rows in the column
            For i = 2 To LastRow

        ' Searches for when the value of the next cell is different than that of the current cell
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
            
                Ticker_Name = ws.Cells(i, 1).Value
            
                'Find the Total Stock Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                ws.Range("L" & Summary_row).Value = Total_Stock_Volume
                
                
                'Put the Ticker names in the Summary Table
                ws.Cells(Summary_row, 9).Value = Ticker_Name
                  
                  
                If ws.Cells(Start, 3) = 0 Then
                    
                    For New_Value = Start To i
                    
                        If ws.Cells(New_Value, 3).Value <> 0 Then
                        
                            Start = New_Value
                            
                            Exit For
                            
                        End If
                        
                    Next New_Value
                        
                End If
          
                  
                Year_Change = ws.Cells(i, 6) - ws.Cells(Start, 3)
                Percent_Change = (Year_Change / ws.Cells(Start, 3)) * 100
                  
                ws.Range("J" & Summary_row).Value = Year_Change
                ws.Range("K" & Summary_row).Value = Percent_Change
            
                
                
                'Color code the yearly change
                
                If Year_Change > 0 Then
                
                    ws.Range("J" & Summary_row).Interior.ColorIndex = 4
                    
                    Else
                
                    ws.Range("J" & Summary_row).Interior.ColorIndex = 3
                    
                    End If
                
                Total_Stock_Volume = 0
                Start = i + 1
                Summary_row = Summary_row + 1
                  
                  
                  
                  
            Else
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                End If
                
                Next i
            
    
    'Find the greatest % increase
        Dim j As Long
        Dim High_Percent As Double
        Dim Ticker_Nm As String
        
        j = 2
        High_Percent = ws.Cells(j, 11).Value
        
        For j = 3 To Summary_row
        
            If ws.Cells(j, 11).Value > High_Percent Then
            
                High_Percent = ws.Cells(j, 11).Value
                
                Ticker_Nm = ws.Cells(j, 9).Value
                
                Else
                
            End If
            
            Next j
            
            
       
      'Paste in greatest increase %
      ws.Range("P2").Value = Ticker_Nm
      ws.Range("Q2").Value = High_Percent
                
            
            
    ' Find the greatest % decrease
    Dim k As Long
    Dim Low_Percent As Double
    Dim Ticker As String
    k = 2
    Low_Percent = ws.Cells(k, 11).Value
        
        For k = 3 To Summary_row
        
            If ws.Cells(k, 11).Value < Low_Percent Then
            
                Low_Percent = ws.Cells(k, 11).Value
                
                Ticker = ws.Cells(k, 9).Value
                
                Else
                
            End If
            
            Next k
  
        'Paste in greatest decrease %
      ws.Range("P3").Value = Ticker
      ws.Range("Q3").Value = Low_Percent
      
      
    ' Find the greatest total volume
    Dim v As Long
    Dim Greatest_Volume As Double
    Dim Tickerv As String
    v = 2
    Greatest_Volume = ws.Cells(v, 11).Value
        
        For v = 3 To Summary_row
        
            If ws.Cells(v, 11).Value > Greatest_Volume Then
            
                Greatest_Volume = ws.Cells(v, 11).Value
                
                Tickerv = ws.Cells(v, 9).Value
                
                Else
                
            End If
            
            Next v
  
        'Paste in greatest total volume
      ws.Range("P4").Value = Tickerv
      ws.Range("Q4").Value = Greatest_Volume
  
  
  
  
  
  
  
'This moves on to the next tab
Next ws






End Sub

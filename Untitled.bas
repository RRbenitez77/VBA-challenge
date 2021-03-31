Attribute VB_Name = "Module1"
Sub stock_data()

'Declare and set worksheet




    'Set Dimensions
Dim total As Double
Dim i As Long
Dim change As Double
Dim j As Long
Dim start As Long
Dim rowTotal As Long
Dim percentChange As Double
Dim dailyChange As Double
Dim averageChange As Double
Dim Ticker As String
Dim ws As Worksheet

'Loop through all stocks for one year
For Each ws In Worksheets

   'Set title row
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percent Change"
 ws.Range("L1").Value = "Total Stock Volume"
 
   'Set initial Values
 j = 0
 total = 0
 change = 0
 start = 2
 vol = 0
 'get the last row with data
     rowTotal = ws.Cells(Rows.Count, 1).End(xlUp).Row
 'Summery_Table_Row=2
 
 For i = 2 To rowTotal
       
 'If ticker changes then print results
 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
 
 
       
' Stores results in variables
       
       total = total + ws.Cells(i, 7).Value
       
'Handle zero total volume
       
       If total = 0 Then
       
'print the results
         
         ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
         ws.Range("J" & 2 + j).Value = 0
         ws.Range("K" & 2 + j).Value = "%" & 0
         ws.Range("L" & 2 + j).Value = 0
         
         Else
         
'Find First non zero with starting value
          If ws.Cells(start, 3) = 0 Then
            For find_value = start To i
            If ws.Cells(find_value, 3).Value <> 0 Then
            start = find_value
            
            
            Exit For
      End If
      Next find_value
      
      End If
    
             
'Calculate Change
         change = (ws.Cells(i, 6) - ws.Cells(start, 3))
         percentChange = Round((change / ws.Cells(start, 3) * 100), 2)
         
'start of next stock ticker abreviation
         
         start = i + 1
         
 'start the results
            ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("J" & 2 + j).Value = Round(change, 2)
            ws.Range("K" & 2 + j).Value = "%" & percentChange
            ws.Range("L" & 2 + j).Value = total
            
         
    'colors the positive with green and negatives with red
         
         Select Case change
            Case Is > 0
                  ws.Range("J" & 2 + j).Interior.ColorIndex = 4
            Case Is < 0
                 ws.Range("J" & 2 + j).Interior.ColorIndex = 3
            Case Else
                  ws.Range("J" & 2 + j).Interior.ColorIndex = 0
            End Select
         
         
     End If
      
      'reset variables for new stock ticker abreviations
            total = 0
            change = 0
            j = j + 1
            days = 0
     
     'If ticker is the same add together
   Else
   total = total + ws.Cells(i, 7).Value
   End If
   

   Next i
   

   
   
   Next ws
   
     
     

     
End Sub


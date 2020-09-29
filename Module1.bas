Attribute VB_Name = "Module1"
Sub stockmarket()

Dim ticker As String
Dim yearly_start_value As Double
Dim yearly_end_value As Double
Dim i As Long
Dim j As Integer


Dim startrow As Long
Dim lastrow As Long

Dim tickertotal As Double
Dim yearlychange As Double
Dim percentchange As Single
tickertotal = 0
yearlychange = 0
j = 2

startrow = 2

'this lastrow will count the total number of rows(column A) for the first loop
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'this lastrow will count the total number of rows (Column J) for the second loop
lastrow2 = Cells(Rows.Count, 10).End(xlUp).Row

Cells(1, 10) = "Ticker"
Cells(1, 11) = "Yearly Change"
Cells(1, 12) = "Percent Change"
Cells(1, 13) = "Total Volume"

Cells(1, 18) = "Ticker"
Cells(1, 19) = "Value"
Cells(2, 17) = "Greatest % Increase"
Cells(3, 17) = "Greatest % Decrease"
Cells(4, 17) = "Greatest Total Volume"


'this first loop will generate the columns J through M (lookup values) and then also do conditional format

    For i = 2 To lastrow
        

    If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
    
        ticker = Cells(i, 1).Value
        
        tickertotal = tickertotal + Cells(i, 7).Value
        
                 
        yearly_start_value = Cells(startrow, 3).Value
            
        yearly_end_value = Cells(i, 6).Value
      
        yearlychange = yearly_end_value - yearly_start_value
        
        percentchange = (yearlychange / yearly_start_value) * 100
        
        
                 
        Cells(j, 10).Value = ticker
        Cells(j, 11).Value = yearlychange
        Cells(j, 12).Value = percentchange
        Cells(j, 13).Value = tickertotal
        
        'these helper columns were used to check that the correct values are being pulled (comment-out later)
        'Cells(j, 15).Value = yearly_start_value
        'Cells(j, 16).Value = yearly_end_value
        
        startrow = i + 1
             j = j + 1
        tickertotal = 0
        yearlychange = 0
        percentchange = 0
                  
        
        Else
        
        tickertotal = tickertotal + Cells(i, 7).Value
             
        
    End If

'IMPORTANT NOTE:  this formatting is requiring to hit the RUN button again!!!

           If Cells(i, 12).Value > 0 Then
                Cells(i, 12).Interior.ColorIndex = 4
            
            ElseIf Cells(i, 12).Value < 0 Then
                Cells(i, 12).Interior.ColorIndex = 3
            
        End If
        
        
 Next i
 
    
'VBA CHALLENGE PORTION

'finding the values

    
    max_volume = WorksheetFunction.Max(Range("M:M"))
    Cells(4, 19).Value = max_volume
    
    max_increase = WorksheetFunction.Max(Range("L:L"))
    Cells(2, 19).Value = max_increase
    
    max_decrease = WorksheetFunction.Min(Range("L:L"))
    Cells(3, 19).Value = max_decrease
    
'this second loop is for finding the corresponding ticker for the values, by looping through column J

'IMPORTANT NOTE:  this second loop requires you to hit the RUN button again!!!!


For i = 2 To lastrow2
    If Cells(i, 13).Value = Cells(4, 19).Value Then
        Range("R4").Value = Cells(i, 10).Value
               
        End If
        
    If Cells(i, 12).Value = Cells(3, 19).Value Then
        Range("R3").Value = Cells(i, 10).Value
              
        End If
    If Cells(i, 12).Value = Cells(2, 19).Value Then
        Range("R2").Value = Cells(i, 10).Value
            
    End If

Next i

End Sub

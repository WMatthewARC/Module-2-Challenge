Attribute VB_Name = "Module11"
Sub LoopStocks()



'Define frist work sheet second  altering oder

    
            Dim Ticker As String
            
            Dim YearlyChange As Double
            Dim PercentChange As Double
            Dim TotalStockVolume As Double
            
                ' add open and close to see them sparate
                
            Dim openingValue As Double
            Dim closeingValue As Double
            
            Dim trackDate2 As Integer
              





              
 'something to look at the whole work book or go sheet by sheet
    Dim ws As Worksheet
                  

  'For Each workbookRun In Worksheet
  
  
For Each ws In Worksheets


'------------------------------------------------------for start worksheet loop ---------------------------------------

 
      'adding heders
    ' 06/21 add workbookRun to jump sheets
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
        
    '--------just the dam title rows -----------
      
             
                  
        'to help read code
        Dim A_COLUMN As Integer
            A_COLUMN = 1
        
       
            
    '--------just the dam title rows -----------
        
        
                            
        'Set number values IDK why not for double(s)
                   
                    trackDate2 = 2
                    TotalStockVolume = 0
                    'PercentChange = 0
                    'YearlyChange = 0
                    'i = 1
                    LastI = 1
                     
           'range but find out how to make it verable not hard codeing I:w/e like MAX row = x or soemthing
                        
          
            EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
                    
  
'start loop and adding

    'CC example find unique name out but to output to I2 and down
        'For i = 2 to EndRow
        'If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    
    
    
'------------------ For " I " ------------START-----------------------
    For i = 2 To EndRow

    'Looking for unique cell, if this is not what was in the last cell
   

    If ws.Cells(i + 1, A_COLUMN).Value <> ws.Cells(i, A_COLUMN).Value Then
   
    
    Ticker = ws.Cells(i, A_COLUMN).Value
        
    LastI = LastI + 1
            
                'open and closeing amounts
                
                
        openingVaule = ws.Cells(LastI, 3).Value
                
        closeingValue = ws.Cells(i, 6).Value
    
    
    '------------------------For "SUM of ROW" ----------START------loop for Total Stock Volume---------------------------
         
         For SumRow = LastI To i
         
                'rolling sum for volume
                
                TotalStockVolume = TotalStockVolume + ws.Cells(SumRow, 7).Value

    
        Next SumRow
    
    
    '--------------------------For "SUM of ROW"---------END------------------------------------
    
        'Start with if empty move on
        

            If openingVaule = 0 Then

                PercentChange = closeingValue


        'Math for yearly closeing and opening.
        
            Else
                YearlyChange = closeingValue - openingVaule

                PercentChange = YearlyChange / openingVaule

            End If
        
        
        'output our running totals
            ws.Cells(trackDate2, 9).Value = Ticker
            
            ws.Cells(trackDate2, 10).Value = YearlyChange
            
            ws.Cells(trackDate2, 11).Value = PercentChange
            
            
               'futher output for format
                                            
        ws.Cells(trackDate2, 11).NumberFormat = "0.00%"
            
        ws.Cells(trackDate2, 12).Value = TotalStockVolume
            
        
            'for date moveing to next row
             
                 trackDate2 = trackDate2 + 1
                                                    
   
                                                    
                                                    
          '-------------------------------Reset values -----------------------
          
                                                    
                                                    
                                    'reset numbers to start next step
                                                    
                                           PercentChange = 0
                                           YearlyChange = 0
                                           'closeingValue = 0
                                           'openingVaule = 0
                                           TotalStockVolume = 0
    
  
            '-------------------------Move down ---------------
        'I am lost on why this does this or how this works. i = i is what we always used to move down but I think I need another one not a I like a I.2 or "lastI"
        
        
                LastI = i
        
                            
                            
            End If
            
     ' I loop done J loop inside
     
        Next i
        
        
  'addtional still not getting the math right but I need to move on.
  
  
        
  kEndRow = ws.Cells(Rows.Count, "K").End(xlUp).Row
  
    Increase = 0
    Decrease = 0
    Greatest = 0
  
    For k = 3 To kEndRow
         
         
    'lable
    ws.Range("N1").Value = "Column Name"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker Name"
    ws.Range("P1").Value = "Value"
         
         
         
         
         
    'Define all the things we are going to use to do the loop math
         
         
    last_k = k - 1
    current_k = ws.Cells(k, 11).Value
    prevous_k = ws.Cells(last_k, 11).Value
    volume = ws.Cells(k, 12).Value
    prevous_vol = ws.Cells(last_k, 12).Value
                       
                
                   
                
                
                
                
        
    ' start with math for each in a "if eles end if," loop

    'Increse, Decrease & Gretest

'-----------------------------------------------------------------------

        
    If Increase > current_k And Increase > prevous_k Then

        
        Increase = Increase
        
    ElseIf current_k > Increase And current_k > prevous_k Then
        
        
        Increase = current_k
        
        
        increase_name = ws.Cells(k, 9).Value
        
        
    ElseIf prevous_k > Increase And prevous_k > current_k Then
  
        Increase = prevous_k
    
        increase_name = ws.Cells(last_k, 9).Value
    
    
    
    End If
    
    '--------------------------------------------------------------------------------
    
       
       If Decrease < current_k And Decrease < prevous_k Then

               
                Decrease = Decrease


            ElseIf current_k < Increase And current_k < prevous_k Then

                Decrease = current_k


                decrease_name = ws.Cells(k, 9).Value

            ElseIf prevous_k < Increase And prevous_k < current_k Then

                Decrease = prevous_k

                decrease_name = ws.Cells(last_k, 9).Value

            End If
 
   '-----------------------------------------------------------------------------------------------------
   
   

        If Greatest > volume And Greatest > prevous_vol Then

                Greatest = Greatest

            ElseIf volume > Greatest And volume > prevous_vol Then

                Greatest = volume

                greatest_name = ws.Cells(k, 9).Value

            ElseIf prevous_vol > Greatest And prevous_vol > volume Then

                Greatest = prevous_vol
                
                greatest_name = ws.Cells(last_k, 9).Value

            End If

    
    Next k
        
        
        
   '-------------------------------------------------------------
   

' get the names for ticker
   
   
    ws.Range("O2").Value = increase_name
    ws.Range("O3").Value = decrease_name
    ws.Range("O4").Value = greatest_name
    ws.Range("P2").Value = Increase
    ws.Range("P3").Value = Decrease
    ws.Range("P4").Value = Greatest
   
     
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"
    ws.Range("P4").NumberFormat = "000,000"
        

 
   
   
 '-----------coloers---------
 
 
 
 ' finds the end of the row for the cloume
 
    jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row

    
            For j = 2 To jEndRow
    
                'if greater than or less than zero
                If ws.Cells(j, 10) > 0 Then
    
                    ws.Cells(j, 10).Interior.ColorIndex = 4
    
                Else
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                    
                    
                End If
    
            Next j

   
        
        
        
        
'------------------ For " I " --END---------------------------------

Next ws


 Call LoopStockv2

    
End Sub

    
Sub LoopStockv2():

    MsgBox ("LoopStock Module has been ran")

End Sub

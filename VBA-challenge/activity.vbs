Sub Act1()


'Set the data range - determines the rows
LastRow = Range("A" & Rows.Count).End(xlUp).Row

'creates a variable of all the ocupied range for the A colum
Dim rangecells As String
rangecells = "A1:A" & LastRow

'Find the unique values into the range and creates an array
Set Data = Range(rangecells)
Dim unique As Variant
unique = WorksheetFunction.unique(Data)


'find the ammount of values into the array
Dim countunique As Integer
countunique = WorksheetFunction.CountA(unique)


'creates the range variable to print the tickers
Dim newtickers As String
newtickers = "H1:H" & countunique

'Creates the colum whit the unique tickers
Range(newtickers) = unique


'Creates a for cicle for each unique ticker value

For i = 2 To countunique
'countunique
currentticker = Cells(i, 8)


            'For loop that finds the range for currentticker
            Dim r As Range
            Dim initialprice As Double
            Dim finalprice As Double
            Dim totalstock As Variant
        
            For j = 2 To LastRow
             
                If Range("A" & j).Value = currentticker Then
                    If r Is Nothing Then
                        Set r = Range("B" & j)
                    Else
                        Set r = Union(r, Range("B" & j))
                    End If
                End If
            Next j
            
           
                            
                        'get the opening day
                    
                        Dim Firstdate As Date
                        Firstdate = WorksheetFunction.Min(r)
                        
                        
                          'get the closing day
                    
                        Dim Lastdate As Date
                        Lastdate = WorksheetFunction.Max(r)
                        
                        
                        'find the last row of a range
                        Dim lastrangerow As Long
                        lastrangerow = r.Rows(r.Rows.Count).Row
                        
                        
                        'sets the initial range value
                        Dim initialrangerow As Long
                         Dim templastrangerow As Long
                    
    
                    'sets the initial range
                       If initialrangerow = 0 Then
                        initialrangerow = 2
                       Else
                        initialrangerow = templastrangerow
                       End If
                       
                        
                    
                        
                        'loop to get the opening price at the first day considering inly the range for the ticker
                        For k = initialrangerow To lastrangerow
                        
                             'looks for the oldest date and the open value
                            If Range("B" & k).Value = Firstdate Then
                            initialprice = Cells(k, 3)
                        
                           Cells(i, 9) = initialprice
                            End If
                            
                            'looks for the oldest date and the open value
                            If Range("B" & k).Value = Lastdate Then
                            finalprice = Cells(k, 6)
                        
                        'calculates the difference in prices
                            Dim difference As Double
                            difference = finalprice - initialprice
                           Cells(i, 9) = difference
                           
                           
                          'calculates the percent difference in prices
                            Dim percent As Double
                            percent = (finalprice / initialprice) - 1
                           Cells(i, 10) = percent
                    
                            End If
                            
                           'Calculates the total volume
                            If Range("A" & k).Value = currentticker Then
                            totalstock = totalstock + Range("G" & k).Value
                            End If
                            Cells(i, 11) = totalstock
                            
                        
 
                            
                            Next k
                            
                        'clean up r to loop again
                        totalstock = 0
                            
                        templastrangerow = lastrangerow + 1
                        
                       'clean up r to loop again
                        Set r = Nothing
                        
                       


       

            
            




Next i



'set format color
 Dim format_range_color As Range
 Set format_range_color = Range("I2", "I" & LastRow)
 Dim condition1 As FormatCondition, condition2 As FormatCondition
 format_range_color.FormatConditions.Delete
 
 Set condition1 = format_range_color.FormatConditions.Add(xlCellValue, xlGreater, "=0")
 Set condition2 = format_range_color.FormatConditions.Add(xlCellValue, xlLess, "=0")


   With condition1
    .Interior.Color = RGB(35, 136, 35)
   End With

   With condition2
     .Interior.Color = RGB(210, 34, 45)
   End With
   
   'set format percent
 Dim format_range_percent As Range
 Set format_range_percent = Range("J2", "J" & LastRow)
 Dim condition3 As FormatCondition, condition4 As FormatCondition
 format_range_percent.FormatConditions.Delete
 
 Set condition3 = format_range_percent.FormatConditions.Add(xlCellValue, xlGreater, "=0")
 Set condition4 = format_range_percent.FormatConditions.Add(xlCellValue, xlLess, "=0")


   With condition3
    .Font.Color = RGB(35, 136, 35)
    .NumberFormat = "0.00%"
   End With

   With condition4
     .Font.Color = RGB(210, 34, 45)
     .NumberFormat = "0.00%"
   End With
    
    
       'set format percent 2
 Dim format_range_percent2 As Range
 Set format_range_percent2 = Range("P2", "P3")
 format_range_percent2.NumberFormat = "0.00%"

    
    'set format date
     Dim format_range_date As Range
    Set format_range_date = Range("B2", "B" & LastRow)
    Set condition5 = format_range_date.FormatConditions.Add(xlCellValue, xlGreater, "=0")
   
   With condition5
    .NumberFormat = "dd-mm-yyy"
   End With
   
   

   
   
'find the biggest increase
        Dim Increaserange As Range
        Set Increaserange = Range("J2", "J" & LastRow)
        Dim increase As Double
        increase = WorksheetFunction.Max(Increaserange)
        decrease = WorksheetFunction.Min(Increaserange)
        Range("P2") = increase
        Range("P3") = decrease

        
        For l = 2 To countunique
        
        If Cells(l, 10) = increase Then
        Range("O2") = Cells(l, 8).Value
        End If
        
        If Cells(l, 10) = decrease Then
        Range("O3") = Cells(l, 8).Value
        End If
        
        Next l
        
        
'Find the greatest total volume

        Dim greatestrange As Range
        Set greatestrange = Range("K2", "K" & countunique)
        Dim greatestvalue As Variant
        greatestvalue = WorksheetFunction.Max(greatestrange)
        Range("P4") = greatestvalue
        
        For m = 2 To countunique
        
        If Cells(m, 11) = greatestvalue Then
        Range("O4") = Cells(m, 8).Value
        End If
        
        Next m
        
        
'Creates colum titles
Range("H1") = "Ticker"
Range("I1") = "Quarterly change"
Range("J1") = "Percent Change"
Range("K1") = "Total Stock volume"
Range("O1") = "Ticker"
Range("P1") = "Value"
Range("N2") = "Greatest % increase"
Range("N3") = "Greatest % decreas"
Range("N4") = "Greatest total volume"



        
    
   
End Sub
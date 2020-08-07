Attribute VB_Name = "Stock"
Sub StockAnalysis()
'PARTE2 CORRE EL CODIGO EN TODAS LAS HOJAS

'LOOP THROUGH ALL SHEETS
For Each ws In Worksheets
    Dim WorksheetName As String
   
    

'PRINT HEADERS TITLES

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    

'DEFINE VARIABLES
    Dim Ticker As String 'Is this Ticker result table or ticker column search?
    Dim YearlyChange As Variant
    Dim PercentChange As Variant
    Dim TotalStockVolume As Variant
    Dim YearOpen As Variant
    Dim YearClose As Variant
    Dim SummaryTableRow As Variant
    Dim openv As Variant 'open value



'START THE COUNTER
    TotalStockVolume = 0
    openv = 2

'SET UP INTEGERS FOR LOOP

    SummaryTableRow = 2
    
'LOOP:KEEP TRACK OF EACH STOCK LOCATION IN THE TABLE

     LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
     WorksheetName = ws.Name
          
'LOOP THROUGH ALL TICKERS IN FIRST COLUMN TO LAST ROW
     For i = 2 To ws.UsedRange.Rows.Count
    

'CHECK IF WE ARE WITHIN THE SAME STOCK, IF NOT THEN...
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
             TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value 'Agrega todo lo anterior cuando llega a la ultima celda
           
'FIND ALL THE VALUES IN SPECIFIC LOCATION
                      
            Ticker = ws.Cells(i, 1).Value
        
            YearOpen = ws.Cells(openv, 3).Value
           
            YearClose = ws.Cells(i, 6).Value
           
                                                     
                      
'DEFINE OPERATIONS OR FUNCTIONS
            YearlyChange = (YearClose - YearOpen)
            
    
'IF ANUAL CHANGE DIFFERENT CERO -para que corra en la ultima hoja P,
'por que tiene ceros y sin el if diferente a cero no termina la iteracion
        
           If YearOpen <> 0 Then
           
                PercentChange = (YearlyChange / YearOpen)
           Else
                PercentChange = 0
            End If
            
                     
           
           
           'INSERT VALUES INTO SUMMARY
           'Se puede utilizar tambien la solucion con Range("I"&2+j)...
           ws.Cells(SummaryTableRow, 9).Value = Ticker
           
           ws.Cells(SummaryTableRow, 10).Value = YearlyChange
           
           'ws.Cells(SummaryTableRow, 11).Value = PercentChange & "%" o tambien
           
           ws.Cells(SummaryTableRow, 11).Value = PercentChange
           ws.Cells(SummaryTableRow, 11).Style = "Percent"
           ws.Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
           
           
           
           ws.Cells(SummaryTableRow, 12).Value = TotalStockVolume
    
            
            
            If YearlyChange > 0 Then
                ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
            End If
            
            
            'WRITES THE NEXT RESULT IN THE NEXT ROW
                SummaryTableRow = SummaryTableRow + 1 'Escribir siguiente resultado en la siguiente fila
                        
            'RESETING VOLUME AT EACH CHANGE OF STOCK
                TotalStockVolume = 0 'Reset a la sumatoria del ticker cada que cambia. de A a B x ejemplo
            
                openv = i + 1
            
            
            Else
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
               
             
    'FINISH IF FUNCTION
        End If
    
'FINISH LOOP
    Next i


'Max-Min SummaryTableRow Percent Change
    Dim MaxValue As Double
    Dim a As Integer
    Dim Tickername As String


    ws.Range("O1").Value = "Tickername"
    ws.Range("P1").Value = "MaxValue"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"

'INSERT GREATEST %INCREASE
    For a = 3 To SummaryTableRow
        If ws.Cells(a, 11).Value > MaxValue Then
            MaxValue = ws.Cells(a, 11).Value
            Tickername = ws.Cells(a, 9).Value
            
        Else

            ws.Range("O2") = Tickername
        
            ws.Range("P2") = MaxValue
            ws.Range("P2").Style = "Percent"
            ws.Range("P2").NumberFormat = "0.00%"
    
    End If
      


    Next a
    
'INSERT GREATEST %DECREASE

    For a = 3 To SummaryTableRow
        If ws.Cells(a, 11).Value < MaxValue Then
           MaxValue = ws.Cells(a, 11).Value
           Tickername = ws.Cells(a, 9).Value
            
            
        Else

            ws.Range("O3") = Tickername
        
            ws.Range("P3") = MaxValue
            ws.Range("P3").Style = "Percent"
            ws.Range("P3").NumberFormat = "0.00%"
        
    
    End If
      


    Next a

'INSERT GREATEST TOTAL VOLUME
    For a = 3 To SummaryTableRow
        If ws.Cells(a, 12).Value > MaxValue Then
            MaxValue = ws.Cells(a, 12).Value
            Tickername = ws.Cells(a, 9).Value
            
        Else

            ws.Range("O4") = Tickername
            
            ws.Range("P4") = MaxValue
            ws.Range("P4").NumberFormat = "0.00"
            
    
    End If
      


    Next a
    
'Move to the next worksheet
Next ws

End Sub




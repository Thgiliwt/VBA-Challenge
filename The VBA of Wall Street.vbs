Sub doitall()

    Dim xsh As Worksheet
    
    Application.ScreenUpdating = False
    
    
    For Each xsh In Worksheets
    
        xsh.Select
        
        Call stockanalysis
    
    Next
    
    Application.ScreenUpdating = True
    

End Sub




Sub stockanalysis()

    '==========Part 1 - Setting text in cells==========
    
    Dim titles As Variant
    
        Cells(1, "I").Value = "Ticker,Yearly Change,Percent Change,Total Stock Volume,,,,Ticker Name,Value"
        
        titles = Split(Cells(1, "I"), ",")
        
    Dim i As Integer
    
        For i = 0 To UBound(titles)
        
        Cells(1, i + 9).Value = titles(i)
        
        Cells(2, "O").Value = "Greatest % Increase"
        Cells(3, "O").Value = "Greatest % Decrease"
        Cells(4, "O").Value = "Greatest Totlal Volume"
        
        Next i
        
    '==========Part 2 - Determining values for Ticker, Yearly Change, Percent Change and Total Stock Volume & Formatting ==========
    
    Dim tickername As String
    
    Dim lastrow, rownumber, a As Long
    
    Dim openvalue, closevalue, changevalue As Double
    
    Dim volume As Variant
    
        lastrow = Cells(Rows.Count, "A").End(xlUp).Row
        
        rownumber = 2
        
        openvalue = Cells(2, "C").Value
        
        For a = 2 To lastrow
        
            If Cells(a, "A").Value <> Cells(a + 1, "A").Value Then
            
                tickername = Cells(a, "A").Value
                Cells(rownumber, "I").Value = tickername
                
                
                closevalue = Cells(a, "F").Value
                changevalue = closevalue - openvalue
                Cells(rownumber, "J").Value = changevalue
                
                If changevalue > 0 Then
                
                    Cells(rownumber, "J").Interior.Color = RGB(0, 255, 0)
                    
                ElseIf changevalue < 0 Then
                
                    Cells(rownumber, "J").Interior.Color = RGB(255, 0, 0)
                    
                Else
                
                    Cells(rownumber, "J").Interior.Color = RGB(0, 0, 255)
                
                End If
                
                If openvalue = 0 Then
                
                    Cells(rownumber, "K").Value = "Not Applicable"
                    
                Else
                                    
                    Cells(rownumber, "K").Value = FormatPercent(changevalue / openvalue)
                
                End If
                
                
                volume = volume + Cells(a, "G").Value
                Cells(rownumber, "L").Value = volume
                
            
            rownumber = rownumber + 1
            
            openvalue = Cells(a + 1, "C").Value
            
            volume = 0
            
            Else
            
                volume = volume + Cells(a, "G").Value
            
            End If
            
        Next a
        
    '==========Part 3 - Determining greatest % change for both postive and negative, greatest total stock volume and their corresponding name==========

    Dim deter_lastrow As Long
    
        deter_lastrow = Cells(Rows.Count, "I").End(xlUp).Row
        
    Dim percent() As Variant
    Dim total_volume() As Variant
    Dim x As Long
    
        x = deter_lastrow
    
    ReDim percent(1 To x)
    ReDim total_volume(1 To x)
            
    Dim b As Long
    
        For b = 1 To deter_lastrow
        
            percent(b) = Range("K" & b + 1)
            
            total_volume(b) = Range("L" & b + 1)
            
        Next b
        
        Cells(2, "Q").Value = FormatPercent(Application.WorksheetFunction.Max(percent))
        Cells(3, "Q").Value = FormatPercent(Application.WorksheetFunction.Min(percent))
        Cells(4, "Q").Value = Application.WorksheetFunction.Max(total_volume)
        
        Cells(2, "p").Value = Range("I" & Application.WorksheetFunction.Match(Range("Q2"), percent, 0) + 1)
        Cells(3, "p").Value = Range("I" & Application.WorksheetFunction.Match(Range("Q3"), percent, 0) + 1)
        Cells(4, "p").Value = Range("I" & Application.WorksheetFunction.Match(Range("Q4"), total_volume, 0) + 1)
   
    '==========Part 4 - Layouts==========
            
    ActiveSheet.UsedRange.EntireColumn.AutoFit
    
            
    
End Sub
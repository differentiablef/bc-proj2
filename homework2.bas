Attribute VB_Name = "Module1"


Sub CreateSummaries()
    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' variable declarations
    Dim inPos As Long, botPos As Long
    Dim volTotal As Double
    
    
    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' loop through available worksheets and create
    '   summary for each
    
    For Each Curr In ActiveWorkbook.Worksheets
    
        ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' setup summary section headers
    
        Curr.Cells(1, 9) = "Ticker"
        Curr.Cells(1, 10) = "Total Volume"
    
        ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' loop through data rows with non-nil first entry
        inPos = 2
        botPos = 2
        While Curr.Cells(inPos, 1).Text <> ""
            ' update summary information
            volTotal = Curr.Cells(inPos, 7) + volTotal
        
            ' does the next entry have the same symbol?
            If Curr.Cells(inPos, 1) <> Curr.Cells(inPos + 1, 1) Then
                ' if not, "push" ticker symbol summary onto the bottom of
                '  the list we are building in the worksheet
            
                Curr.Cells(botPos, 9) = Curr.Cells(inPos, 1)  ' ticker symbol
                Curr.Cells(botPos, 10) = volTotal        ' total volume
            
                botPos = 1 + botPos ' set new bottom position
            
                ' reset summary variables for next symbol
                volTotal = 0#
            End If
        
            inPos = 1 + inPos
        Wend
    Next Curr
End Sub

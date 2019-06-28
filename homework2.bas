Attribute VB_Name = "Module1"


Sub WorksheetCreateSummary()
    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' variable declarations
    Dim inPos As Integer, botPos As Integer
    Dim volTotal As Double
    
    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' setup summary section headers
    
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Total Volume"
    
    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' loop through data rows with non-nil first entry
    inPos = 2
    botPos = 2
    While Cells(inPos, 1).Text <> ""
        ' update summary information
        volTotal = Cells(inPos, 7) + volTotal
        
        ' does the next entry have the same symbol?
        If Cells(inPos, 1) <> Cells(inPos + 1, 1) Then
            ' if not, "push" ticker symbol summary onto the bottom of
            '  the list we are building in the worksheet
            
            Cells(botPos, 9) = Cells(inPos, 1)  ' ticker symbol
            Cells(botPos, 10) = volTotal        ' total volume
            
            botPos = 1 + botPos ' set new bottom position
            
            ' reset summary variables for next symbol
            volTotal = 0#
        End If
        
        inPos = 1 + inPos
    Wend
    

End Sub

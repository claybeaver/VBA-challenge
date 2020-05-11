Attribute VB_Name = "Module1"
Sub StockAnalysis()
    
    
    ' ws Decleeration
    
    Dim ws As Variant
    
    ' Loop through all worksheets
    
    For Each ws In Worksheets
    
    
    ' Set an initial variable for holding the Ticker name
    Dim Ticker_Name As String
    Ticker_Name = Blank
    
    ' Set an initial variable for holding the totals for each Ticker & Information
    Dim Ticker_Total As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Stock_Open As Double
    Dim Stock_Close As Double
    Dim Total_Stock_Volume As LongLong
    Yearly_Change = 0
    Stock_Open = ws.Cells(2, 3).Value
    Stock_Close = 0
    Percent_Change = 0
    Total_Stock_Volume = 0
    
    ' Final Row
    Dim FinalRow_Stocks As Long
    FinalRow_Stocks = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    
    ' Keep track of the location for each new Ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Format the Output Table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Format Percentage Change Column
    ws.Columns("K").NumberFormat = "0.00%"
    
    
    ' Loop through all trade data
    For i = 2 To FinalRow_Stocks
    
        ' Check if we have encountered a new ticker group
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ' Set the Ticker Name
            Ticker_Name = ws.Cells(i, 1).Value
            
            ' Make last addition to the Ticker groups Yearly Change, Percent Change, and Total Stock Volume
                
                ' Yearly Change & Percent Change
                Stock_Close = ws.Range("F" & i).Value
                Yearly_Change = Stock_Open - Stock_Close
                
                    'Conditional for zero
                    If Stock_Open <> 0 Then
                Percent_Change = ((Stock_Open - Stock_Close) / Stock_Open)
                Else: Percent_Change = 0
                End If
                    
                
                ' Total Stock Volume
                Total_Stock_Volume = ws.Range("G" & i).Value + Total_Stock_Volume
            
            ' Print the Ticker and corresponding information in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
            
            ' Print the Yearly Change in the Summary Table
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                ' Conditionally Format Green and Red for Positive and Negative values
                If ws.Range("J" & Summary_Table_Row).Value < 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                End If
            
            ' Print the Percent Change in the Summary Table
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            
            ' Print the Total Stock Volume in the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
            
            ' Add one to the Summary Table Row
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset Data
            Yearly_Change = 0
            Percent_Change = 0
            Total_Stock_Volume = 0
            
            ' Set next Ticker's stock open
            Stock_Open = ws.Cells(i + 1, 3).Value
            
        ' If the cell immediately following a row is the same ticker...
        Else
        
            ' Add to the Total Stock Volume total
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7)
            
        End If
        
    Next i

    '
    '
    '
    '
    '
    '
    '
    ' Challenge to return "Greatest % Increase", "Greatest % Decrease", and "Greatest total volume"

    ' Declare Variables
    Dim Greatest_Percent_Increase As Double
    Dim Greatest_Percent_Decrease As Double
    Dim Greatest_Total_Volume As Double
    Dim GPI_Holder As Double
    Dim GPD_Holder As Double
    Dim GTV_Holder As Double
    Greatest_Percent_Increase = ws.Range("K2").Value
    Greatest_Percent_Decrease = ws.Range("K2").Value
    Greatest_Total_Volume = ws.Range("L2").Value
    GPI_Holder = ws.Range("K2").Value
    GPD_Holder = ws.Range("K2").Value
    GTV_Holder = ws.Range("L2").Value
    Dim Ticker_Holder As Double
    Dim Ticker_Holder2 As Double
    Dim Ticker_Holder3 As Double
    
    
    
    ' Format the Output Table
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    
    ' Final Row Outputs
    Dim FinalRow_Outputs As Long
    FinalRow_Outputs = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    ' Solve using For Loop
    For j = 2 To FinalRow_Outputs
    
        'Check for Greatest % Increase
        If ws.Range("K" & j).Value > ws.Range("K" & j + 1).Value Then
        GPI_Holder = ws.Range("K" & j).Value
        End If
        If GPI_Holder > Greatest_Percent_Increase Then
        Greatest_Percent_Increase = GPI_Holder
        Ticket_Holder = ws.Range("I" & j).Value
        End If
        
        'Check for Greatest % Decrease
        If ws.Range("K" & j).Value < ws.Range("K" & j + 1).Value Then
        GPD_Holder = ws.Range("K" & j).Value
        End If
        If GPD_Holder < Greatest_Percent_Decrease Then
        Greatest_Percent_Decrease = GPD_Holder
        Ticket_Holder2 = ws.Range("I" & j).Value
        End If
        
        'Check for Greatest Total Volume
        If ws.Range("L" & j).Value > ws.Range("L" & j + 1).Value Then
        GTV_Holder = ws.Range("L" & j).Value
        End If
        If GTV_Holder > Greatest_Total_Volume Then
        Greatest_Total_Volume = GTV_Holder
        Ticket_Holder3 = ws.Range("I" & j).Value
        End If
        
    
    Next j
    
    ' Display Results
    ws.Range("P2").Value = Ticket_Holder
    ws.Range("P3").Value = Ticket_Holder2
    ws.Range("P4").Value = Ticket_Holder3
    ws.Range("Q2").Value = Greatest_Percent_Increase
    ws.Range("Q3").Value = Greatest_Percent_Decrease
    ws.Range("Q4").Value = Greatest_Total_Volume
    
    'Format Output cells
    ws.Range("Q2:Q3").NumberFormat = "0.00%"




Next ws

End Sub

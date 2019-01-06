VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub HOMEWORK_2_VBA()

For Each ws In Worksheets

    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Dim Ticker_Symbol As String

    Dim Total_Volume As LongLong
    Total_Volume = 0

    Dim Open_Price As Double
    Open_Price = ws.Range("C2").Value
    
    Dim Summary_Row As Integer
    Summary_Row = 2

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

        For i = 2 To LastRow
       
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker_Symbol = ws.Cells(i, 1).Value
                                
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                                
                ws.Range("I" & Summary_Row).Value = Ticker_Symbol
                                
                Dim Close_Price As Double
                Close_Price = ws.Cells(i, 6).Value
                
                ws.Range("J" & Summary_Row).Value = Close_Price - Open_Price
                ws.Range("J" & Summary_Row).NumberFormat = "0.000000000"
                
                    If (ws.Range("J" & Summary_Row).Value > 0) Then
                        ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
                        
                    Else
                        ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
                
                    End If
                
                If Open_Price > 0 Then
                    ws.Range("K" & Summary_Row).Value = (Close_Price - Open_Price) / Open_Price
                Else
                    ws.Range("K" & Summary_Row).Value = 0
                End If
                
                ws.Range("K" & Summary_Row).NumberFormat = "0.00%"
                
                ws.Range("L" & Summary_Row).Value = Total_Volume
                    Total_Volume = 0
                
                Summary_Row = Summary_Row + 1
                
                Open_Price = ws.Cells(i + 1, 3).Value
                
            Else
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
            End If
        Next i
        
    'Hard portion of homework
        
    Dim LastRow2 As Integer
    LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
        
    Dim Greatest_Increase As Double
    Greatest_Increase = ws.Range("K2").Value
    
    Dim Increase_Ticker As String
    Increase_Ticker = ws.Range("I2").Value
       
        For i = 2 To LastRow2
                        
            If ws.Cells(i, 11).Value > Greatest_Increase Then
                Greatest_Increase = ws.Cells(i, 11).Value
                Increase_Ticker = ws.Cells(i, 9).Value
            End If
        
        Next i
        
    Dim Greatest_Decrease As Double
    Greatest_Decrease = ws.Range("K2").Value
        
    Dim Decrease_Ticker As String
    Decrease_Ticker = ws.Range("I2").Value
        
        For i = 2 To LastRow2
                        
            If ws.Cells(i, 11).Value < Greatest_Decrease Then
                Greatest_Decrease = ws.Cells(i, 11).Value
                Decrease_Ticker = ws.Cells(i, 9).Value
            End If
        
        Next i
        
    Dim Greatest_Volume As LongLong
    Greatest_Volume = ws.Range("K3").Value
    
    Dim Volume_Ticker As String
    Volume_Ticker = ws.Range("I2").Value
    
        For i = 2 To LastRow2
                        
            If ws.Cells(i, 12).Value > Greatest_Volume Then
                Greatest_Volume = ws.Cells(i, 12).Value
                Volume_Ticker = ws.Cells(i, 9).Value
            End If
        
        Next i
        
    ws.Range("P2").Value = Greatest_Increase
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("O2").Value = Increase_Ticker
        
    ws.Range("P3").Value = Greatest_Decrease
    ws.Range("P3").NumberFormat = "0.00%"
    ws.Range("O3").Value = Decrease_Ticker
        
    ws.Range("P4").Value = Greatest_Volume
    ws.Range("O4").Value = Volume_Ticker
            
Next ws
      
End Sub

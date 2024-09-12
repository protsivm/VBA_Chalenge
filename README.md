# VBA_Chalenge

Sub TickerPriceVolume()

    ' Define Variables
    Dim i As Long
    Dim j As Long
    Dim LastRow As Long
    Dim ws As Worksheet
    Dim Ticker As String
    Dim QTRopen As Double
    Dim QTRclose As Double
    Dim VolumeDay As Double
    Dim VolumeQTR As Double
    Dim MaxPosChange As Double
    Dim MaxNegChange As Double
    Dim MaxVolume As Double
    
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        
        ' Find the last row in column A for the current worksheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        'Create column and cell labels
        ws.Range("I1").Value = ("Ticker")
        ws.Range("J1").Value = ("Quaterly Change")
        ws.Range("K1").Value = ("Percent Change")
        ws.Range("L1").Value = ("Total Stock Volume")
        ws.Range("P1").Value = ("Ticker")
        ws.Range("Q1").Value = ("Value")
        ws.Range("O2").Value = ("Greatest % Increase")
        ws.Range("O3").Value = ("Greatest % Decrease")
        ws.Range("O4").Value = ("Greatest Total Volume")

        ' Initialize row counter for output
        j = 2
        
        'Set Values
        MaxPosChange = 0
        MaxNegChange = 0
        MaxVolume = 0
    
        ' Loop through each row, starting from row 2
        For i = 2 To LastRow
        
            ' If a new ticker is encountered (current is different from the previous)
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' Capture ticker, open QTR price, and initialize volume counter
                Ticker = ws.Cells(i, 1).Value
                QTRopen = ws.Cells(i, 3).Value
                VolumeQTR = ws.Cells(i, 7).Value
            End If
        
            If ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value Then
                ' Accumulate volume for the same ticker
                VolumeQTR = VolumeQTR + ws.Cells(i, 7).Value
            End If

            ' If we reach the last instance of the ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ' Capture the closing QTR Price
                QTRclose = ws.Cells(i, 6).Value
            
                ' Output the results to new columns
                ws.Cells(j, 9).Value = Ticker
                ws.Cells(j, 10).Value = QTRclose - QTRopen
                ws.Cells(j, 11).Value = (QTRclose - QTRopen) / QTRopen
                ws.Cells(j, 11).NumberFormat = "0.00%"
                ws.Cells(j, 12).Value = VolumeQTR
                
                    'Change Cell Interior Color
                    If QTRclose > QTRopen Then
                        ws.Cells(j, 10).Interior.ColorIndex = 4
                    ElseIf QTRclose < QTRopen Then
                        ws.Cells(j, 10).Interior.ColorIndex = 3
                    End If
                
                
                    ' Capture the Max Change
                    If ws.Cells(j, 11).Value > MaxPosChange Then
                        MaxPosChange = ws.Cells(j, 11).Value
                        ws.Range("P2").Value = ws.Cells(j, 9).Value
                        ws.Range("Q2").Value = MaxPosChange
                        ws.Range("Q2").NumberFormat = "0.00%"
                    End If
                
                    If ws.Cells(j, 11).Value < MaxNegChange Then
                        MaxNegChange = ws.Cells(j, 11).Value
                        ws.Range("P3").Value = ws.Cells(j, 9).Value
                        ws.Range("Q3").Value = MaxNegChange
                        ws.Range("Q3").NumberFormat = "0.00%"
                    End If
                
                    ' Capture the Max Volume
                    If ws.Cells(j, 12).Value > MaxVolume Then
                        MaxVolume = ws.Cells(j, 12).Value
                        ws.Range("Q4").Value = MaxVolume
                        ws.Range("P4").Value = ws.Cells(j, 9).Value
                    End If
                    
                ' Move to the next output row
                j = j + 1
                
            End If
                
          
               
                             
        Next i
       
        
    Next ws
                
End Sub

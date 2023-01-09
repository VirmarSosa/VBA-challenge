Sub VBAChallenge()
Dim x As Integer
Dim z As Long
Dim j As Integer

x = 251
Cells(1, 10).Value = "Ticker"
Cells(1, 11).Value = "Open Price"
Cells(1, 12).Value = "Close Price"
Cells(1, 13).Value = "Yearly Change"
Cells(1, 14).Value = "Percentage Change"
Cells(1, 15).Value = "Total Stock Volume"
Cells(1, 18) = "Ticker"
Cells(1, 19) = "Value"
Cells(2, 17) = "Greatest % Increase"
Cells(3, 17) = "Greatest % Decrease"
Cells(4, 17) = "Greatest Total Volume"

For z = 1 To 3050

'Ticker

Cells(2, 10).Value = Cells(2, 1).Value
Cells(2 + z, 10).Value = Cells(2 + (x * z), 1).Value

'Open

Cells(2, 11).Value = Cells(2, 3).Value
Cells(2 + z, 11).Value = Cells(2 + (x * z), 3).Value

'Close


            Cells(2, 12).Value = Cells(2 + (x - 1), 6).Value
            Cells(2 + z, 12).Value = Cells(2 + (x * z) + 250, 6).Value


'Yearly Change

Cells(2, 13).Value = Cells(2, 12).Value - Cells(2, 11).Value
Cells(2 + z, 13).Value = Cells(2 + z, 12).Value - Cells(2 + z, 11).Value

'Percent Change

    If Cells(2 + z, 11).Value = "" Then
    
            Cells(2 + z, 11).Value = ""
        
        Else
    
            Cells(2, 14).Value = ((((Cells(2, 12).Value * 100) / Cells(2, 11).Value) - 100) / 100)
            Cells(2 + z, 14).Value = ((((Cells(2 + z, 12).Value * 100) / Cells(2 + z, 11).Value) - 100) / 100)

    End If


'Total Stock

Cells(2, 15).Value = "=Sum(G2:G252)"
Cells(2 + z, 15).Value = Application.Sum(Range("G" & (2 + (x * z)), "G" & ((x * z) + 252)))

'Conditional Formatting

If Cells(2, 14).Value >= 0 Then
    
    Cells(2, 14).Interior.Color = vbGreen
   
Else: Cells(2, 14).Interior.Color = vbRed

End If

If Cells(2, 14).Value >= 0 Then

    Cells(2, 14).Interior.Color = vbGreen
   
    Else: Cells(2, 14).Interior.Color = vbRed

End If

If Cells(2 + z, 14).Value >= 0 Then

    Cells(2 + z, 14).Interior.Color = vbGreen
   
Else: Cells(2 + z, 14).Interior.Color = vbRed
      
End If

' % Max Increase

Cells(2, 19).Value = WorksheetFunction.Max(Range("N2", "N" & (2 + z)))

' % Min Decrease

Cells(3, 19).Value = WorksheetFunction.Min(Range("N2", "N" & (2 + z)))

' % Max Volumen

Cells(4, 19).Value = WorksheetFunction.Max(Range("O2", "O" & (2 + z)))

'Format

Range("Q2:Q4").Font.Size = 12
Range("Q2:Q4").Font.Bold = True
Range("Q2:Q4").VerticalAlignment = xlCenter
Range("Q2:Q4").EntireColumn.AutoFit
Range("N2", "N" & (2 + z)).NumberFormat = "0.00%"
Range("O2", "O" & (2 + z)).NumberFormat = "0"
Range("S2:S3").NumberFormat = "0.00%"
Range("J1:S1").Font.Size = 12
Range("J1:S1").Font.Bold = True
Range("J1:S1").VerticalAlignment = xlCenter
Range("J1:S1").EntireColumn.AutoFit
'Range("K:L").EntireColumn.Hidden = True

Next z

For j = 2 To 3500
    
    'Greatest Increase Ticker

    If Cells(j, 14).Value = Cells(2, 19).Value Then
        
        Cells(2, 18).Value = Cells(j, 10).Value

    End If
    
    'Greatest Decrease Ticker
    
    If Cells(j, 14).Value = Cells(3, 19).Value Then
        
        Cells(3, 18).Value = Cells(j, 10).Value

    End If
    
    'Greatest Volume Ticker
    
    If Cells(j, 15).Value = Cells(4, 19).Value Then
        
        Cells(4, 18).Value = Cells(j, 10).Value

    End If
    
Next j

End Sub

Sub Stock_Breakdown()
    'define worksheet to process
    Dim ws As Worksheet
    
    
    For Each ws In ActiveWorkbook.Worksheets
        'set headers for each new column and pretty up spacing in all worksheets
        ws.Columns("I:Q").AutoFit
        ws.[I1] = "Ticker"
        ws.[J1] = "Yearly Change"
        ws.[K1] = "Percent Change"
        ws.[L1] = "Total Stock Volume"
        ws.[O2] = "Greatest % Increase"
        ws.[O3] = "Greatest % Decrease"
        ws.[O4] = "Greatest Total Volume"
        ws.[P1] = "Ticker"
        ws.[Q1] = "Value"

        'define last row (code contribution from class 10/31/2023)
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'set solution index (for when year/ticker resolves)
        si = 2
        'define firstOpen
        firstOpen = 0
      
        For i = 2 To lastrow
            if firstOpen = 0 Then
                firstOpen = ws.Cells(i, "C")

            End If

            'pull ticker symbols from column A and note them in column I
            If ws.Cells(i, "A") <> ws.Cells(i + 1, "A") Then
                ws.Cells(si, "I") = ws.Cells(i, "A")
                'define and calculate yearly change to populate column J (code contribution/suggestion from Geronimo Perez via tutoring session 11/08/2023. Mr. Perez values brevity in code and it made things much more readable this way)
                yearlyCh = ws.Cells(i, "F") - firstOpen
                ws.Cells(si, "J") = yearlyCh

                    'add conditional formatting to yearly change column
                     if yearlyCh > 0 Then
                        ws.Cells(si, "J").interior.colorIndex =4
                     else 
                         ws.Cells(si, "J").interior.colorIndex =3
                    end if
                
                'calculate percent yearly change
                percentCh = yearlyCh / firstOpen * 100
                ws.Cells(si, "K") = percentCh

                'where the total stock volume goes after it's finished with its else conditional
                ws.Cells(si, "L") = totalCh

                si = si + 1
                firstOpen = 0

            'total yearly stock volume 
            else
                totalCh = totalCh + ws.Cells(i, "G")
            End If

            
        Next i
    Next ws
End Sub


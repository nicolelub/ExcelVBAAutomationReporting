Sub FinalSort()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim hasColorF_O As Boolean
    Dim hasColorP_Y As Boolean
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Sheets
        Application.ScreenUpdating = False
        
        ' Find last row with data
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Make sure we have data to sort
        If lastRow > 1 Then
            ' Add helper columns Z, AA, and AB
            ws.Range("Z1").Value = "ColorPriority"  ' Lower number = higher priority
            ws.Range("AA1").Value = "F-O Color"     ' Color priority for F-O columns
            ws.Range("AB1").Value = "P-Y Color"     ' Color priority for P-Y columns
            
            ' Fill helper columns based on cell colors and locations
            For i = 2 To lastRow
                ' Check if F-O has any colored cells
                hasColorF_O = False
                For j = 6 To 15 ' Columns F to O
                    If ws.Cells(i, j).Interior.ColorIndex <> xlNone Then
                        hasColorF_O = True
                        Exit For
                    End If
                Next j
                
                ' Check if P-Y has any colored cells
                hasColorP_Y = False
                For j = 16 To 25 ' Columns P to Y
                    If ws.Cells(i, j).Interior.ColorIndex <> xlNone Then
                        hasColorP_Y = True
                        Exit For
                    End If
                Next j
                
                ' Set ColorPriority
                If hasColorF_O Then
                    ws.Cells(i, "Z").Value = 1 ' Highest priority - F-O colored
                ElseIf hasColorP_Y Then
                    ws.Cells(i, "Z").Value = 2 ' Second priority - P-Y colored
                Else
                    ws.Cells(i, "Z").Value = 3 ' Lowest priority - no colors
                End If
                
                ' Color priority for F-O columns (Columns F-O)
                Dim foColor As Long
                foColor = xlNone
                For j = 6 To 15
                    If ws.Cells(i, j).Interior.ColorIndex <> xlNone Then
                        foColor = ws.Cells(i, j).Interior.Color
                        Exit For
                    End If
                Next j
                
                ' Set F-O color priority
                Select Case foColor
                    Case RGB(255, 0, 0)     ' Red
                        ws.Cells(i, "AA").Value = 1
                    Case RGB(255, 255, 0)   ' Yellow
                        ws.Cells(i, "AA").Value = 2
                    Case RGB(255, 165, 0)   ' Orange
                        ws.Cells(i, "AA").Value = 3
                    Case Else
                        ws.Cells(i, "AA").Value = 4
                End Select
                
                ' Color priority for P-Y columns (Columns P-Y)
                Dim pyColor As Long
                pyColor = xlNone
                For j = 16 To 25
                    If ws.Cells(i, j).Interior.ColorIndex <> xlNone Then
                        pyColor = ws.Cells(i, j).Interior.Color
                        Exit For
                    End If
                Next j
                
                ' Set P-Y color priority
                Select Case pyColor
                    Case RGB(255, 0, 0)     ' Red
                        ws.Cells(i, "AB").Value = 1
                    Case RGB(255, 255, 0)   ' Yellow
                        ws.Cells(i, "AB").Value = 2
                    Case RGB(255, 165, 0)   ' Orange
                        ws.Cells(i, "AB").Value = 3
                    Case Else
                        ws.Cells(i, "AB").Value = 4
                End Select
            Next i
            
            ' Sort using the helper columns
            With ws.Sort
                .SortFields.Clear
                ' First sort by color location group (F-O first, then P-Y, then none)
                .SortFields.Add Key:=ws.Range("Z2:Z" & lastRow), Order:=xlAscending
                ' Then by F-O column color (Red, Yellow, Orange)
                .SortFields.Add Key:=ws.Range("AA2:AA" & lastRow), Order:=xlAscending
                ' Then by P-Y column color (Red, Yellow, Orange)
                .SortFields.Add Key:=ws.Range("AB2:AB" & lastRow), Order:=xlAscending
                .SetRange ws.Range("A1:AB" & lastRow)
                .Header = xlYes
                .Apply
            End With
            
            ' Clean up helper columns
            ws.Range("Z1:AB" & lastRow).ClearContents
        End If
        
        Application.ScreenUpdating = True
    Next ws
    
    MsgBox "Sorting complete for all sheets!", vbInformation
End Sub

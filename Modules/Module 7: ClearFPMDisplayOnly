Sub ClearFPMDisplayOnly()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim colPlanNameF As Integer, colPlanNameP As Integer
    
    ' Set worksheet
    Set ws = ThisWorkbook.Sheets("Results")
    
    ' Define column indices
    colPlanNameF = 6  ' Column F (Plan Name in F-O range)
    colPlanNameP = 16 ' Column P (Plan Name in P-Y range)
    
    ' Find last row based on column F (assuming F always has data)
    lastRow = ws.Cells(ws.Rows.Count, colPlanNameF).End(xlUp).Row
    
    ' Loop through the rows
    For Each cell In ws.Range(ws.Cells(2, colPlanNameF), ws.Cells(lastRow, colPlanNameF))
        If Trim(LCase(cell.Value)) = "fpm display only" Then
            ' Clear contents in F-O (columns 6 to 15)
            ws.Range(ws.Cells(cell.Row, 6), ws.Cells(cell.Row, 15)).ClearContents
            ' Clear contents in P-Y (columns 16 to 25)
            ws.Range(ws.Cells(cell.Row, 16), ws.Cells(cell.Row, 25)).ClearContents
        End If
    Next cell
    
    ' Loop through the rows again for column P
    For Each cell In ws.Range(ws.Cells(2, colPlanNameP), ws.Cells(lastRow, colPlanNameP))
        If Trim(LCase(cell.Value)) = "fpm display only" Then
            ' Clear contents in P-Y (columns 16 to 25)
            ws.Range(ws.Cells(cell.Row, 16), ws.Cells(cell.Row, 25)).ClearContents
        End If
    Next cell
    
    MsgBox "Clearing complete!", vbInformation
End Sub



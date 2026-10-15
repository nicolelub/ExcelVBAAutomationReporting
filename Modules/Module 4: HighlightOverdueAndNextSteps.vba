Sub HighlightOverdueAndNextSteps()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim todayDate As Date
    Dim overdueDate As Date
    Dim nextStepDate As Date
    
    ' Define the fixed column ranges to highlight
    Const COL_Z As Integer = 26  ' Overdue Step Expected Date (Z)
    Const COL_AA As Integer = 27 ' Overdue Step Summary (AA)
    Const COL_AB As Integer = 28 ' Next Step Expected Date (AB)
    Const COL_AC As Integer = 29 ' Next Step Summary (AC)
    
    ' Set worksheet and today's date
    Set ws = ThisWorkbook.Sheets("Results")
    todayDate = Date
    
    ' Find last row with data (using a reliable column)
    lastRow = ws.Cells(ws.Rows.Count, COL_Z).End(xlUp).Row
    
    ' Clear existing formatting only in our target columns (Z-AC)
    ws.Range(ws.Cells(2, COL_Z), ws.Cells(lastRow, COL_AC)).Interior.ColorIndex = xlNone
    
    ' Process each row for Overdue Step Expected Date
    For Each cell In ws.Range(ws.Cells(2, COL_Z), ws.Cells(lastRow, COL_Z))
        ' Check if cell contains a valid date
        If Not IsEmpty(cell.Value) And IsDate(cell.Value) Then
            ' Convert date string to date
            overdueDate = CDate(cell.Value)
            
            ' Highlight in red if date is in the past
            If overdueDate < todayDate Then
                ' Highlight both Overdue Step Expected Date and Overdue Step Summary
                ws.Range(ws.Cells(cell.Row, COL_Z), ws.Cells(cell.Row, COL_AA)).Interior.ColorIndex = 3 ' Red
            End If
        End If
    Next cell
    
    ' Process each row for Next Step Expected Date
    For Each cell In ws.Range(ws.Cells(2, COL_AB), ws.Cells(lastRow, COL_AB))
        ' Check if cell contains a valid date
        If Not IsEmpty(cell.Value) And IsDate(cell.Value) Then
            ' Convert date string to date
            nextStepDate = CDate(cell.Value)
            
            ' Highlight in red if date is in the past
            If nextStepDate < todayDate Then
                ' Highlight both Next Step Expected Date and Next Step Summary
                ws.Range(ws.Cells(cell.Row, COL_AB), ws.Cells(cell.Row, COL_AC)).Interior.ColorIndex = 3 ' Red
            End If
        End If
    Next cell
    
    MsgBox "Overdue steps highlighting complete!", vbInformation
End Sub

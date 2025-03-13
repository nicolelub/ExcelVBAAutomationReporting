Sub DeleteUnmatchedRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim keepRow As Boolean
    Dim checkColumns As Variant
    Dim checkNames As Variant
    Dim colValue As String
    Dim colIndex As Variant
    Dim nameToCheck As Variant
    
    ' Set worksheet (Name of first tab - change to your fit)
    Set ws = ThisWorkbook.Sheets("Results")
    
    ' Define the columns to check (C, G, H, Q, R)
    checkColumns = Array(3, 7, 8, 17, 18) ' Column numbers
    
    ' Define the names to look for (change based on your last names - my names are anonymized/changed)
    checkNames = Array("Tenn", "Smith", "Monroe", "Garel", "Hally", "Donovan", "Power", "Brown")
    
    ' Find the last row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Application.ScreenUpdating = False
    
    ' Counter for rows deleted
    Dim rowsDeleted As Long
    rowsDeleted = 0
    
    ' Loop from bottom to top (to avoid skipping rows when deleting)
    For i = lastRow To 2 Step -1
        keepRow = False ' Assume row should be deleted
        
        ' Check each specified column
        For Each colIndex In checkColumns
            ' Get cell value
            colValue = CStr(ws.Cells(i, colIndex).Value)
            
            ' Check for each name
            For Each nameToCheck In checkNames
                If InStr(1, colValue, CStr(nameToCheck), vbTextCompare) > 0 Then
                    keepRow = True ' Name found, keep row
                    Exit For
                End If
            Next nameToCheck
            
            If keepRow Then Exit For
        Next colIndex
        
        ' Delete row if no match found
        If Not keepRow Then
            ws.Rows(i).Delete
            rowsDeleted = rowsDeleted + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox rowsDeleted & " rows without specified names have been removed!", vbInformation
End Sub


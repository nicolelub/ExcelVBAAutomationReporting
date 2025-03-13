Sub CreatePMTabsAndDistributeData()
    Dim ws As Worksheet
    Dim wsNew As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim headerRow As Range
    Dim pmNames As Variant
    Dim pmTabs As Object
    Dim wsExists As Boolean
    Dim targetRow As Long
    
    ' Define manager column indices
    Const COL_C As Integer = 3  ' Prospect Manager
    Const COL_G As Integer = 7  ' Plan Manager
    Const COL_H As Integer = 8  ' Secondary Plan Manager
    Const COL_Q As Integer = 17 ' 2 Plan Manager
    Const COL_R As Integer = 18 ' 2 Secondary Plan Manager
    
    ' Define array of column indices to check
    Dim managerColumns As Variant
    managerColumns = Array(COL_C, COL_G, COL_H, COL_Q, COL_R)
    
    ' Define the names to look for
    pmNames = Array("Tenn", "Smith", "Monroe", "Garel", "Hally", "Donovan", "Power", "Brown")
    
    ' Set source worksheet
    Set ws = ThisWorkbook.Sheets("Results")
    
    ' Determine last row and column with data
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Define header row
    Set headerRow = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    
    ' Create a dictionary to track created worksheets
    Set pmTabs = CreateObject("Scripting.Dictionary")
    
    ' Sort the data by Column C (Prospect Manager) and then by Column G (Plan Manager)
    ws.Range("A1:" & ws.Cells(lastRow, lastCol).Address).Sort _
        Key1:=ws.Range("C2"), Order1:=xlAscending, _
        Key2:=ws.Range("G2"), Order2:=xlAscending, _
        Header:=xlYes
    
    ' First, create tabs for each PM in the list
    For j = LBound(pmNames) To UBound(pmNames)
        ' Check if worksheet with this name already exists
        wsExists = False
        For Each wsNew In ThisWorkbook.Worksheets
            If wsNew.Name = CStr(pmNames(j)) Then
                wsExists = True
                Exit For
            End If
        Next wsNew
        
        ' Create new worksheet if it doesn't exist
        If Not wsExists Then
            Set wsNew = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            wsNew.Name = CStr(pmNames(j))
        Else
            ' If sheet exists, clear its contents except for headers
            Set wsNew = ThisWorkbook.Worksheets(CStr(pmNames(j)))
            wsNew.Cells.Clear
        End If
        
        ' Copy header row
        headerRow.Copy wsNew.Range("A1")
        
        ' Store reference to the sheet in dictionary
        If Not pmTabs.Exists(CStr(pmNames(j))) Then
            pmTabs.Add CStr(pmNames(j)), wsNew
        End If
    Next j
    
    ' Dictionary to track which rows have been copied to which tabs
    Dim rowCopied As Object
    Set rowCopied = CreateObject("Scripting.Dictionary")
    
    ' Process each row and check all manager columns
    For i = 2 To lastRow
        ' Reset the dictionary for each row
        Set rowCopied = CreateObject("Scripting.Dictionary")
        
        ' Check each manager column
        For Each colIndex In managerColumns
            Dim managerName As String
            managerName = Trim(ws.Cells(i, colIndex).Value)
            
            ' Find if this manager is in our PM list
            For j = LBound(pmNames) To UBound(pmNames)
                If InStr(1, managerName, CStr(pmNames(j)), vbTextCompare) > 0 Then
                    Dim pmTab As String
                    pmTab = CStr(pmNames(j))
                    
                    ' Check if we've already copied this row to this tab
                    If Not rowCopied.Exists(pmTab) Then
                        ' Copy row to the appropriate tab
                        Set wsNew = pmTabs(pmTab)
                        targetRow = wsNew.Cells(wsNew.Rows.Count, "A").End(xlUp).Row + 1
                        ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Copy wsNew.Range("A" & targetRow)
                        
                        ' Mark as copied to this tab
                        rowCopied.Add pmTab, True
                    End If
                    
                    Exit For
                End If
            Next j
        Next
    Next i
    
    ' Format all PM tabs
    For j = LBound(pmNames) To UBound(pmNames)
        If pmTabs.Exists(CStr(pmNames(j))) Then
            Set wsNew = pmTabs(CStr(pmNames(j)))
            wsNew.Columns.AutoFit
        End If
    Next j
    
    MsgBox "Data has been sorted and distributed to PM tabs!", vbInformation
End Sub

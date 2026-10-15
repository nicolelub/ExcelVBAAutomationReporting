Sub ConsolidatedSortAndHighlightPlan2()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim todayDate As Date
    Dim firstOfMonth As Date
    Dim lastOfMonth As Date
    Dim askDate As Date
    Dim responseDate As Date
    Dim planStatus As String
    Dim dateCell As Range
    Dim fiscalYear As Integer
    Dim currentFY As Integer
    Dim planName As String
    Dim headerRow As Range
    
    ' Define the fixed column ranges to highlight (P to Y)
    Const COL_P As Integer = 16 ' Column P
    Const COL_Y As Integer = 25 ' Column Y
    
    ' Define allowed plan managers
    Dim checkNames As Variant
    checkNames = Array("Tenn", "Smith", "Monroe", "Garel", "Hally", "Donovan", "Power", "Brown")
    
    ' Set worksheet and dates
    Set ws = ThisWorkbook.Sheets("Results")
    todayDate = Date
    firstOfMonth = DateSerial(Year(todayDate), Month(todayDate), 1)
    lastOfMonth = DateSerial(Year(todayDate), Month(todayDate) + 1, 0)
    
    ' Determine current fiscal year (FY starts on July 1)
    If Month(todayDate) >= 7 Then
        currentFY = Year(todayDate) + 1
    Else
        currentFY = Year(todayDate)
    End If
    
    ' Fiscal year start date
    Dim fiscalYearStart As Date
    fiscalYearStart = DateSerial(currentFY - 1, 7, 1)
    
    ' Find header row and corresponding columns
    Set headerRow = ws.Rows(1)
    
    ' Locate key columns by header names
    Dim colPlanName As Integer
    Dim colPlanStatus As Integer
    Dim colExpectedAskDate As Integer
    Dim colResponseDate As Integer
    Dim colPlanManager As Integer
    Dim colSecondaryPlanManager As Integer
    
    colPlanName = 0
    colPlanStatus = 0
    colExpectedAskDate = 0
    colResponseDate = 0
    colPlanManager = 0
    colSecondaryPlanManager = 0
    
    ' Find column indices by header names
    Dim i As Integer
    For i = 1 To headerRow.Cells.Count
        If Not IsEmpty(headerRow.Cells(1, i).Value) Then
            Select Case Trim(headerRow.Cells(1, i).Value)
                Case "2 Plan Name"
                    colPlanName = i
                Case "2 Plan Status"
                    colPlanStatus = i
                Case "2 Plan Expected Ask Date"
                    colExpectedAskDate = i
                Case "2 Plan Resonse Date"
                    colResponseDate = i
                Case "2 Plan Manager"
                    colPlanManager = i
                Case "2 Secondary Plan Manager"
                    colSecondaryPlanManager = i
            End Select
        End If
    Next i
    
    ' Validate required columns were found
    If colPlanStatus = 0 Then
        MsgBox "Could not find '2 Plan Status' column. Please check the header row.", vbExclamation
        Exit Sub
    End If
    
    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.Count, colPlanStatus).End(xlUp).Row
    
    ' Convert text dates to actual dates
    If colExpectedAskDate > 0 Then
        For Each dateCell In ws.Range(ws.Cells(2, colExpectedAskDate), ws.Cells(lastRow, colExpectedAskDate))
            If Not IsEmpty(dateCell) And IsDate(dateCell.Value) Then
                dateCell.Value = CDate(dateCell.Value)
                dateCell.NumberFormat = "mm/dd/yyyy"
            End If
        Next dateCell
    End If
    
    If colResponseDate > 0 Then
        For Each dateCell In ws.Range(ws.Cells(2, colResponseDate), ws.Cells(lastRow, colResponseDate))
            If Not IsEmpty(dateCell) And IsDate(dateCell.Value) Then
                dateCell.Value = CDate(dateCell.Value)
                dateCell.NumberFormat = "mm/dd/yyyy"
            End If
        Next dateCell
    End If
    
    ' Sort by Plan Status and Response Date
    With ws.Sort
        .SortFields.Clear
        If colPlanStatus > 0 Then
            .SortFields.Add2 Key:=ws.Range(ws.Cells(2, colPlanStatus), ws.Cells(lastRow, colPlanStatus)), Order:=xlAscending
        End If
        If colResponseDate > 0 Then
            .SortFields.Add2 Key:=ws.Range(ws.Cells(2, colResponseDate), ws.Cells(lastRow, colResponseDate)), Order:=xlAscending
        End If
        .SetRange ws.Range("A1:AZ" & lastRow) ' Covers all possible columns
        .Header = xlYes
        .Apply
    End With
    
    ' Clear existing formatting in columns P-Y
    ws.Range(ws.Cells(2, COL_P), ws.Cells(lastRow, COL_Y)).Interior.ColorIndex = xlNone
    
    ' Apply color formatting based on new rules
    If colPlanStatus > 0 Then
        For Each cell In ws.Range(ws.Cells(2, colPlanStatus), ws.Cells(lastRow, colPlanStatus))
            Dim targetRange As Range
            Set targetRange = ws.Range(ws.Cells(cell.Row, COL_P), ws.Cells(cell.Row, COL_Y))
            
            ' Get plan status and name
            planStatus = LCase(Trim(cell.Value))
            planName = LCase(ws.Cells(cell.Row, colPlanName).Value)
            
            ' Get dates safely
            If colExpectedAskDate > 0 And Not IsEmpty(ws.Cells(cell.Row, colExpectedAskDate).Value) And IsDate(ws.Cells(cell.Row, colExpectedAskDate).Value) Then
                askDate = CDate(ws.Cells(cell.Row, colExpectedAskDate).Value)
            Else
                askDate = 0
            End If
            
            If colResponseDate > 0 And Not IsEmpty(ws.Cells(cell.Row, colResponseDate).Value) And IsDate(ws.Cells(cell.Row, colResponseDate).Value) Then
                responseDate = CDate(ws.Cells(cell.Row, colResponseDate).Value)
            Else
                responseDate = 0
            End If
            
            ' Check for allowed plan managers
            Dim isAllowedManager As Boolean
            isAllowedManager = False
            If colPlanManager > 0 Then
                Dim planManager As String
                planManager = LCase(Trim(ws.Cells(cell.Row, colPlanManager).Value))
                
                For i = LBound(checkNames) To UBound(checkNames)
                    If InStr(LCase(planManager), LCase(checkNames(i))) > 0 Then
                        isAllowedManager = True
                        Exit For
                    End If
                Next i
            End If
            
            If colSecondaryPlanManager > 0 And Not isAllowedManager Then
                Dim secondaryPlanManager As String
                secondaryPlanManager = LCase(Trim(ws.Cells(cell.Row, colSecondaryPlanManager).Value))
                
                For i = LBound(checkNames) To UBound(checkNames)
                    If InStr(LCase(secondaryPlanManager), LCase(checkNames(i))) > 0 Then
                        isAllowedManager = True
                        Exit For
                    End If
                Next i
            End If
            
            ' Highlighting rules
            If Not isAllowedManager Then
                GoTo NextRow
            End If
            
            If (planStatus = "accepted" Or planStatus = "rejected" Or planStatus = "canceled") And askDate < fiscalYearStart Then
                targetRange.Interior.ColorIndex = 45 ' Orange
            ElseIf askDate <= todayDate And (planStatus = "qualified" Or planStatus = "unqualified") Then
                targetRange.Interior.ColorIndex = 3 ' Red
            ElseIf askDate >= firstOfMonth And askDate <= lastOfMonth And planStatus = "qualified" Then
                targetRange.Interior.ColorIndex = 6 ' Yellow
            ElseIf planStatus = "response pending" And responseDate <= todayDate Then
                targetRange.Interior.ColorIndex = 3 ' Red
            ElseIf responseDate >= firstOfMonth And responseDate <= lastOfMonth Then
                targetRange.Interior.ColorIndex = 6 ' Yellow
            End If
            
NextRow:
        Next cell
    End If
    
    MsgBox "Sorting and highlighting complete!", vbInformation
End Sub


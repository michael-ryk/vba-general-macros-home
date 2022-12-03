Sub EmphasizeSimilar()
    ' ==========================================================================
    ' Description: Emphasize rows with similar tags and gray out all remaining
    ' Excel where used: My code write repeat
    ' ==========================================================================

    Debug.Print ("================ Start =================")
    Application.ScreenUpdating = False
    
    'Constants
    Const startingRow As Integer = 7
    Const tagColumn As String = "H"
    Const filterColumn = "I"
    Const subjectColumn = "D"
    Const colorStartColumn As String = "A"
    Const colorEndColumn As String = "I"
    Const boldStartColumn = "D"
    Const boldEndColumn = "E"
    Const previousSelectedRowCellAddress = "F2"
    
    'Declare variables
    Dim currentRow As Integer
    Dim lastRow As Long
    Dim tagList As String
    Dim tagArray() As String
    Dim flagTagMatch As Boolean
    Dim flagSubjectMatch As Boolean
    Dim currentSubject As String
    Dim previousSelectedSubject As String
    Dim i As Long
    
    'Validate selected row in valid range
    currentRow = ActiveCell.Row
    If (currentRow < startingRow) Then Exit Sub
    
    lastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    tagList = Cells(currentRow, tagColumn)
    tagArray = Split(tagList, " ")
    previousSelectedSubject = ActiveSheet.Range(previousSelectedRowCellAddress).Value
    currentSubject = ActiveSheet.Cells(currentRow, subjectColumn).Value
    'Debug.Print ("Current selected row: " & currentRow)
    Debug.Print ("tag list from current row: " & tagList)
    Debug.Print ("Current selected subject: " & currentSubject)
    Debug.Print ("Previous selected subject: " & previousSelectedSubject)
    ActiveSheet.Range(previousSelectedRowCellAddress).Value = currentSubject
    
    'Set bold and colors to default for all rows
    ActiveSheet.Range(Cells(startingRow, boldStartColumn), Cells(lastRow, colorEndColumn)).Font.Bold = False
    ActiveSheet.Range(Cells(startingRow, colorStartColumn), Cells(lastRow, colorEndColumn)).Font.Color = RGB(56, 56, 56)
    
    'Color Active line
    ActiveSheet.Range(Cells(currentRow, colorStartColumn), Cells(currentRow, colorEndColumn)).Font.Color = RGB(48, 84, 150)
    
    '=== Main Loop ===
    For i = startingRow To lastRow
        
        flagTagMatch = False
        flagSubjectMatch = False
        
        For Each Item In tagArray
            'Mark row which have one tag which included in selected row
            If InStr(1, Cells(i, tagColumn).Value, Item) Then
                flagTagMatch = True
            End If
            'Mark row which have at least one keyword from tag section in subject
            If InStr(1, Cells(i, "F").Value, Item) Then
                flagSubjectMatch = True
            End If
        Next Item
        
        'Set row filter result value for future sorting
        If (flagTagMatch) Then
            ActiveSheet.Range(Cells(i, boldStartColumn), Cells(i, boldEndColumn)).Font.Bold = True
            ActiveSheet.Cells(i, filterColumn).Value = "2"
        ElseIf (flagSubjectMatch) Then
            ActiveSheet.Cells(i, filterColumn).Value = "3"
            ActiveSheet.Range(Cells(i, colorStartColumn), Cells(i, colorEndColumn)).Font.Color = RGB(128, 128, 128)
        Else
            ActiveSheet.Cells(i, filterColumn).Value = "4"
            ActiveSheet.Range(Cells(i, colorStartColumn), Cells(i, colorEndColumn)).Font.Color = RGB(217, 217, 217)
        End If
        
        'Selected row = 1 to make it first after sorting
        If (i = currentRow) Then
            ActiveSheet.Cells(i, filterColumn).Value = "1"
        End If
        
        'Color previous row
        If (ActiveSheet.Cells(i, subjectColumn) = previousSelectedSubject) Then
            ActiveSheet.Range(Cells(i, colorStartColumn), Cells(i, colorEndColumn)).Font.Color = RGB(142, 169, 219)
        End If
        
    Next i
    
    'Filter relevant match - Think if i want it
    'ActiveSheet.ListObjects("Concepts").Range.AutoFilter Field:=11, Criteria1:="1"
    
    'Sort Data
    ActiveSheet.ListObjects("Concepts").Sort.SortFields.Clear
    ActiveSheet.ListObjects("Concepts").Sort.SortFields. _
        Add2 Key:=Range("Concepts[[#All],[Filter]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.ListObjects("Concepts").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Restore initial settings
    Application.ScreenUpdating = True

End Sub
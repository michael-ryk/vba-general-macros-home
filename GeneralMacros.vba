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
    Const colorStartColumn As String = "A"
    Const colorEndColumn As String = "I"
    Const boldStartColumn = "D"
    Const boldEndColumn = "E"
    
    'Declare variables
    Dim currentRow As Integer
    Dim lastRow As Long
    Dim tagList As String
    Dim tagArray() As String
    Dim flagTagMatch As Boolean
    Dim flagSubjectMatch As Boolean
    
    'Get Current row and decide if exit
    currentRow = ActiveCell.Row
    If (currentRow < startingRow) Then Exit Sub
    
    lastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    tagList = Cells(currentRow, tagColumn)
    tagArray = Split(tagList, " ")
    
    ActiveSheet.Range(Cells(startingRow, boldStartColumn), Cells(lastRow, colorEndColumn)).Font.Bold = False
    ActiveSheet.Range(Cells(startingRow, colorStartColumn), Cells(lastRow, colorEndColumn)).Font.Color = RGB(56, 56, 56)
    
    Debug.Print ("Current selected row: " & currentRow)
    Debug.Print ("tag list from current row: " & tagList)
    
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
            Cells(i, filterColumn).Value = "1"
        ElseIf (flagSubjectMatch) Then
            Cells(i, filterColumn).Value = "2"
            ActiveSheet.Range(Cells(i, colorStartColumn), Cells(i, colorEndColumn)).Font.Color = RGB(128, 128, 128)
        Else
            Cells(i, filterColumn).Value = "3"
            ActiveSheet.Range(Cells(i, colorStartColumn), Cells(i, colorEndColumn)).Font.Color = RGB(217, 217, 217)
        End If
        
    Next i
    
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
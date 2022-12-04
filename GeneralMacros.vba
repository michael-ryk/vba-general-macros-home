Sub EmphasizeSimilar()
    ' ==========================================================================
    ' Version: v1.0
    ' Description: Emphasize rows with similar tags and gray out all remaining
    ' Excel where used: My code write repeat
    ' ==========================================================================

    Debug.Print ("================ Start =================")
    Application.ScreenUpdating = False
    
    'Constants
    Const tableName As String = "Knowledge"
    Const startingRow As Integer = 7
    Const tagColumn As String = "H"
    Const filterColumn = "I"
    Const subjectColumn = "D"
    Const lockColumn = "J"
    Const dateColumn = "K"
    Const quantityColumn = "L"
    Const previousSelectedRowCellAddress = "D2"
    Const colorStartColumn As String = "A"
    Const colorEndColumn As String = "J"
    Const boldStartColumn = "D"
    Const boldEndColumn = "E"
    
    'Declare variables
    Dim currentRow As Integer
    Dim lastRow As Long
    Dim tagList As String
    Dim tagArray() As String
    Dim flagTagMatch As Boolean
    Dim flagSubjectMatch As Boolean
    Dim currentSubject As String
    Dim previousSelectedSubject As String
    Dim todayDate As Date
    Dim counter As Integer
    Dim i As Long
    
    'Validate selected row in valid range
    currentRow = ActiveCell.row
    lastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).row
    If (currentRow < startingRow) Then
        ActiveSheet.Range(Cells(startingRow, boldStartColumn), Cells(lastRow, colorEndColumn)).Font.Bold = False
        ActiveSheet.Range(Cells(startingRow, colorStartColumn), Cells(lastRow, colorEndColumn)).Font.Color = RGB(56, 56, 56)
        Exit Sub
    End If
    
    tagList = Cells(currentRow, tagColumn)
    tagArray = Split(tagList, " ")
    previousSelectedSubject = ActiveSheet.Range(previousSelectedRowCellAddress).value
    currentSubject = ActiveSheet.Cells(currentRow, subjectColumn).value
    todayDate = Date
    counter = 0
    'Debug.Print ("Current selected row: " & currentRow)
    Debug.Print ("tag list from current row: " & tagList)
    Debug.Print ("Current selected subject: " & currentSubject)
    Debug.Print ("Previous selected subject: " & previousSelectedSubject)
    ActiveSheet.Range(previousSelectedRowCellAddress).value = currentSubject
    
    'Set bold and colors to default for all rows
    ActiveSheet.Range(Cells(startingRow, boldStartColumn), Cells(lastRow, colorEndColumn)).Font.Bold = False
    ActiveSheet.Range(Cells(startingRow, colorStartColumn), Cells(lastRow, colorEndColumn)).Font.Color = RGB(56, 56, 56)
    
    '=== Main Loop ===
    For i = startingRow To lastRow
        
        flagTagMatch = False
        flagSubjectMatch = False
        
        For Each Item In tagArray
            'Mark row which have one tag which included in selected row
            If InStr(1, Cells(i, tagColumn).value, Item) Then
                flagTagMatch = True
            End If
            'Mark row which have at least one keyword from tag section in subject
            If InStr(1, Cells(i, subjectColumn).value, Item) Then
                flagSubjectMatch = True
            End If
        Next Item
        
        'Set row filter result value for future sorting
        If (flagTagMatch) Then
            'Tags matched in tags cell - color black + bold
            ActiveSheet.Range(Cells(i, boldStartColumn), Cells(i, boldEndColumn)).Font.Bold = True
            ActiveSheet.Cells(i, filterColumn).value = "2"
            counter = counter + 1
        ElseIf (flagSubjectMatch) Then
            'tags included subject cell - color grey
            ActiveSheet.Cells(i, filterColumn).value = "3"
            ActiveSheet.Range(Cells(i, colorStartColumn), Cells(i, colorEndColumn)).Font.Color = RGB(128, 128, 128)
        Else
            'All remained rows - very light grey
            ActiveSheet.Cells(i, filterColumn).value = "4"
            ActiveSheet.Range(Cells(i, colorStartColumn), Cells(i, colorEndColumn)).Font.Color = RGB(217, 217, 217)
        End If
        
        'Set lock rows before active row + color green
        If (ActiveSheet.Cells(i, lockColumn).value = "yes") Then
            ActiveSheet.Cells(i, filterColumn).value = "0"
            ActiveSheet.Range(Cells(i, colorStartColumn), Cells(i, colorEndColumn)).Font.Color = RGB(0, 176, 80)
        End If
        
        'Color previous row - light blue
        If (ActiveSheet.Cells(i, subjectColumn) = previousSelectedSubject) Then
            ActiveSheet.Range(Cells(i, colorStartColumn), Cells(i, colorEndColumn)).Font.Color = RGB(142, 169, 219)
        End If
        
        'Selected row = 1 to make it first after sorting + color Dark blue + update date
        If (i = currentRow) Then
            ActiveSheet.Cells(i, filterColumn).value = "1"
            ActiveSheet.Cells(i, dateColumn).value = todayDate
            ActiveSheet.Range(Cells(currentRow, colorStartColumn), Cells(currentRow, colorEndColumn)).Font.Color = RGB(48, 84, 150)
        End If
        
    Next i
    
    'Filter relevant match - Think if i want it
    'ActiveSheet.ListObjects("Concepts").Range.AutoFilter Field:=11, Criteria1:="1"
    
    'Save quantity of connections to current selected row
    ActiveSheet.Cells(currentRow, quantityColumn).value = counter
    
    'Sort Data
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Clear
    ActiveSheet.ListObjects(tableName).Sort.SortFields. _
        Add2 Key:=Range(tableName & "[[#All],[Filter]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.ListObjects(tableName).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Restore initial settings
    Application.ScreenUpdating = True

End Sub
Sub EmphasizeSimilar()
    ' ==========================================================================
    ' Version: v2.0
    ' Description: Emphasize rows with similar tags and gray out all remaining
    ' Excel where used: My code write repeat
    ' ==========================================================================

    Debug.Print ("================ Start =================")
    Application.ScreenUpdating = False
    
    'Constants
    Const subjectColumn = "D"
    Const SavedAsideSubjectCellAddress = "D2"
    Const SavedAsideTagsCellAddress = "D3"
    Const SavedAsideLocationCellAddress = "D4"
    Const colorStartColumn As String = "A"
    Const colorEndColumn As String = "J"
    Const boldStartColumn = "D"
    Const boldEndColumn = "E"
    
    'Declare variables
    Dim startingRow As Integer
    Dim currentRow As Integer
    Dim lastRow As Long
    Dim locationCulmn As String
    Dim filterColumn As Integer
    Dim lockColumn As Integer
    Dim dateColumn As Integer
    Dim connectiosColumn As Integer
    Dim tagColumn As Integer
    Dim tagList As String
    Dim selectedRowTagArray() As String
    Dim targetRowTagArray() As String
    Dim flagTagMatch As Boolean
    Dim flagSubjectMatch As Boolean
    Dim currentSubject As String
    Dim previousSelectedSubject As String
    Dim todayDate As Date
    Dim counter As Integer
    Dim tableName As String
    Dim i As Long
    
    'Assign variables based on current excel file
    tableName = ActiveSheet.ListObjects(1).Name
    startingRow = ActiveSheet.ListObjects(1).Range.Cells(1, 1).Row + 1
    filterColumn = ActiveSheet.ListObjects(1).ListColumns("Filter").Range.Column
    lockColumn = ActiveSheet.ListObjects(1).ListColumns("Lock").Range.Column
    dateColumn = ActiveSheet.ListObjects(1).ListColumns("Date").Range.Column
    connectionsColumn = ActiveSheet.ListObjects(1).ListColumns("Connections").Range.Column
    tagColumn = ActiveSheet.ListObjects(1).ListColumns("Tags").Range.Column
    locationColumn = ActiveSheet.ListObjects(1).ListColumns("Location").Range.Column
    
    'Validate selected row in valid range
    currentRow = ActiveCell.Row
    lastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    If (currentRow < startingRow) Then
        ActiveSheet.Range(Cells(startingRow, boldStartColumn), Cells(lastRow, colorEndColumn)).Font.Bold = False
        ActiveSheet.Range(Cells(startingRow, colorStartColumn), Cells(lastRow, colorEndColumn)).Font.Color = RGB(56, 56, 56)
        Exit Sub
    End If
    
    'Clear filter if applied
    On Error Resume Next
    ActiveSheet.ShowAllData
    
    tagList = Cells(currentRow, tagColumn)
    selectedRowTagArray = Split(tagList, " ")
    previousSelectedSubject = ActiveSheet.Range(SavedAsideSubjectCellAddress).Value
    currentSubject = ActiveSheet.Cells(currentRow, subjectColumn).Value
    todayDate = Date
    counter = 0
    'Debug.Print ("Current selected row: " & currentRow)
    Debug.Print ("tag list from current row: " & tagList)
    Debug.Print ("Current selected subject: " & currentSubject)
    Debug.Print ("Previous selected subject: " & previousSelectedSubject)
    Debug.Print (ActiveSheet.ListObjects(1).ListColumns("Filter").Range.Column)
    ActiveSheet.Range(SavedAsideSubjectCellAddress).Value = currentSubject
    ActiveSheet.Range(SavedAsideTagsCellAddress).Value = ActiveSheet.Cells(currentRow, tagColumn).Value
    ActiveSheet.Range(SavedAsideLocationCellAddress).Value = ActiveSheet.Cells(currentRow, locationColumn).Value
    
    
    'Set bold and colors to default for all rows
    ActiveSheet.Range(Cells(startingRow, boldStartColumn), Cells(lastRow, colorEndColumn)).Font.Bold = False
    ActiveSheet.Range(Cells(startingRow, colorStartColumn), Cells(lastRow, colorEndColumn)).Font.Color = RGB(56, 56, 56)
    
    '=== Main Loop ===
    For i = startingRow To lastRow
        
        flagTagMatch = False
        flagSubjectMatch = False
        targetRowTagArray = Split(Cells(i, tagColumn), " ")
        Debug.Print ("Row: " & i)
        
        For Each selectedTag In selectedRowTagArray
            
            'Mark row which have one tag which included in selected row
            For Each targetTag In targetRowTagArray
                If (selectedTag = targetTag) Then flagTagMatch = True
            Next targetTag
            
            'Mark row which have at least one keyword from tag section in subject
            If InStr(1, Cells(i, subjectColumn).Value, selectedTag) Then
                flagSubjectMatch = True
            End If
            
        Next selectedTag
        
        'Set row filter result value for future sorting
        If (flagTagMatch) Then
            'Tags matched in tags cell - color black + bold
            ActiveSheet.Range(Cells(i, boldStartColumn), Cells(i, boldEndColumn)).Font.Bold = True
            ActiveSheet.Cells(i, filterColumn).Value = "2"
            counter = counter + 1
        ElseIf (flagSubjectMatch) Then
            'tags included subject cell - color grey
            ActiveSheet.Cells(i, filterColumn).Value = "3"
            ActiveSheet.Range(Cells(i, colorStartColumn), Cells(i, colorEndColumn)).Font.Color = RGB(128, 128, 128)
        Else
            'All remained rows - very light grey
            ActiveSheet.Cells(i, filterColumn).Value = "4"
            ActiveSheet.Range(Cells(i, colorStartColumn), Cells(i, colorEndColumn)).Font.Color = RGB(217, 217, 217)
        End If
        
        'Set lock rows before active row + color green
        If (ActiveSheet.Cells(i, lockColumn).Value = "yes") Then
            ActiveSheet.Cells(i, filterColumn).Value = "0"
            ActiveSheet.Range(Cells(i, colorStartColumn), Cells(i, colorEndColumn)).Font.Color = RGB(0, 176, 80)
        End If
        
        'Color previous row - light blue
        If (ActiveSheet.Cells(i, subjectColumn) = previousSelectedSubject) Then
            ActiveSheet.Range(Cells(i, colorStartColumn), Cells(i, colorEndColumn)).Font.Color = RGB(142, 169, 219)
        End If
        
        'Selected row = 1 to make it first after sorting + color Dark blue + update date
        If (i = currentRow) Then
            ActiveSheet.Cells(i, filterColumn).Value = "1"
            ActiveSheet.Cells(i, dateColumn).Value = todayDate
            ActiveSheet.Range(Cells(currentRow, colorStartColumn), Cells(currentRow, colorEndColumn)).Font.Color = RGB(48, 84, 150)
        End If
        
    Next i
    
    'Filter relevant match - Think if i want it
    'ActiveSheet.ListObjects("Concepts").Range.AutoFilter Field:=11, Criteria1:="1"
    
    'Save quantity of connections to current selected row
    ActiveSheet.Cells(currentRow, connectionsColumn).Value = counter
    
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
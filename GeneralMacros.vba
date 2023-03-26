Option Explicit

Sub EmphasizeSimilar()
    ' ==========================================================================
    ' Version: v4.0
    ' Description: Emphasize rows with similar tags and gray out all remaining
    ' Excel where used: My code write repeat
    ' ==========================================================================

    Debug.Print ("================ Start =================")
    Application.ScreenUpdating = False
    
    '==================================================
    'Declare variables
    '==================================================
    
    'Constants
    Const addrSavedSubject = "D2"
    Const addrSavedTags = "D3"
    Const addrSavedLocation = "D4"
    Const colorStartColumn          As String = "A"
    Const colorEndColumn            As String = "J"
    
    '==================================================
    'Assign variables
    '==================================================
    
    Dim wbMain                      As Workbook
    Dim shtMain                     As Worksheet
    Dim lo                          As ListObject
    
    Set wbMain = ThisWorkbook
    Set shtMain = ActiveSheet
    Set lo = shtMain.ListObjects(1)
    
    Dim iColFilter                  As Integer
    Dim iColLock                    As Integer
    Dim iColDate                    As Integer
    Dim iColConnections             As Integer
    Dim iColTags                    As Integer
    Dim iColLocation                As Integer
    Dim iColSubject                 As Integer
    
    'Get Columns letters based on found headings in table
    iColFilter = lo.ListColumns("Filter").Range.Column
    iColLock = lo.ListColumns("Lock").Range.Column
    iColDate = lo.ListColumns("Date").Range.Column
    iColConnections = lo.ListColumns("Connections").Range.Column
    iColTags = lo.ListColumns("Tags").Range.Column
    iColLocation = lo.ListColumns("Location").Range.Column
    iColSubject = lo.ListColumns("Subject").Range.Column
    
    'todo - validate all columns exist in excel
    
    'Clear filter in case it was alredy applied
    On Error Resume Next
    lo.AutoFilter.ShowAllData
    
    'Get First and last row index
    Dim iFirstTableRow              As Integer
    Dim lRowLastInTable             As Long
    Dim sSelectedRow                As Long
    iFirstTableRow = lo.Range.Cells(1, 1).Row + 1
    sSelectedRow = ActiveCell.Row
    lRowLastInTable = shtMain.Range("A" & Rows.Count).End(xlUp).Row
    
    'Validate selected row in valid range
    If (sSelectedRow < iFirstTableRow) Then
        Call ClearFocus(iFirstTableRow, lRowLastInTable, iColSubject, iColTags)
        Exit Sub
    End If
        
    'Set Variables for currently selected row
    Dim arrSelectedTagList()        As String
    Dim sSelectedTagList            As String
    Dim sSelectedSubject            As String
    Dim sPreviousSubject            As String
    Dim todayDate                   As Date
    Dim iNumberOfConnections        As Integer
    sSelectedTagList = Cells(sSelectedRow, iColTags)
    arrSelectedTagList = Split(sSelectedTagList, " ")
    sSelectedSubject = shtMain.Cells(sSelectedRow, iColSubject).Value
    todayDate = Date
    iNumberOfConnections = 0
    
    'Save aside current selected row details for next run
    Dim rngSavedSubject             As Range
    Dim rngSavedTagsList            As Range
    Dim rngSavedLocation            As Range
    Set rngSavedSubject = shtMain.Range(addrSavedSubject)
    Set rngSavedTagsList = shtMain.Range(addrSavedTags)
    Set rngSavedLocation = shtMain.Range(addrSavedLocation)
    sPreviousSubject = rngSavedSubject.Value
    rngSavedSubject.Value = sSelectedSubject
    rngSavedTagsList.Value = shtMain.Cells(sSelectedRow, iColTags).Value
    rngSavedLocation.Value = shtMain.Cells(sSelectedRow, iColLocation).Value
    
    'Debug Prints
    Debug.Print ("tag list from current row: " & sSelectedTagList)
    Debug.Print ("Current selected subject: " & sSelectedSubject)
    Debug.Print ("Previous selected subject: " & sPreviousSubject)
    
    'Set default style for all rows
    Dim rngStyleApply               As Range
    Set rngStyleApply = shtMain.Range(Cells(iFirstTableRow, colorStartColumn), Cells(lRowLastInTable, colorEndColumn))
    With rngStyleApply.Font
        .Bold = False
        .color = RGB(56, 56, 56)
    End With

    '==================================================
    'Cycle through lines
    '==================================================
    Dim lRowIndex                   As Long
    Dim sRowTagList                 As String
    Dim bTagMatch                   As Boolean
    Dim bSubjectMatch               As Boolean
    Dim arrRowTagList()             As String
    
    For lRowIndex = iFirstTableRow To lRowLastInTable
        
        sRowTagList = Cells(lRowIndex, iColTags)
        
        If (Len(sRowTagList)) Then
            
            bTagMatch = False
            bSubjectMatch = False
            arrRowTagList = Split(sRowTagList, " ")
            
            '==================================================
            'Cycle through tags from selected row
            '==================================================
            Dim selectedTag As Variant
            For Each selectedTag In arrSelectedTagList
                
                'Mark row which have one tag which included in selected row
                If Not (bTagMatch) Then
                    Dim targetTag As Variant
                    For Each targetTag In arrRowTagList
                        If (selectedTag = targetTag) Then
                            bTagMatch = True
                            Exit For
                        End If
                    Next targetTag
                End If
                
                'Mark row which have at least one keyword from tag section in subject
                If InStr(1, Cells(lRowIndex, iColSubject).Value, selectedTag) Then
                    bSubjectMatch = True
                End If
                
            Next selectedTag
            
            '==================================================
            'Set row filter result value for future sorting
            '==================================================
            
            If (bTagMatch) Then
                'Tags matched in tags cell - color black + bold
                shtMain.Range(Cells(lRowIndex, iColSubject), Cells(lRowIndex, iColSubject)).Font.Bold = True
                shtMain.Cells(lRowIndex, iColFilter).Value = "Match"
                iNumberOfConnections = iNumberOfConnections + 1
            ElseIf (bSubjectMatch) Then
                'tags included subject cell - color grey
                shtMain.Cells(lRowIndex, iColFilter).Value = "Sugest"
                Call colorRow(lRowIndex, colorStartColumn, colorEndColumn, RGB(128, 128, 128))
            Else
                'All remained rows - very light grey
                shtMain.Cells(lRowIndex, iColFilter).Value = "Others"
                Call colorRow(lRowIndex, colorStartColumn, colorEndColumn, RGB(190, 190, 190))
            End If
            
            'Lock rows have highest priority of sorting above current row + color green
            If (shtMain.Cells(lRowIndex, iColLock).Value = "yes") Then
                shtMain.Cells(lRowIndex, iColFilter).Value = "Lock"
                Call colorRow(lRowIndex, colorStartColumn, colorEndColumn, RGB(0, 176, 80))
            End If
            
            'Color previous row - light blue
            If (shtMain.Cells(lRowIndex, iColSubject) = sPreviousSubject) Then
                Call colorRow(lRowIndex, colorStartColumn, colorEndColumn, RGB(142, 169, 219))
            End If
            
            'Selected row = 1 to make it before results + color Dark blue + update date
            If (lRowIndex = sSelectedRow) Then
                shtMain.Cells(lRowIndex, iColFilter).Value = "Main"
                shtMain.Cells(lRowIndex, iColDate).Value = todayDate
                Call colorRow(lRowIndex, colorStartColumn, colorEndColumn, RGB(48, 84, 150))
            End If
            
        End If
    
    Next lRowIndex
    
    '==================================================
    'Final configs
    '==================================================
    
    'Save quantity of connections to current selected row
    shtMain.Cells(sSelectedRow, iColConnections).Value = iNumberOfConnections
    
    'Filter all matches and blank lines
    lo.Range.AutoFilter Field:=iColFilter, Operator:=xlFilterValues, _
        Criteria1:=Array("", "Main", "Match", "Sugest", "Lock")
    'TO DO - Make field dynamic if this column in different place
        
    'Restore initial settings
    Application.ScreenUpdating = True

End Sub

Function colorRow(Row As Long, startCol As String, endCol As String, rgbColor As Long)
    ActiveSheet.Range(Cells(Row, startCol), Cells(Row, endCol)).Font.color = rgbColor
End Function

Sub ClearFocus(startRow As Integer, endRow As Long, iColSubject As Integer, endColumn As Integer)
'Clear filter and make all rows with same style
    ActiveSheet.Range(Cells(startRow, iColSubject), Cells(endRow, iColSubject)).Font.Bold = False
    ActiveSheet.Range(Cells(startRow, "A"), Cells(endRow, endColumn)).Font.color = RGB(56, 56, 56)
End Sub
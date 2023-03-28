Option Explicit

Sub EmphasizeSimilar()
    ' ==========================================================================
    ' Version: v4.0
    ' Description: Emphasize rows with similar tags and gray out all remaining
    ' Excel where used: My code write repeat
    ' ==========================================================================

    Debug.Print ("================ Start =================")
    Application.ScreenUpdating = False
    Dim StartTime                   As Double
    Dim SecondsElapsed              As Double
    StartTime = Timer
    
    'Constants - address of cells that are free in excel
    Const addrSavedSubject          As String = "D2"
    Const addrSavedTags             As String = "D3"
    Const addrSavedLocation         As String = "D4"
    Const addrColorColStart         As String = "F1"
    Const addrColorColEnd           As String = "F2"
    
    Dim wbMain                      As Workbook
    Dim shtMain                     As Worksheet
    Dim lo                          As ListObject
    
    Set wbMain = ThisWorkbook
    Set shtMain = ActiveSheet
    On Error Resume Next
    Set lo = shtMain.ListObjects(1)
    
    ' Test if Table exist in active sheet
    If lo Is Nothing Then
        MsgBox "Current sheet doesn't have any table", vbExclamation
        Exit Sub
    End If
    
    Dim iColFilter                  As Integer
    Dim iColLock                    As Integer
    Dim iColDate                    As Integer
    Dim iColConnections             As Integer
    Dim iColTags                    As Integer
    Dim iColLocation                As Integer
    Dim iColSubject                 As Integer
    Dim iColFoundTag                As Integer
    
    'Get Columns indexes based on found headings in table - All headings must present
    iColFilter = lo.ListColumns("Filter").Range.Column
    iColLock = lo.ListColumns("Lock").Range.Column
    iColDate = lo.ListColumns("Date").Range.Column
    iColConnections = lo.ListColumns("Connections").Range.Column
    iColTags = lo.ListColumns("Tags").Range.Column
    iColLocation = lo.ListColumns("Location").Range.Column
    iColSubject = lo.ListColumns("Subject").Range.Column
    iColFoundTag = lo.ListColumns("Found Tag").Range.Column
    
    'todo - validate all columns exist in excel
    
    'Clear autofilter in case it was alredy applied
    On Error Resume Next
    lo.AutoFilter.ShowAllData
        
    'Clear all contents of Filter column
    lo.ListColumns("Filter").DataBodyRange.ClearContents
    lo.ListColumns("Found Tag").DataBodyRange.ClearContents
        
    'Get First and last row index
    Dim iFirstTableRow              As Integer
    Dim lRowLastInTable             As Long
    Dim sSelectedRow                As Long
    iFirstTableRow = lo.Range.Cells(1, 1).Row + 1
    sSelectedRow = ActiveCell.Row
    lRowLastInTable = shtMain.Range("A" & Rows.Count).End(xlUp).Row
    
    'Set default style for all rows
    Dim colorStartColumn            As String
    Dim colorEndColumn              As String
    Dim rngStyleApply               As Range
    colorStartColumn = shtMain.Range(addrColorColStart).Value
    colorEndColumn = shtMain.Range(addrColorColEnd).Value
    Set rngStyleApply = shtMain.Range(Cells(iFirstTableRow, colorStartColumn), Cells(lRowLastInTable, colorEndColumn))
    With rngStyleApply.Font
        .Bold = False
        .Color = RGB(190, 190, 190)
    End With
    
    ' If selected row outside of table - stop macro - Keep unfiltered
    If (sSelectedRow < iFirstTableRow) Then
        rngStyleApply.Font.Color = RGB(56, 56, 56)
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
    'Debug.Print ("tag list from current row: " & sSelectedTagList)
    'Debug.Print ("Current selected subject: " & sSelectedSubject)
    'Debug.Print ("Previous selected subject: " & sPreviousSubject)

    '==================================================
    'Cycle through tags from selected row
    '==================================================
    Dim rngFilter           As Range
    Dim rngBold             As Range
    Dim rngLock             As Range
    Dim rngSubject          As Range
    Dim rngColorApply       As Range
    Dim rngTags             As Range
    Dim rngFoundTag         As Range
    Dim selectedTag         As Variant
    
    For Each selectedTag In arrSelectedTagList
    
        Debug.Print (selectedTag)
        
        '==================================================
        'Cycle through excel rows
        '==================================================
        Dim lRowIndex                   As Long
        Dim sRowTagList                 As String
        Dim sRowSubject                 As String
    
        For lRowIndex = iFirstTableRow To lRowLastInTable
        
            Set rngFilter = shtMain.Cells(lRowIndex, iColFilter)
            Set rngBold = shtMain.Range(Cells(lRowIndex, iColSubject), Cells(lRowIndex, iColSubject))
            Set rngLock = shtMain.Cells(lRowIndex, iColLock)
            Set rngSubject = shtMain.Cells(lRowIndex, iColSubject)
            Set rngColorApply = shtMain.Range(Cells(lRowIndex, colorStartColumn), Cells(lRowIndex, colorEndColumn))
            Set rngTags = shtMain.Cells(lRowIndex, iColTags)
            Set rngFoundTag = shtMain.Cells(lRowIndex, iColFoundTag)
            
            sRowTagList = rngTags.Value
            sRowSubject = rngSubject.Value
            
            ' Do only if tag cell not empty
            If (Len(sRowTagList)) Then
                
                ' Mark Match or Suggest
                If InStr(sRowTagList, selectedTag) > 0 Then
                    'Debug.Print ("Selected tag found in list of this row tag - Mark it")
                    rngBold.Font.Bold = True
                    rngFilter.Value = "Match"
                    rngFoundTag.Value = rngFoundTag.Value & " " & selectedTag
                    rngColorApply.Font.Color = RGB(56, 56, 56)
                    iNumberOfConnections = iNumberOfConnections + 1
                ElseIf InStr(sRowSubject, selectedTag) > 0 Then
                    'Debug.Print ("Selected tag found in subject - suggest it")
                    rngFilter.Value = "Sugest"
                    rngColorApply.Font.Color = RGB(128, 128, 128)
                End If
                
                'Lock rows have highest priority of sorting above current row + color green
                If (rngLock.Value = "yes") Then
                    rngFilter.Value = "Lock"
                    rngColorApply.Font.Color = RGB(0, 176, 80)
                End If
                
                'Color previous row - light blue
                If (rngSubject.Value = sPreviousSubject) Then
                    rngColorApply.Font.Color = RGB(142, 169, 219)
                End If
                
                'Selected row = 1 to make it before results + color Dark blue + update date
                If (lRowIndex = sSelectedRow) Then
                    rngFilter.Value = "Main"
                    shtMain.Cells(lRowIndex, iColDate).Value = todayDate
                    rngColorApply.Font.Color = RGB(48, 84, 150)
                End If
                
            End If
        
        Next lRowIndex
        
    Next selectedTag

    '==================================================
    'Final configs
    '==================================================
    
    'Save quantity of connections to current selected row
    shtMain.Cells(sSelectedRow, iColConnections).Value = iNumberOfConnections
    
    'Filter all matches and blank lines
    lo.Range.AutoFilter Field:=iColFilter, Operator:=xlFilterValues, _
        Criteria1:=Array("Main", "Match", "Sugest", "Lock")
    'TO DO - Make field dynamic if this column in different place
        
    'Restore initial settings
    Application.ScreenUpdating = True

    'Calculate Macro run time and print it
    SecondsElapsed = Round(Timer - StartTime, 2)
    Debug.Print ("Time took to run: " & SecondsElapsed & " sec")
    Debug.Print ("================ Finish =================")
    
End Sub
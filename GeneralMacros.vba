Sub EmphasizeSimilar()
' ==========================================================================
' Version: 1.0
' Description: Emphasize rows with similar tags and gray out all remaining
' Excel where used: My code write repeat
' ==========================================================================

    Debug.Print ("================ Start =================")
    Application.ScreenUpdating = False
    
    
    'Constants
    Const startingRow As Integer = 7
    Const tagColumn As String = "J"
    
    'Declare variables
    Dim currentRow As Integer
    Dim lastRow As Long
    Dim tagList As String
    Dim tagArray() As String
    Dim flag As Boolean
    
    'Get Current row and decide if exit
    currentRow = ActiveCell.Row
    If (currentRow < startingRow) Then Exit Sub
    
    lastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    tagList = Cells(currentRow, tagColumn)
    tagArray = Split(tagList, " ")
    
    ActiveSheet.Range(Cells(startingRow, "A"), Cells(lastRow, "J")).Font.Bold = False
    
    Debug.Print ("Current selected row: " & currentRow)
    Debug.Print ("tag list from current row: " & tagList)
    
    For i = startingRow To lastRow

        flag = False
        
        For Each Item In tagArray

            If InStr(1, Cells(i, tagColumn).Value, Item) Then
                flag = True
            End If
            
        Next Item
        
        If (flag) Then
            'ActiveSheet.Range(Cells(i, "A"), Cells(i, "J")).Font.Bold = True
            Debug.Print ("family row: " & i)
            Cells(i, "K").Value = "yes"
        Else
            Cells(i, "K").Value = "no"
        End If
        
    Next i
    
    ActiveSheet.ListObjects("Concepts").Range.AutoFilter Field:=11, Criteria1:="yes"
    Application.ScreenUpdating = True

End Sub
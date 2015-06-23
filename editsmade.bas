Attribute VB_Name = "Module3"
Sub editsmade()
'Records the edits made to a document and stores the text before and after the edits were made for later analysis

    Dim source As Document, target As Document
    Dim rtable As Table
    Dim rrow As Row
    Dim arev As Revision

    Set source = ActiveDocument
    Set target = Documents.Add
    
    'Write number of edits made at the top of the markup document
'    Dim mytext As String
'    mytext = source.Revisions.Count
'    Selection.TypeText (mytext)
    
    Set rtable = target.Tables.Add(target.range, 1, 4)
    With rtable
        .Cell(1, 1).range.text = "Revision"
        .Cell(1, 2).range.text = "Revision Type"
        .Cell(1, 3).range.text = "Author"
        .Cell(1, 4).range.text = "Page Number"
    End With
    
    For Each arev In source.Revisions
        Set rrow = rtable.Rows.Add
        With rrow
            .Cells(1).range.text = arev.range.text
            If arev.Type = 1 Then      'Work on changing over to using function below to convert to enumerations
                .Cells(2).range.text = "Insertion"
            ElseIf arev.Type = 2 Then
                .Cells(2).range.text = "Deletion"
            Else
                .Cells(2).range.text = arev.Type
            End If
            .Cells(3).range.text = arev.Author
            .Cells(4).range.text = arev.range.Information(wdActiveEndPageNumber)
        End With
    Next arev

    'Save with filename appended with "_edited" in AJE folder
    Dim filename As String
    Dim extension As String
    Dim pathname As String
    
    With source
        filename = .Name
        If Right(.Name, 1) = "x" Then
            filename = Left(.Name, Len(.Name) - 5)
            extension = ".docx"
        Else
            filename = Left(.Name, Len(.Name) - 4)
            extension = ".doc"
        End If
    End With
    filename = filename & "_markup"
    
    pathname = "Mac OS X:Users:jamesharper:Github:editsmade:"
    
    ActiveDocument.SaveAs (pathname & filename & extension)
        
End Sub
Sub ConvertRevisionTypeValueToEnumeration(value)
    
    Dim dict As Dictionary
    Dim v As Variant

    'Create the dictionary
    Set dict = New Dictionary

   'Add some (key, value) pairs
    dict.Add 0, "wdNoRevision"
    dict.Add 1, "wdRevisionInsert"
    dict.Add 2, "wdRevisionDelete"
    dict.Add 3, "wdRevisionProperty"
    dict.Add 4, "wdRevisionParagraphNumber"
    dict.Add 5, "wdRevisionDisplayField"
    dict.Add 6, "wdRevisionReconcile"
    dict.Add 7, "wdRevisionConflict"
    dict.Add 8, "wdRevisionStyle"
    dict.Add 9, "wdRevisionReplace"
    dict.Add 10, "wdRevisionParagraphProperty"
    dict.Add 11, "wdRevisionTableProperty"
    dict.Add 12, "wdRevisionSectionProperty"
    dict.Add 13, "wdRevisionStyleDefinition"
    dict.Add 14, "wdRevisionMovedFrom"
    dict.Add 15, "wdRevisionMovedTo"
    dict.Add 16, "wdRevisionCellInsertion"
    dict.Add 17, "wdRevisionCellDeletion"
    dict.Add 18, "wdRevisionCellMerge"
    dict.Add 20, "wdRevisionConflictInsert"
    dict.Add 21, "wdRevisionConflictDelete"

    'How many items do we have?
    Debug.Print "Number of items stored: " & dict.Count

    'We can retrieve an item based on the key
    Debug.Print "Ted is " & dict.Item(4) & " years old"

   'We can test whether an item exists
    Debug.Print "We have Jane's age: " & dict.Exists(17)
    Debug.Print "We have Zak's age " & dict.Exists(21)

   'And we can iterate through the complete dictionary
    For Each v In dict.Keys
        Debug.Print "Value: " & v & "Enumeration: "; dict.Item(v)
    Next

End Sub
    


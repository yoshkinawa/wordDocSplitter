Sub ExtractSectionsToNewDocuments()
    Dim originalDoc As Document
    Set originalDoc = ActiveDocument
    
    ' Create a new folder for the created files
    Dim folderPath As String
    folderPath = "path\to\folder"
    
    ' Check if the folder exists
    If Dir(folderPath, vbDirectory) <> "" Then

        ' If the folder exists, delete it
        On Error Resume Next
        Kill folderPath & "\*.*"
        RmDir folderPath
        On Error GoTo 0

    End If
    
    ' Create the folder
    MkDir folderPath
    
    Dim section As section
    Dim newDoc As Document
    Dim i As Integer
    
    ' Loop through each section in the original document
    For Each section In originalDoc.Sections

        ' Create a new document for each section
        Set newDoc = Documents.Add
        
        ' Copy the content of the section to the new document with formatting and styles
        section.Range.Copy
        newDoc.Range.Paste

        ' remove links
        Dim oField As Field
        For Each oField In newDoc.Fields
            If oField.Type = wdFieldHyperlink Then
                oField.Unlink
            End If
        Next
        
        ' Copy the "Normal" style definition from the original document to the new document
        newDoc.Styles("Normal").ParagraphFormat = originalDoc.Styles("Normal").ParagraphFormat
        newDoc.Styles("Normal").Font = originalDoc.Styles("Normal").Font
        
        ' Remove section breaks in the new document
        With newDoc.Range.Find
            .Text = "^b" ' ^b is the code for a section break
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll
        End With
        
        ' Set the file name for saving the new document in the created folder
        Dim fileName As String
        fileName = folderPath & "\section " & i & ".docx"
        
        ' Save the new document with the generated file name
        newDoc.SaveAs2 fileName
        
        ' Close the new document
        newDoc.Close
        
        ' Increment the section counter
        i = i + 1

    Next section
End Sub

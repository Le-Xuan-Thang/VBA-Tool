Sub DeleteUnusedStyles()
    ' Macro to delete unused styles in Word document - macOS Compatible
    
    Dim doc As Document
    Dim styleObj As Style
    Dim para As Paragraph
    Dim usedStyles As Collection
    Dim stylesToDelete As Collection
    Dim i As Integer
    Dim deletedCount As Integer
    Dim totalStyles As Integer
    
    ' Initialize current document
    Set doc = ActiveDocument
    
    ' Create collections to store used styles (thay thế Dictionary)
    Set usedStyles = New Collection
    Set stylesToDelete = New Collection
    
    ' Display start message
    Application.ScreenUpdating = False
    StatusBar = "Checking used styles..."
    
    ' Loop through all paragraphs to find used styles
    For Each para In doc.Paragraphs
        If Not IsStyleInCollection(usedStyles, para.Style.NameLocal) Then
            usedStyles.Add para.Style.NameLocal, para.Style.NameLocal
        End If
    Next para
    
    ' Find character styles through Find function
    Dim rng As Range
    Set rng = doc.Range
    Dim charStyle As Style
    
    ' Find character styles using Find function
    rng.Find.ClearFormatting
    For Each charStyle In doc.Styles
        If charStyle.Type = wdStyleTypeCharacter And Not charStyle.BuiltIn Then
            Set rng = doc.Range
            rng.Find.Style = charStyle
            If rng.Find.Execute Then
                If Not IsStyleInCollection(usedStyles, charStyle.NameLocal) Then
                    usedStyles.Add charStyle.NameLocal, charStyle.NameLocal
                End If
            End If
        End If
    Next charStyle
    
    ' Check styles in tables
    Dim tbl As Table
    Dim cellRange As Cell
    For Each tbl In doc.Tables
        For Each cellRange In tbl.Range.Cells
            If Not IsStyleInCollection(usedStyles, cellRange.Range.Style.NameLocal) Then
                usedStyles.Add cellRange.Range.Style.NameLocal, cellRange.Range.Style.NameLocal
            End If
        Next cellRange
    Next tbl
    
    ' Check styles in headers and footers
    Dim docSection As Section
    Dim hdrFtr As HeaderFooter
    For Each docSection In doc.Sections
        For Each hdrFtr In docSection.Headers
            If hdrFtr.Exists Then
                For Each para In hdrFtr.Range.Paragraphs
                    If Not IsStyleInCollection(usedStyles, para.Style.NameLocal) Then
                        usedStyles.Add para.Style.NameLocal, para.Style.NameLocal
                    End If
                Next para
            End If
        Next hdrFtr
        
        For Each hdrFtr In docSection.Footers
            If hdrFtr.Exists Then
                For Each para In hdrFtr.Range.Paragraphs
                    If Not IsStyleInCollection(usedStyles, para.Style.NameLocal) Then
                        usedStyles.Add para.Style.NameLocal, para.Style.NameLocal
                    End If
                Next para
            End If
        Next hdrFtr
    Next docSection
    
    StatusBar = "Identifying styles to delete..."
    
    ' Identify styles to delete
    totalStyles = doc.Styles.Count
    
    For Each styleObj In doc.Styles
        ' Only delete custom styles (not built-in) and not used
        If Not styleObj.BuiltIn And Not IsStyleInCollection(usedStyles, styleObj.NameLocal) Then
            If Not IsStyleInCollection(stylesToDelete, styleObj.NameLocal) Then
                stylesToDelete.Add styleObj.NameLocal, styleObj.NameLocal
            End If
        End If
    Next styleObj
    
    ' Delete styles
    StatusBar = "Deleting unused styles..."
    deletedCount = 0
    
    For i = 1 To stylesToDelete.Count
        On Error Resume Next
        doc.Styles(stylesToDelete.Item(i)).Delete
        If Err.Number = 0 Then
            deletedCount = deletedCount + 1
        End If
        On Error GoTo 0
    Next i
    
    ' Restore screen and status bar
    Application.ScreenUpdating = True
    StatusBar = ""
    
    ' Display results
    MsgBox "Completed!" & vbCrLf & vbCrLf & _
           "Total styles: " & totalStyles & vbCrLf & _
           "Deleted styles: " & deletedCount & vbCrLf & _
           "Remaining styles: " & doc.Styles.Count, _
           vbInformation, "Delete Unused Styles"
    
    ' Clean up objects
    Set usedStyles = Nothing
    Set stylesToDelete = Nothing
    Set doc = Nothing
End Sub

' Helper function to check if style exists in collection
Function IsStyleInCollection(col As Collection, styleName As String) As Boolean
    Dim i As Integer
    IsStyleInCollection = False
    
    On Error Resume Next
    Dim temp As String
    temp = col.Item(styleName)
    If Err.Number = 0 Then
        IsStyleInCollection = True
    End If
    On Error GoTo 0
End Function

' Helper macro: Display list of all styles
Sub ListAllStyles()
    Dim doc As Document
    Dim styleObj As Style
    Dim styleList As String
    Dim builtInCount As Integer
    Dim customCount As Integer
    
    Set doc = ActiveDocument
    
    styleList = "LIST OF ALL STYLES:" & vbCrLf & vbCrLf
    
    For Each styleObj In doc.Styles
        If styleObj.BuiltIn Then
            styleList = styleList & "[Built-in] " & styleObj.NameLocal & vbCrLf
            builtInCount = builtInCount + 1
        Else
            styleList = styleList & "[Custom] " & styleObj.NameLocal & vbCrLf
            customCount = customCount + 1
        End If
    Next styleObj
    
    styleList = styleList & vbCrLf & "Total: " & doc.Styles.Count & " styles" & vbCrLf
    styleList = styleList & "Built-in: " & builtInCount & vbCrLf
    styleList = styleList & "Custom: " & customCount
    
    MsgBox styleList, vbInformation, "Styles List"
End Sub

' Helper macro: Backup document before deleting
Sub BackupAndDeleteUnusedStyles()
    Dim response As Integer
    
    response = MsgBox("Do you want to create a backup before deleting styles?" & vbCrLf & _
                     "Press Yes to backup, No to continue without backup, Cancel to abort.", _
                     vbYesNoCancel + vbQuestion, "Backup Document")
    
    Select Case response
        Case vbYes
            ' Create backup
            ActiveDocument.SaveAs2 FileName:=ActiveDocument.Path & "/" & _
                                  Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1) & _
                                  "_backup_" & Format(Now, "yyyymmdd_hhmmss") & ".docx"
            MsgBox "Backup created successfully!", vbInformation
            Call DeleteUnusedStyles
        Case vbNo
            Call DeleteUnusedStyles
        Case vbCancel
            Exit Sub
    End Select
End Sub

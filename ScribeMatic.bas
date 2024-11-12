' Module: ScribeMatic

Private INVALIDCHARS As Variant

' Define the change type enumeration
Private Enum changeType
    Insert = 1
    delete = -1
    Replace = 0
End Enum

' Class-like structure within a module (simulated)
Private Type LevenshteinItem
    position As Integer
    oldChar As String
    newChar As String
    changeType As changeType
End Type

' Main procedure for performing text comparison and displaying changes
Sub ScribeMatic()
    INVALIDCHARS = Array(vbCrLf, vbCr, vbLf)
    Dim doc As Document
    Dim selectedText As String
    Dim startRemoved As Integer
    Dim filePath As String
    Dim fileText As String
    Dim fileDialog As fileDialog
    Dim changes As New Collection, cleanedChanges As New Collection
    Dim change As Object
    Dim i As Integer, j As Integer
    Dim oldChar As String, newChar As String
    Dim keystrokeGoal As Integer
    Dim userInput As String

    ' Set the active document
    Set doc = ActiveDocument

    ' Get the selected text and clean it
    Dim dict As Object
    Set dict = cleanText(selection.text)
    selectedText = dict("cleanedText")
    startRemoved = dict("startRemoved")
    
    selection.Start = selection.Start + startRemoved
    selection.End = selection.Start + Len(selectedText)
    
    If Len(selectedText) = 0 Then
        MsgBox "No text selected."
        Exit Sub
    End If
    
    selectedTextKeystrokes = countKeystrokesFromLine(selectedText)

    ' Open file dialog to select a text file
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    If fileDialog.Show = -1 Then
        filePath = fileDialog.SelectedItems(1)
    Else
        MsgBox "No file selected."
        Exit Sub
    End If
    
    ' Prompt for keystroke count
    userInput = InputBox("Please enter the number of keystrokes (character count) to start comparison from:", "Keystroke Count", "1")
    If Not IsNumeric(userInput) Then
       MsgBox "Invalid input. Please enter a numeric value."
        Exit Sub
    End If
    keystrokeGoal = CInt(userInput)
    
    ' Load text from the selected file
    Open filePath For Input As #1
    fileText = Input$(LOF(1), 1)
    Close #1

    ' Initialize collection to store changes
    Dim int1 As Long, int2 As Long, int3 As Long
    Set changes = LevenshteinDifferences(selectedText, fileText)

    ' Process and clean up changes based on keystroke limits
    If changes.Count > 0 Then
        ' Add keystroke count to each change
        changes(changes.Count).Add "keystroke", countKeystrokesFromLine(mid(selectedText, 1, changes(changes.Count)("position"))) + countKeystrokesFromLine(changes(changes.Count)("newChar")) - countKeystrokesFromLine(changes(changes.Count)("oldChar"))
        For i = changes.Count - 1 To 1 Step -1
            int1 = changes(i + 1)("keystroke")
            int2 = countKeystrokesFromLine(mid(selectedText, changes(i + 1)("position") + 1, changes(i)("position") - changes(i + 1)("position")))
            int3 = countKeystrokesFromLine(changes(i)("newChar")) - countKeystrokesFromLine(changes(i)("oldChar"))
            changes(i).Add "keystroke", int1 + int2 + int3
        Next i
        
        ' Filter changes based on keystrokeGoal
        Dim endPos As Long
        endPos = changes(1)("position")
        For Each elem In changes
            If Not (elem("changeType") = Insert And elem("keystroke") > keystrokeGoal) Or endPos <> elem("position") Then
                cleanedChanges.Add elem
            End If
        Next elem
        Set changes = cleanedChanges
    End If
    
    
    ' Mark changes in selection
    Call editChanges(changes, selection)
    
    ' Add important information below the table
    Dim keystrokeText As String
    keystrokeText = selectedTextKeystrokes & " / " & keystrokeGoal & " keystrokes reached."
    doc.Range(doc.Content.End - 1, doc.Content.End).InsertAfter vbCrLf & keystrokeText

    MsgBox "Comparison complete! Changes have been listed."
End Sub

' Function to calculate Levenshtein differences between selected and file text
Function LevenshteinDifferences(selectedText As String, fileText As String) As Collection
    Dim len1 As Long, len2 As Long
    Dim d() As Long
    Dim i As Long, j As Long
    Dim change As Object
    Dim changes As New Collection

    len1 = Len(selectedText)
    len2 = Len(fileText)
    
    ' Initialize the distance matrix
    ReDim d(len1, len2)
    
    ' Fill the matrix with base values
    For i = 0 To len1
        d(i, 0) = i
    Next i
    For j = 0 To len2
        d(0, j) = j
    Next j
    
    ' Compute Levenshtein distance and extract differences
    For i = 1 To len1
        For j = 1 To len2
            If mid(selectedText, i, 1) = mid(fileText, j, 1) Then
                d(i, j) = d(i - 1, j - 1) ' No change
            Else
                d(i, j) = Min(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + 1) ' Insert, Delete, Replace
            End If
        Next j
    Next i
    
    ' Backtrack to build the changes collection
    i = len1
    j = len2
    While i > 0 Or j > 0
        If i > 0 And d(i - 1, j) + 1 = d(i, j) Then
            ' Delete (from selectedText)
            Set change = getChangeDict(delete, i, mid(selectedText, i, 1), "")
            Call addSafety(changes, change)
            i = i - 1
        ElseIf j > 0 And d(i, j - 1) + 1 = d(i, j) Then
            ' Insert (in fileText)
            Set change = getChangeDict(Insert, i, "", mid(fileText, j, 1))
            Call addSafety(changes, change)
            j = j - 1
        ElseIf i > 0 And j > 0 And d(i - 1, j - 1) + 1 = d(i, j) Then
            ' Replace
            Set change = getChangeDict(Replace, i, mid(selectedText, i, 1), mid(fileText, j, 1))
            Call addSafety(changes, change)
            i = i - 1
            j = j - 1
        Else
            ' No difference
            i = i - 1
            j = j - 1
        End If
    Wend

    Set LevenshteinDifferences = changes
End Function

' Function to clean text by removing invalid characters from both ends
Private Function cleanText(selectedText As String) As Object
    Dim cleaned As Boolean
    Set dict = CreateObject("Scripting.Dictionary")
    Dim startRemoved As Long
    startRemoved = 0
    
    ' Loop to remove invalid characters from the start of the text
    cleaned = False
    Do While Not cleaned And Len(selectedText) > 0
        cleaned = True
        If characterInArray(left(selectedText, 1), INVALIDCHARS) Then
            selectedText = Right(selectedText, Len(selectedText) - 1)
            cleaned = False
            startRemoved = startRemoved + 1
        End If
    Loop
    
    ' Loop to remove invalid characters from the end of the text
    cleaned = False
    Do While Not cleaned And Len(selectedText) > 0
        cleaned = True
        If characterInArray(Right(selectedText, 1), INVALIDCHARS) Then
            selectedText = left(selectedText, Len(selectedText) - 1)
            cleaned = False
        End If
    Loop
    dict.Add "cleanedText", selectedText
    dict.Add "startRemoved", startRemoved
    Set cleanText = dict
End Function

' Ensures only valid changes are added to the collection by excluding changes with invalid characters
Private Sub addSafety(ByRef changes As Collection, change As Object)
    If characterInArray(change("newChar"), INVALIDCHARS) Or characterInArray(change("oldChar"), INVALIDCHARS) Then
        Exit Sub
    End If
    changes.Add change
End Sub

' Creates and returns a dictionary object for a change, storing position, oldChar, newChar, and changeType
Private Function getChangeDict(ByVal changeType As changeType, ByVal position As Long, ByVal oldChar As String, ByVal newChar As String)
    Set change = CreateObject("Scripting.Dictionary")
    change.Add "position", position
    change.Add "oldChar", oldChar
    change.Add "newChar", newChar
    change.Add "changeType", changeType
    Set getChangeDict = change
End Function


' Calculates the minimum of three values
Private Function Min(val1 As Long, val2 As Long, val3 As Long) As Long
    Min = val1
    If val2 < Min Then Min = val2
    If val3 < Min Then Min = val3
End Function

' Converts the ChangeType enum value to a string
Private Function getChangeTypeString(ByVal changeType As changeType) As String
    Select Case changeType
        Case Insert
            getChangeTypeString = "Insert"
        Case delete
            getChangeTypeString = "Delete"
        Case Replace
            getChangeTypeString = "Replace"
        Case Else
            getChangeTypeString = "Unknown"
    End Select
End Function

' Converts the ChangeType enum value to an integer
Private Function getChangeTypeInteger(ByVal inputType As changeType) As Integer
    Select Case inputType
        Case changeType.Insert
            getChangeTypeInteger = changeType.Insert
        Case changeType.delete
            getChangeTypeInteger = changeType.delete
        Case changeType.Replace
            getChangeTypeInteger = changeType.Replace
        Case Else
            getChangeTypeInteger = -99 ' Falls der Typ nicht definiert ist
    End Select
End Function

' Counts the total number of keystrokes required to type all characters in a string
Private Function countKeystrokesFromLine(ByVal inputString As String) As Integer
    Dim keystrokes As Integer
    Dim character As String
    keystrokes = 0
    For i = 1 To Len(inputString)
        keystrokes = keystrokes + getKeystrokeFromCharacter(mid(inputString, i, 1))
    Next i
    countKeystrokesFromLine = keystrokes
End Function

' Determines the number of keystrokes required for a specific character
Private Function getKeystrokeFromCharacter(character As String)
    doubleKeystrokes = Array("€", "\", "{", "[", "]", "}", "²", "³", "°", "!", """", "§", "$", "%", "&", "/", "(", ")", "=", "?", "*", ">", ";", ":", "_", "@", "|", "'")
    
    If character = "…" Then
       getKeystrokeFromCharacter = 3
    ElseIf character Like "[A-Z]" Or characterInArray(character, doubleKeystrokes) Then
        getKeystrokeFromCharacter = 2
    ElseIf Not characterInArray(character, INVALIDCHARS) Then
        getKeystrokeFromCharacter = 1
    End If
End Function

' Checks if a character is present in a given array
Private Function characterInArray(character As String, arr As Variant) As Boolean
    characterInArray = False
    For Each element In arr
        If element = character Then
            characterInArray = True
            Exit Function
        End If
    Next element
End Function

' Processes a list of changes (Insert, delete, Replace) to a selected text range in a Word document, adding textboxes to highlight changes and adjusting the format of modified text areas.
Private Sub editChanges(changes As Object, ByRef selection As Object)
    Dim originalStart As Long, originalEnd As Long
    Dim lastInsertPos As Long
    Dim lastInsertTextBox As Object
    lastInsertPos = -1
    originalStart = selection.Start
    
    ' Set line spacing for selected paragraph
    selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
    selection.ParagraphFormat.LineSpacing = 36
    
    For Each change In changes
        originalEnd = selection.End
        If Len(selection.text) >= change("position") Then
            Select Case change("changeType")
                Case Insert
                    ' Check if insertion position is the same as the last one
                    If lastInsertPos = change("position") Then
                        Dim textLen As Integer
                        Dim textBoxStr As String
                        textLen = Len(lastInsertTextBox.TextFrame.TextRange.text)
                        textBoxStr = lastInsertTextBox.TextFrame.TextRange.text
                        If textLen < 15 Then
                            lastInsertTextBox.left = lastInsertTextBox.left - 5
                            lastInsertTextBox.width = lastInsertTextBox.width + 10
                            lastInsertTextBox.TextFrame.TextRange.text = change("newChar") & left(textBoxStr, textLen - 1)
                        ElseIf textLen = 15 Then
                            lastInsertTextBox.left = lastInsertTextBox.left - 5
                            lastInsertTextBox.width = lastInsertTextBox.width + 10
                            lastInsertTextBox.TextFrame.TextRange.text = change("newChar") & left(textBoxStr, 5) & "[…]" & left(Right(textBoxStr, 9), 8)
                        Else
                            lastInsertTextBox.TextFrame.TextRange.text = change("newChar") & left(textBoxStr, 5) & "[…]" & left(Right(textBoxStr, 9), 8)
                        End If
                    Else
                        ' Set new insertion position
                        lastInsertPos = change("position")
                        selection.Start = selection.Start + change("position") - 1
                        selection.End = selection.Start + 1
                        
                        ' Add new text box for insertion
                        Set lastInsertTextBox = ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                            selection.Range.Information(wdHorizontalPositionRelativeToTextBoundary), _
                            selection.Range.Information(wdVerticalPositionRelativeToTextBoundary), _
                            10, 35)
                            
                        lastInsertTextBox.left = selection.Range.Information(wdHorizontalPositionRelativeToTextBoundary) + 2.5
                        lastInsertTextBox.Top = selection.Range.Information(wdVerticalPositionRelativeToTextBoundary)
                        
                        ' Format text box
                        formatTextBox (textbox)
                        
                        ' Set content of text box
                        lastInsertTextBox.TextFrame.TextRange.text = change("newChar") & vbCrLf & ChrW(&H22A5)
                        
                        ' Update selection
                        selection.End = originalEnd + 1
                    End If
                Case delete
                    selection.Start = selection.Start + change("position") - 1
                    selection.End = selection.Start + 1
                    selection.Font.Underline = wdUnderlineThick
                    selection.Font.UnderlineColor = RGB(255, 0, 1)
                    
                    ' Add text box with delete marker
                    Set textbox = ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                        selection.Range.Information(wdHorizontalPositionRelativeToTextBoundary), _
                        selection.Range.Information(wdVerticalPositionRelativeToTextBoundary), _
                        15, 20)
                        
                    textbox.left = selection.Range.Information(wdHorizontalPositionRelativeToTextBoundary) - 4
                    textbox.Top = selection.Range.Information(wdVerticalPositionRelativeToTextBoundary) + 12
                    
                    ' Format text box
                    formatTextBox (textbox)
                    
                    ' Set delete marker
                    textbox.TextFrame.TextRange.text = "/"
                    
                    ' Update selection
                    selection.End = originalEnd + 1
                Case Replace
                    selection.Start = selection.Start + change("position") - 1
                    selection.End = selection.Start + 1
                    selection.Font.Underline = wdUnderlineThick
                    selection.Font.UnderlineColor = RGB(255, 0, 1)
                    
                    ' Add text box with replacement character
                    Set textbox = ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                        selection.Range.Information(wdHorizontalPositionRelativeToTextBoundary), _
                        selection.Range.Information(wdVerticalPositionRelativeToTextBoundary), _
                        15, 20)
                        
                    textbox.left = selection.Range.Information(wdHorizontalPositionRelativeToTextBoundary) - 4
                    textbox.Top = selection.Range.Information(wdVerticalPositionRelativeToTextBoundary)
                                        
                    ' Format text box
                    formatTextBox (textbox)
                    
                    ' Set replacement text
                    textbox.TextFrame.TextRange.text = change("newChar")
                    
                    ' Update selection
                    selection.End = originalEnd + 1
            End Select
            ' Reset selection to original start
            selection.Start = originalStart
        End If
    Next change
End Sub

' Formats the given text box by setting its font, margins, alignment, and making the text transparent with no border.
Private Sub formatTextBox(ByRef textbox As Object)
    textbox.Fill.Transparency = 1
    textbox.TextFrame.TextRange.Font.Size = 12
    textbox.TextFrame.TextRange.Font.Name = "Courier New"
    textbox.TextFrame.TextRange.Font.COLOR = RGB(255, 0, 1)
    textbox.TextFrame.TextRange.Font.Bold = True
    textbox.TextFrame.MarginTop = 0
    textbox.TextFrame.MarginBottom = 0
    textbox.TextFrame.MarginLeft = 0
    textbox.TextFrame.MarginRight = 0
    textbox.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
    textbox.TextFrame.VerticalAnchor = msoAnchorBottom
    textbox.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
    textbox.line.Visible = msoFalse
End Sub

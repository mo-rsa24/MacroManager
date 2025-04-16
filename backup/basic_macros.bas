1.
Sub CreateHierarchicalBiDirectionalLinksb()
    Dim oDoc As Object, oText As Object, oVCursor As Object
    Dim sSectionTitle As String, sCleanTitle As String
    Dim oBookmarks As Object, oTextCursor As Object
    Dim oInsertCursor As Object, oBookmark As Object
    Dim sParentBookmark As String, sBookmarkMain As String, sBookmarkTOC As String

    oDoc = ThisComponent
    oText = oDoc.Text
    oVCursor = oDoc.CurrentController.ViewCursor

    ' Get selected text explicitly
    sSectionTitle = Trim(oVCursor.getString())

    ' Validate colon at end
    If Right(sSectionTitle, 1) <> ":" Then
        MsgBox "Selection must end with a colon (:)", 48, "Formatting Error"
        Exit Sub
    End If

    ' Clean title (remove trailing colon)
    sCleanTitle = Trim(Left(sSectionTitle, Len(sSectionTitle) - 1))

    ' Parent bookmark input
    sParentBookmark = InputBox("Enter exact parent bookmark name (e.g., Section 1 To Now)", _
                               "Parent Bookmark")
    If sParentBookmark = "" Then Exit Sub

    ' Build hierarchical bookmark names
    sBookmarkMain = sParentBookmark & " " & sCleanTitle
    sBookmarkTOC = sBookmarkMain & " Contents"

    oBookmarks = oDoc.getBookmarks()

    ' Ensure bookmarks don't exist already
    If oBookmarks.hasByName(sBookmarkMain) Or oBookmarks.hasByName(sBookmarkTOC) Then
        MsgBox "Bookmark already exists. Choose a unique title or parent.", 48, "Duplicate Bookmark"
        Exit Sub
    End If

    '--- Select just the colon from the current selection ---
    oTextCursor = oText.createTextCursorByRange(oVCursor)
    oTextCursor.collapseToEnd()
    oTextCursor.goLeft(1, True) ' Precisely select the colon
    oTextCursor.HyperLinkURL = ""

    ' Create main bookmark by directly adding it to the bookmark container.
    ' This attaches the bookmark to the exact text range (the colon) so it will move with the text.
    oBookmark = oBookmarks.addNew(sBookmarkMain, oTextCursor)

    ' Set hyperlink on the same text range to later point to the TOC bookmark
    oTextCursor.HyperLinkURL = "#" & sBookmarkTOC

    ' Insert the navigation line above the current paragraph
    oInsertCursor = oText.createTextCursorByRange(oVCursor.Start)
    oInsertCursor.gotoStartOfParagraph(False)
    oText.insertString(oInsertCursor, sCleanTitle & Chr(13), False)

    ' Hyperlink the navigation line to the main bookmark
    oInsertCursor.goLeft(Len(sCleanTitle) + 1, True)
    oInsertCursor.HyperLinkURL = "#" & sBookmarkMain

    ' Create the TOC bookmark on the navigation line by attaching it to that text range
    oBookmark = oBookmarks.addNew(sBookmarkTOC, oInsertCursor)

    MsgBox "✅ Hierarchical navigation created for: " & sBookmarkMain, 64, "Success"
End Sub


2.

Sub CreateHierarchicalBiDirectionalLinksa()
    Dim oDoc As Object, oText As Object, oVCursor As Object
    Dim sSectionTitle As String, sCleanTitle As String
    Dim oBookmarks As Object, oTextCursor As Object
    Dim oInsertCursor As Object, oBookmark As Object
    Dim sParentBookmark As String, sBookmarkMain As String, sBookmarkTOC As String

    oDoc = ThisComponent
    oText = oDoc.Text
    oVCursor = oDoc.CurrentController.ViewCursor

    ' Get selected text explicitly
    sSectionTitle = Trim(oVCursor.getString())

    ' Validate colon at end
    If Right(sSectionTitle, 1) <> ":" Then
        MsgBox "Selection must end with a colon (:)", 48, "Formatting Error"
        Exit Sub
    End If

    ' Clean title
    sCleanTitle = Trim(Left(sSectionTitle, Len(sSectionTitle) - 1))

    ' Parent bookmark input
    sParentBookmark = InputBox("Enter exact parent bookmark name (e.g., Section 1 To Now)", _
                               "Parent Bookmark")
    If sParentBookmark = "" Then Exit Sub

    ' Build hierarchical bookmarks
    sBookmarkMain = sParentBookmark & " " & sCleanTitle
    sBookmarkTOC = sBookmarkMain & " Contents"

    oBookmarks = oDoc.getBookmarks()

    ' Ensure bookmarks don't exist
    If oBookmarks.hasByName(sBookmarkMain) Or oBookmarks.hasByName(sBookmarkTOC) Then
        MsgBox "Bookmark already exists. Choose a unique title or parent.", 48, "Duplicate Bookmark"
        Exit Sub
    End If

    '--- Select just the colon from the current selection ---
    oTextCursor = oText.createTextCursorByRange(oVCursor)
    oTextCursor.collapseToEnd()
    oTextCursor.goLeft(1, True) ' Precisely select the colon
    oTextCursor.HyperLinkURL = ""

    ' Create main bookmark and embed it with the text so it moves on cut/paste
    oBookmark = oDoc.createInstance("com.sun.star.text.Bookmark")
    oBookmark.Name = sBookmarkMain
    oText.insertTextContent(oTextCursor, oBookmark, True)  ' Use True to replace the selected text

    ' Hyperlink the (now bookmarked) colon to the TOC bookmark
    oTextCursor.HyperLinkURL = "#" & sBookmarkTOC

    ' Insert navigation line above
    oInsertCursor = oText.createTextCursorByRange(oVCursor.Start)
    oInsertCursor.gotoStartOfParagraph(False)
    oText.insertString(oInsertCursor, sCleanTitle & Chr(13), False)

    ' Hyperlink navigation line to main bookmark
    oInsertCursor.goLeft(Len(sCleanTitle) + 1, True)
    oInsertCursor.HyperLinkURL = "#" & sBookmarkMain

    ' Create TOC bookmark on the navigation line and embed it with the text
    oBookmark = oDoc.createInstance("com.sun.star.text.Bookmark")
    oBookmark.Name = sBookmarkTOC
    oText.insertTextContent(oInsertCursor, oBookmark, True)  ' Use True to have it attached to the text

    MsgBox "✅ Hierarchical navigation created for: " & sBookmarkMain, 64, "Success"
End Sub
3.


Sub CreateCommaSeparatedListFromBullets()
    Dim oDoc As Object, oSelection As Object
    Dim oEnum As Object, oPara As Object
    Dim sList As String, sText As String
    Dim pos As Integer
    Dim i As Integer

    oDoc = ThisComponent
    oSelection = oDoc.CurrentSelection

    ' Ensure there is a selection
    If oSelection.getCount() = 0 Then
        MsgBox "Please select the nested bullet structure first."
        Exit Sub
    End If

    sList = ""

    ' Loop through each selected text range (a contiguous selection may have 1 item with many paragraphs)
    For i = 0 To oSelection.getCount() - 1
        ' Create an enumeration for the paragraphs in this portion of the selection
        oEnum = oSelection.getByIndex(i).createEnumeration()
        While oEnum.hasMoreElements()
            oPara = oEnum.nextElement()
            sText = oPara.getString()

            ' Look for the colon character in the paragraph
            pos = InStr(sText, ":")
            If pos > 0 Then
                ' Extract text before the colon and trim spaces
                sText = Trim(Left(sText, pos - 1))
                ' Remove bullet characters if they are present at the very start
                sText = RemoveBullet(sText)
                ' Append to list; add comma only if sList already has content
                If sList = "" Then
                    sList = sText
                Else
                    sList = sList & ", " & sText
                End If
            End If
        Wend
    Next i

    ' Insert the generated comma-separated string above the first paragraph of the selection.
    ' We obtain a reference to the start of the first selected text range.
    Dim oFirstRange As Object
    oFirstRange = oSelection.getByIndex(0).getStart()
    oFirstRange.getText().insertString(oFirstRange, sList & CHR(13), False)

    ' Optional: Notify that the string has been inserted.
    MsgBox "The list '" & sList & "' has been inserted."
	End Sub

Sub CreateBiDirectionalLinks()
    Dim oDoc As Object
    Dim oText As Object
    Dim oVCursor As Object
    Dim sSectionTitle As String
    Dim sCleanTitle As String
    Dim sSectionNum As String
    Dim oBookmarks As Object
    Dim oTextCursor As Object
    Dim oInsertCursor As Object
    Dim oBookmark As Object
    Dim colonPos As Integer

    oDoc = ThisComponent
    oText = oDoc.Text
    oVCursor = oDoc.CurrentController.ViewCursor

    ' Get the section title from the current selection
    sSectionTitle = oVCursor.getString()

    ' Locate the first colon in the section title.
    colonPos = InStr(sSectionTitle, ":")
    If colonPos = 0 Then
        MsgBox "Please include a colon (:) in the section title.", 48, "Formatting Error"
        Exit Sub
    End If

    ' Remove the colon and trim any extra spaces to get the clean section title
    sCleanTitle = Trim(Left(sSectionTitle, colonPos - 1))

    ' Ask for a section number
    sSectionNum = InputBox("Enter Section Number (e.g., 1)", "Section Number")
    If sSectionNum = "" Then Exit Sub

    ' Construct bookmark names
    Dim sBookmarkMain As String
    Dim sBookmarkTOC As String
    sBookmarkMain = "Section " & sSectionNum & " " & sCleanTitle
    sBookmarkTOC = sBookmarkMain & " Contents"

    ' Get the bookmarks container
    oBookmarks = oDoc.getBookmarks()

    ' Check if the bookmark names are already used
    If oBookmarks.hasByName(sBookmarkMain) Then
        MsgBox "A bookmark with the name '" & sBookmarkMain & "' already exists." & _
               Chr(13) & "Please choose a different section number or section title.", 48, "Bookmark Exists"
        Exit Sub
    End If

    If oBookmarks.hasByName(sBookmarkTOC) Then
        MsgBox "A bookmark with the name '" & sBookmarkTOC & "' already exists." & _
               Chr(13) & "Please choose a different section number or section title.", 48, "Bookmark Exists"
        Exit Sub
    End If

    ' --- Step 1: Create bookmark on the colon and hyperlink it back to TOC ---
    ' Create a new text cursor starting from the beginning of the selection and collapse it to the start.
    oTextCursor = oText.createTextCursorByRange(oVCursor)
    oTextCursor.collapseToStart()
    ' Move the cursor to the colon position (colon is at position colonPos in the string).
    oTextCursor.goRight(colonPos - 1, False)
    ' Now select exactly one character (the colon)
    oTextCursor.goRight(1, True)

    ' Clear any existing hyperlink attribute on that character range
    oTextCursor.HyperLinkURL = ""

    ' Create and insert the bookmark for the main section into the document
    oBookmark = oDoc.createInstance("com.sun.star.text.Bookmark")
    oBookmark.Name = sBookmarkMain
    oText.insertTextContent(oTextCursor, oBookmark, False)

    ' Link the colon back to the TOC bookmark target
    oTextCursor.HyperLinkURL = "#" & sBookmarkTOC

    ' --- Step 2: Insert navigation line above the heading ---
    oInsertCursor = oText.createTextCursorByRange(oVCursor)
    oInsertCursor.gotoStartOfParagraph(False)
    oText.insertString(oInsertCursor, sCleanTitle & Chr(13), False)

    ' --- Step 3: Create hyperlink to the main section from the nav title ---
    oInsertCursor.goLeft(Len(sCleanTitle) + 1, True)
    oInsertCursor.HyperLinkURL = "#" & sBookmarkMain

    ' --- Step 4: Bookmark the navigation title for TOC use ---
    oBookmark = oDoc.createInstance("com.sun.star.text.Bookmark")
    oBookmark.Name = sBookmarkTOC
    oText.insertTextContent(oInsertCursor, oBookmark, False)

    MsgBox "✅ Bi-directional navigation created for: " & sCleanTitle, 64, "Complete"
End Sub

Sub CreateHierarchicalBiDirectionalLinks()
    Dim oDoc As Object, oText As Object, oVCursor As Object
    Dim sSectionTitle As String, sCleanTitle As String
    Dim oBookmarks As Object, oTextCursor As Object
    Dim oInsertCursor As Object, oBookmark As Object
    Dim sParentBookmark As String, sBookmarkMain As String, sBookmarkTOC As String

    oDoc = ThisComponent
    oText = oDoc.Text
    oVCursor = oDoc.CurrentController.ViewCursor

    ' Get selected text explicitly
    sSectionTitle = Trim(oVCursor.getString())

    ' Validate colon at end
    If Right(sSectionTitle, 1) <> ":" Then
        MsgBox "Selection must end with a colon (:)", 48, "Formatting Error"
        Exit Sub
    End If

    ' Clean title
    sCleanTitle = Trim(Left(sSectionTitle, Len(sSectionTitle) - 1))

    ' Parent bookmark input
    sParentBookmark = InputBox("Enter exact parent bookmark name (e.g., Section 1 To Now)", _
                               "Parent Bookmark")
    If sParentBookmark = "" Then Exit Sub

    ' Build hierarchical bookmarks
    sBookmarkMain = sParentBookmark & " " & sCleanTitle
    sBookmarkTOC = sBookmarkMain & " Contents"

    oBookmarks = oDoc.getBookmarks()

    ' Ensure bookmarks don't exist
    If oBookmarks.hasByName(sBookmarkMain) Or oBookmarks.hasByName(sBookmarkTOC) Then
        MsgBox "Bookmark already exists. Choose a unique title or parent.", 48, "Duplicate Bookmark"
        Exit Sub
    End If

    '--- FIX: Properly select just the colon from the user's current selection ---
    oTextCursor = oText.createTextCursorByRange(oVCursor)
    oTextCursor.collapseToEnd()
    oTextCursor.goLeft(1, True) ' Precisely select the colon
    oTextCursor.HyperLinkURL = ""

    oBookmark = oDoc.createInstance("com.sun.star.text.Bookmark")
    oBookmark.Name = sBookmarkMain
    oText.insertTextContent(oTextCursor, oBookmark, False)

    ' Hyperlink colon to TOC bookmark
    oTextCursor.HyperLinkURL = "#" & sBookmarkTOC

    ' Insert navigation line above
    oInsertCursor = oText.createTextCursorByRange(oVCursor.Start)
    oInsertCursor.gotoStartOfParagraph(False)
    oText.insertString(oInsertCursor, sCleanTitle & Chr(13), False)

    ' Hyperlink navigation line to main bookmark
    oInsertCursor.goLeft(Len(sCleanTitle) + 1, True)
    oInsertCursor.HyperLinkURL = "#" & sBookmarkMain

    ' Bookmark navigation line
    oBookmark = oDoc.createInstance("com.sun.star.text.Bookmark")
    oBookmark.Name = sBookmarkTOC
    oText.insertTextContent(oInsertCursor, oBookmark, False)

    MsgBox "✅ Hierarchical navigation created for: " & sBookmarkMain, 64, "Success"
	End Sub


Sub CreateBiDirectionalLinksCustomTitle()
    Dim oDoc As Object
    Dim oText As Object
    Dim oVCursor As Object
    Dim sSectionTitle As String
    Dim sCleanTitle As String
    Dim sCustomTitle As String
    Dim oBookmarks As Object
    Dim oTextCursor As Object
    Dim oInsertCursor As Object
    Dim oBookmark As Object
    Dim colonPos As Integer
    Dim sDocURL As String

    oDoc = ThisComponent
    oText = oDoc.Text
    oVCursor = oDoc.CurrentController.ViewCursor

    ' Get the section title from the current selection
    sSectionTitle = oVCursor.getString()

    ' Locate the first colon in the section title.
    colonPos = InStr(sSectionTitle, ":")
    If colonPos = 0 Then
        MsgBox "Please include a colon (:) in the section title.", 48, "Formatting Error"
        Exit Sub
    End If

    ' Remove the colon and trim any extra spaces to get the clean section title.
    sCleanTitle = Trim(Left(sSectionTitle, colonPos - 1))

    ' Prompt the user to enter a custom bookmark title.
    sCustomTitle = InputBox("Enter Bookmark Title", "Bookmark Title")
    If sCustomTitle = "" Then Exit Sub

    ' Construct bookmark names using the custom title.
    Dim sBookmarkMain As String
    Dim sBookmarkTOC As String
    sBookmarkMain = sCustomTitle
    sBookmarkTOC = sBookmarkMain & " Contents"

    ' Get the bookmarks container.
    oBookmarks = oDoc.getBookmarks()

    ' Check if the bookmark names are already defined.
    If oBookmarks.hasByName(sBookmarkMain) Then
        MsgBox "A bookmark with the name '" & sBookmarkMain & "' already exists." & _
               Chr(13) & "Please choose a different bookmark title.", 48, "Bookmark Exists"
        Exit Sub
    End If

    If oBookmarks.hasByName(sBookmarkTOC) Then
        MsgBox "A bookmark with the name '" & sBookmarkTOC & "' already exists." & _
               Chr(13) & "Please choose a different bookmark title.", 48, "Bookmark Exists"
        Exit Sub
    End If

    ' Retrieve the document URL
    sDocURL = oDoc.URL
    If sDocURL = "" Then
        MsgBox "Document must be saved to create bi-directional navigation.", 48, "Error"
        Exit Sub
    End If

    ' --- Step 1: Create bookmark on the colon and hyperlink it back to TOC ---
    ' Create a new text cursor starting at the beginning of the selection.
    oTextCursor = oText.createTextCursorByRange(oVCursor)
    oTextCursor.collapseToStart()
    ' Move the cursor to the colon position.
    oTextCursor.goRight(colonPos - 1, False)
    ' Now select exactly one character (the colon).
    oTextCursor.goRight(1, True)

    ' Clear any existing hyperlink attribute on that character range.
    oTextCursor.HyperLinkURL = ""

    ' Create and insert the bookmark for the main section into the document.
    oBookmark = oDoc.createInstance("com.sun.star.text.Bookmark")
    oBookmark.Name = sBookmarkMain
    oText.insertTextContent(oTextCursor, oBookmark, False)

    ' Link the colon back to the TOC bookmark target.
    oTextCursor.HyperLinkURL = "#" & sBookmarkTOC

    ' --- Step 2: Insert navigation line above the heading ---
    oInsertCursor = oText.createTextCursorByRange(oVCursor)
    oInsertCursor.gotoStartOfParagraph(False)
    oText.insertString(oInsertCursor, sCleanTitle & Chr(13), False)

    ' --- Step 3: Create hyperlink to the main section from the nav title ---
    oInsertCursor.goLeft(Len(sCleanTitle) + 1, True)
    oInsertCursor.HyperLinkURL = sDocURL & "#" & sBookmarkMain

    ' --- Step 4: Bookmark the navigation title for TOC use ---
    oBookmark = oDoc.createInstance("com.sun.star.text.Bookmark")
    oBookmark.Name = sBookmarkTOC
    oText.insertTextContent(oInsertCursor, oBookmark, False)

    MsgBox "✅ Bi-directional navigation created for: " & sCleanTitle, 64, "Complete"
End Sub



' This helper function removes bullet characters at the very start of a string.
Function RemoveBullet(s As String) As String
    Dim firstChar As String
    firstChar = Left(s, 1)
    If firstChar = "•" Or firstChar = "◦" Or firstChar = "-" Then
        RemoveBullet = Trim(Mid(s, 2))
    Else
        RemoveBullet = s
    End If
End Function

Sub CreateBiDirectionalLinksCustomTitleForCode()
    Dim oDoc As Object
    Dim oText As Object
    Dim oVCursor As Object
    Dim sSectionTitle As String
    Dim sCleanTitle As String
    Dim sCustomTitle As String
    Dim oBookmarks As Object
    Dim oTextCursor As Object
    Dim oInsertCursor As Object
    Dim oBookmark As Object
    Dim colonPos As Integer
    Dim sDocURL As String

    ' Initialize document objects.
    oDoc = ThisComponent
    oText = oDoc.Text
    oVCursor = oDoc.CurrentController.ViewCursor

    ' Get the text from the current selection.
    sSectionTitle = oVCursor.getString()

    ' Locate the first colon in the section title.
    colonPos = InStr(sSectionTitle, ":")
    If colonPos = 0 Then
        MsgBox "Please include a colon (:) in the code snippet.", 48, "Formatting Error"
        Exit Sub
    End If

    ' Remove the colon and trim spaces to get the clean title.
    sCleanTitle = Trim(Left(sSectionTitle, colonPos - 1))

    ' Prompt the user for a custom bookmark title.
    sCustomTitle = InputBox("Enter Bookmark Title", "Bookmark Title")
    If sCustomTitle = "" Then Exit Sub

    ' Construct bookmark names using the custom title.
    Dim sBookmarkMain As String
    Dim sBookmarkTOC As String
    sBookmarkMain = sCustomTitle
    sBookmarkTOC = sBookmarkMain & " Contents"

    ' Get the bookmarks container.
    oBookmarks = oDoc.getBookmarks()

    ' Check if the bookmark names already exist.
    If oBookmarks.hasByName(sBookmarkMain) Then
        MsgBox "A bookmark with the name '" & sBookmarkMain & "' already exists." & _
               Chr(13) & "Please choose a different bookmark title.", 48, "Bookmark Exists"
        Exit Sub
    End If

    If oBookmarks.hasByName(sBookmarkTOC) Then
        MsgBox "A bookmark with the name '" & sBookmarkTOC & "' already exists." & _
               Chr(13) & "Please choose a different bookmark title.", 48, "Bookmark Exists"
        Exit Sub
    End If

    ' Retrieve the full document URL.
    sDocURL = oDoc.URL
    If sDocURL = "" Then
        MsgBox "Document must be saved to create bi-directional navigation.", 48, "Error"
        Exit Sub
    End If

    ' --- Step 1: Replace colon with upward arrow and create bookmark ---
    ' Create a new text cursor starting at the beginning of the selection.
    oTextCursor = oText.createTextCursorByRange(oVCursor)
    oTextCursor.collapseToStart()

    ' Move the cursor to the colon's location.
    oTextCursor.goRight(colonPos - 1, False)
    ' Now select exactly one character (the colon).
    oTextCursor.goRight(1, True)

    ' Replace the colon with the arrow character.
    oTextCursor.String = "↑"

    ' Insert the bookmark for the main section at this location.
    oBookmark = oDoc.createInstance("com.sun.star.text.Bookmark")
    oBookmark.Name = sBookmarkMain
    oText.insertTextContent(oTextCursor, oBookmark, False)

    ' Apply hyperlink to point back to the TOC bookmark.
    oTextCursor.HyperLinkURL = sDocURL & "#" & sBookmarkTOC

    ' --- Step 2: Insert a navigation line above the heading ---
    oInsertCursor = oText.createTextCursorByRange(oVCursor)
    oInsertCursor.gotoStartOfParagraph(False)
    oText.insertString(oInsertCursor, sCleanTitle & Chr(13), False)

    ' --- Step 3: Create hyperlink from the navigation title to the main section ---
    oInsertCursor.goLeft(Len(sCleanTitle) + 1, True)
    oInsertCursor.HyperLinkURL = sDocURL & "#" & sBookmarkMain

    ' --- Step 4: Bookmark the navigation title for TOC use ---
    oBookmark = oDoc.createInstance("com.sun.star.text.Bookmark")
    oBookmark.Name = sBookmarkTOC
    oText.insertTextContent(oInsertCursor, oBookmark, False)

    MsgBox "✅ Bi-directional navigation created for code: " & sCleanTitle, 64, "Complete"
End Sub


Sub CreateHyperlinkContentsBookmark()
    Dim oDoc As Object
    Dim oVCursor As Object
    Dim oText As Object
    Dim oBookmark As Object
    Dim oBookmarks As Object
    Dim oTextCursor As Object
    Dim sHLTarget As String
    Dim sCleanTarget As String
    Dim sUserInput As String
    Dim sNewBookmark As String

    ' Get the current document, view cursor, and text object.
    oDoc = ThisComponent
    oVCursor = oDoc.getCurrentController().getViewCursor()
    oText = oDoc.Text

    ' Read the hyperlink target from the current selection.
    sHLTarget = oVCursor.HyperLinkURL

    If sHLTarget = "" Then
        MsgBox "The current selection does not have a hyperlink.", 48, "No Hyperlink Found"
        Exit Sub
    End If

    sCleanTarget = Split(sHLTarget, "#")(UBound(Split(sHLTarget, "#")))

    ' Prompt the user to confirm or edit the base text using an input box.
    sUserInput = InputBox("Confirm or edit the base text for the bookmark:", "Bookmark Base Text", sCleanTarget)
    If sUserInput = "" Then
        MsgBox "No text entered. Operation cancelled.", 48, "Cancelled"
        Exit Sub
    End If

    ' Append " Contents" to the text from the input box.
    sNewBookmark = sUserInput & " Contents"

    ' Retrieve the document's bookmarks collection.
    oBookmarks = oDoc.getBookmarks()
    If oBookmarks.hasByName(sNewBookmark) Then
        MsgBox "A bookmark with the name '" & sNewBookmark & "' already exists.", 48, "Bookmark Exists"
        Exit Sub
    End If

    ' Create the new bookmark instance.
    oBookmark = oDoc.createInstance("com.sun.star.text.Bookmark")
    oBookmark.Name = sNewBookmark

    ' Insert the new bookmark at the current selection as a character.
    oTextCursor = oText.createTextCursorByRange(oVCursor)
    oText.insertTextContent(oTextCursor, oBookmark, True)

    MsgBox "Bookmark '" & sNewBookmark & "' created.", 64, "Bookmark Created"
End Sub

Sub CreateCommaSeparatedListFromBulletsee()
    Dim oDoc As Object, oSelection As Object
    Dim oEnum As Object, oPara As Object
    Dim sText As String, sCleanText As String
    Dim pos As Integer, i As Integer, bFirst As Boolean
    Dim oCursor As Object, oText As Object
    Dim oHyperlink As Object

    oDoc = ThisComponent
    oSelection = oDoc.CurrentSelection
    oText = oDoc.Text

    If oSelection.getCount() = 0 Then
        MsgBox "Please select the nested bullet structure first."
        Exit Sub
    End If

    ' Create a cursor at the end of the selection so that new insertions are appended.
    oCursor = oText.createTextCursorByRange(oSelection.getByIndex(0).getStart())
    oCursor.gotoRange(oSelection.getByIndex(0).getEnd(), True)
    oCursor.collapseToEnd()

    bFirst = True

    For i = 0 To oSelection.getCount() - 1
        oEnum = oSelection.getByIndex(i).createEnumeration()
        While oEnum.hasMoreElements()
            oPara = oEnum.nextElement()
            sText = oPara.getString()

            pos = InStr(sText, ":")
            If pos > 0 Then
                sCleanText = Trim(Left(sText, pos - 1))
                sCleanText = RemoveBullet(sCleanText)

                ' Create hyperlink text field
                oHyperlink = oDoc.createInstance("com.sun.star.text.TextField.URL")
                oHyperlink.Representation = sCleanText
                oHyperlink.URL = "#Test"

                ' Insert comma separator if needed
                If Not bFirst Then
                    oText.insertString(oCursor, ", ", False)
                Else
                    bFirst = False
                End If

                ' Insert hyperlink text field
                oText.insertTextContent(oCursor, oHyperlink, False)

                ' Reposition the cursor to the end of the inserted hyperlink field.
                ' This moves the cursor outside the field so that further insertions are valid.
                oCursor.gotoRange(oHyperlink.Anchor, False)
                oCursor.collapseToEnd()
            End If
        Wend
    Next i

    MsgBox "Native hyperlinks with #Test target inserted"
End Sub

Function InputBoxWrapper(Optional Prompt As String, Optional Title As String, Optional Default As String) As String
    InputBoxWrapper = InputBox(Prompt, Title, Default)
End Function
import uno
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
from com.sun.star.awt import Rectangle  # Import the Rectangle class
from com.sun.star.awt.MessageBoxType import ERRORBOX, INFOBOX



def show_message(message, title="Message", boxtype=INFOBOX):
    """
    Displays a message box.
    boxtype: Use MessageBoxType enums such as INFOBOX or ERRORBOX.
    """
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.ServiceManager
    doc = XSCRIPTCONTEXT.getDocument()
    frame = doc.CurrentController.Frame
    container_window = frame.ContainerWindow
    toolkit = smgr.createInstance("com.sun.star.awt.Toolkit")
    box = toolkit.createMessageBox(container_window, boxtype, BUTTONS_OK, title, message)
    box.execute()


def get_input_with_default(prompt, title, default_value):
    """
    Prompts the user for input using a BASIC input box.
    If the user enters nothing (or only whitespace), the provided default_value is returned.
    """
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.ServiceManager
    script_provider = smgr.createInstanceWithContext(
        "com.sun.star.script.provider.MasterScriptProviderFactory", ctx
    ).createScriptProvider("")
    script = script_provider.getScript(
        "vnd.sun.star.script:Standard.Module1.InputBoxWrapper?language=Basic&location=application"
    )
    args = (f"{prompt}\nLeave blank to use default: {default_value}", title, "")
    result_tuple = script.invoke(args, (), ())
    if result_tuple and isinstance(result_tuple, tuple):
        result = result_tuple[0]
        if result and isinstance(result, str) and result.strip():
            return result.strip()
    return default_value


def bidirectional_link():
    """
    Creates bi-directional bookmarks and hyperlinks for a selected heading like "Text One:".
    - The main bookmark covers the complete title text ("Text One").
    - The colon receives a hyperlink.
    - A navigation line (the clean title) is inserted above the original paragraph.
    - The TOC bookmark is applied only to the inserted navigation text (without the trailing newline).
    """
    doc = XSCRIPTCONTEXT.getDocument()
    text = doc.Text
    view_cursor = doc.CurrentController.getViewCursor()

    # --- Step 0: Get and process the selected text ---
    selected_text = view_cursor.getString()
    colon_index = selected_text.find(":")
    if colon_index == -1:
        show_message("Please include a colon (:) in the selected text.", "Formatting Error", boxtype=ERRORBOX)
        return

    clean_title = selected_text[:colon_index].strip()
    section_number = get_input_with_default("Enter Section Number (e.g., 1)", "Section Number", "1")
    main_bookmark = f"Section {section_number} {clean_title}"
    toc_bookmark = main_bookmark + " Contents"

    bookmarks = doc.getBookmarks()
    if bookmarks.hasByName(main_bookmark) or bookmarks.hasByName(toc_bookmark):
        show_message("Bookmark name already exists. Please choose a different section number or title.",
                     "Bookmark Exists", boxtype=ERRORBOX)
        return

    # --- Step 1: Create the main bookmark covering the entire clean_title ---
    line_cursor = text.createTextCursorByRange(view_cursor.getStart())
    line_cursor.gotoStartOfParagraph(False)
    if not line_cursor.goRight(len(clean_title), True):
        show_message("Error selecting the title text.", "Error", boxtype=ERRORBOX)
        return

    main_bm_obj = doc.createInstance("com.sun.star.text.Bookmark")
    main_bm_obj.Name = main_bookmark
    text.insertTextContent(line_cursor, main_bm_obj, True)

    # --- Step 2: Apply hyperlink to the colon following the title ---
    colon_cursor = text.createTextCursorByRange(line_cursor.getStart())
    if colon_cursor.goRight(len(clean_title), False) and colon_cursor.goRight(1, True):
        if colon_cursor.getString() == ":":
            full_doc_url = doc.URL
            if not full_doc_url:
                show_message("Please save the document before running this macro.", "Save Required", boxtype=ERRORBOX)
                return
            colon_cursor.HyperLinkURL = full_doc_url + "#" + toc_bookmark
            colon_cursor.HyperLinkName = toc_bookmark
            colon_cursor.HyperLinkTarget = ""

    # --- Step 3: Insert the navigation line above the original paragraph ---
    # Get an insertion cursor at the start of the current paragraph.
    insert_cursor = text.createTextCursorByRange(view_cursor.getStart())
    insert_cursor.gotoStartOfParagraph(False)
    # Navigation string includes the clean title and a newline.
    navigation_line = clean_title + chr(13)
    text.insertString(insert_cursor, navigation_line, False)
    # Now, insert_cursor is at the end of the inserted text.

    # --- Step 4: Rewind the insertion cursor to select just the inserted text ---
    # Move the cursor left by the entire navigation_line so it selects the inserted text.
    if not insert_cursor.goLeft(len(navigation_line), True):
        show_message("Error selecting the inserted navigation text.", "Error", boxtype=ERRORBOX)
        return
    # Collapse the selection to its start.
    insert_cursor.collapseToStart()
    # Now move right exactly by the length of clean_title (excluding the newline)
    if not insert_cursor.goRight(len(clean_title), True):
        show_message("Error refining the selection for the navigation text.", "Error", boxtype=ERRORBOX)
        return

    full_doc_url = doc.URL
    if not full_doc_url:
        show_message("Please save the document before running this macro.",
                     "Save Required", boxtype=ERRORBOX)
        return
    # Apply the hyperlink to the inserted navigation text to link back to the main bookmark.
    insert_cursor.HyperLinkURL = full_doc_url + "#" + main_bookmark

    # --- Step 5: Apply the TOC bookmark over the inserted navigation text ---
    toc_bm_obj = doc.createInstance("com.sun.star.text.Bookmark")
    toc_bm_obj.Name = toc_bookmark
    text.insertTextContent(insert_cursor, toc_bm_obj, True)

    show_message(f"✅ Bi-directional link created for: {clean_title}", "Success", boxtype=INFOBOX)

def bidirectional_link_with_parent():
    """
    Creates bi-directional bookmarks and hyperlinks for a selected heading.
    This function performs the same steps as bidirectional_link but additionally:
      - Prompts the user for the parent bookmark name.
      - Prepends the entered parent bookmark to the selected clean title.
    For example, if the selected text is "Text Two:" and the user enters
    "Section 1 Text One", then the main bookmark becomes "Section 1 Text One Text Two"
    and the TOC bookmark becomes "Section 1 Text One Text Two Contents".
    """
    doc = XSCRIPTCONTEXT.getDocument()
    text = doc.Text
    view_cursor = doc.CurrentController.getViewCursor()

    # --- Step 0: Get and process the selected text ---
    selected_text = view_cursor.getString()
    colon_index = selected_text.find(":")
    if colon_index == -1:
        show_message("Please include a colon (:) in the selected text.",
                     "Formatting Error", boxtype=ERRORBOX)
        return

    # Extract clean text (without the colon and any trailing spaces)
    clean_title = selected_text[:colon_index].strip()

    # Prompt for the parent bookmark name
    parent_bm = get_input_with_default("Enter the name of the parent bookmark",
                                       "Parent Bookmark", "")
    if not parent_bm:
        show_message("Parent bookmark name cannot be empty.",
                     "Input Error", boxtype=ERRORBOX)
        return

    # Create the main and TOC bookmark names by prepending the parent bookmark
    main_bookmark = f"{parent_bm} {clean_title}"
    toc_bookmark = main_bookmark + " Contents"

    bookmarks = doc.getBookmarks()
    if bookmarks.hasByName(main_bookmark) or bookmarks.hasByName(toc_bookmark):
        show_message("Bookmark name already exists. Please choose a different parent name or title.",
                     "Bookmark Exists", boxtype=ERRORBOX)
        return

    # --- Step 1: Create the main bookmark covering the clean_title ---
    line_cursor = text.createTextCursorByRange(view_cursor.getStart())
    line_cursor.gotoStartOfParagraph(False)
    if not line_cursor.goRight(len(clean_title), True):
        show_message("Error selecting the title text.", "Error", boxtype=ERRORBOX)
        return

    main_bm_obj = doc.createInstance("com.sun.star.text.Bookmark")
    main_bm_obj.Name = main_bookmark
    text.insertTextContent(line_cursor, main_bm_obj, True)

    # --- Step 2: Apply hyperlink to the colon following the title ---
    colon_cursor = text.createTextCursorByRange(line_cursor.getStart())
    if colon_cursor.goRight(len(clean_title), False) and colon_cursor.goRight(1, True):
        if colon_cursor.getString() == ":":
            full_doc_url = doc.URL
            if not full_doc_url:
                show_message("Please save the document before running this macro.",
                             "Save Required", boxtype=ERRORBOX)
                return
            colon_cursor.HyperLinkURL = full_doc_url + "#" + toc_bookmark
            colon_cursor.HyperLinkName = toc_bookmark
            colon_cursor.HyperLinkTarget = ""

    # --- Step 3: Insert the navigation line above the original paragraph ---
    insert_cursor = text.createTextCursorByRange(view_cursor.getStart())
    insert_cursor.gotoStartOfParagraph(False)
    navigation_line = clean_title + chr(13)
    text.insertString(insert_cursor, navigation_line, False)

    # --- Step 4: Rewind the insertion cursor to select the inserted navigation text ---
    if not insert_cursor.goLeft(len(navigation_line), True):
        show_message("Error selecting the inserted navigation text.",
                     "Error", boxtype=ERRORBOX)
        return
    insert_cursor.collapseToStart()
    if not insert_cursor.goRight(len(clean_title), True):
        show_message("Error refining the selection for the navigation text.",
                     "Error", boxtype=ERRORBOX)
        return

    full_doc_url = doc.URL
    if not full_doc_url:
        show_message("Please save the document before running this macro.",
                     "Save Required", boxtype=ERRORBOX)
        return
    # Apply the hyperlink to the inserted navigation text (linking back to the main bookmark)
    insert_cursor.HyperLinkURL = full_doc_url + "#" + main_bookmark

    # --- Step 5: Apply the TOC bookmark over the inserted navigation text ---
    toc_bm_obj = doc.createInstance("com.sun.star.text.Bookmark")
    toc_bm_obj.Name = toc_bookmark
    text.insertTextContent(insert_cursor, toc_bm_obj, True)

    show_message(f"✅ Bi-directional link created for: {main_bookmark}",
                 "Success", boxtype=INFOBOX)


def get_parent_bookmark_from_user(default_value):
    """
    Prompts the user for a parent bookmark using an input dialog.
    If the user leaves the input blank, returns the provided default_value.
    """
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.getServiceManager()
    script_provider = smgr.createInstanceWithContext(
        "com.sun.star.script.provider.MasterScriptProviderFactory", ctx
    ).createScriptProvider("")
    script = script_provider.getScript(
        "vnd.sun.star.script:Standard.Module1.InputBoxWrapper?language=Basic&location=application"
    )
    # Note: The input message includes an explanation with the default_value.
    args = (f"Leave blank to use default:\n{default_value}", "Parent Bookmark", "")
    result_tuple = script.invoke(args, (), ())
    if result_tuple and isinstance(result_tuple, tuple):
        result = result_tuple[0]
        return result.strip() if result and isinstance(result, str) else default_value
    return default_value

def get_custom_bookmark_title():
    """
    Prompts the user to enter a custom bookmark title.
    Returns the trimmed result as a string, or an empty string if the user cancels.
    """
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.getServiceManager()
    script_provider = smgr.createInstanceWithContext(
        "com.sun.star.script.provider.MasterScriptProviderFactory", ctx
    ).createScriptProvider("")
    # The arguments are: prompt message, title, and default text.
    args = ("Enter Bookmark Title", "Bookmark Title", "")
    result_tuple = script_provider.getScript(
        "vnd.sun.star.script:Standard.Module1.InputBoxWrapper?language=Basic&location=application"
    ).invoke(args, (), ())
    if result_tuple and isinstance(result_tuple, tuple):
        result = result_tuple[0]
        if result and isinstance(result, str):
            return result.strip()
    return ""


def custom_bidirectional_link():
    """
        Creates bi-directional bookmarks and hyperlinks for a selected heading.
        This function performs the same steps as bidirectional_link but additionally:
          - Prompts the user for the parent bookmark name.
          - Prepends the entered parent bookmark to the selected clean title.
        For example, if the selected text is "Text Two:" and the user enters
        "Section 1 Text One", then the main bookmark becomes "Section 1 Text One Text Two"
        and the TOC bookmark becomes "Section 1 Text One Text Two Contents".
        """
    doc = XSCRIPTCONTEXT.getDocument()
    text = doc.Text
    view_cursor = doc.CurrentController.getViewCursor()

    # --- Step 0: Get and process the selected text ---
    selected_text = view_cursor.getString()
    colon_index = selected_text.find(":")
    if colon_index == -1:
        show_message("Please include a colon (:) in the selected text.",
                     "Formatting Error", boxtype=ERRORBOX)
        return

    # Extract clean text (without the colon and any trailing spaces)
    clean_title = selected_text[:colon_index].strip()

    # Prompt for the parent bookmark name
    bm = get_input_with_default("Enter the name of your bookmark",
                                       "Bookmark", clean_title)
    if not bm:
        show_message("Parent bookmark name cannot be empty.",
                     "Input Error", boxtype=ERRORBOX)
        return

    # Create the main and TOC bookmark names by prepending the parent bookmark
    main_bookmark = bm
    toc_bookmark = main_bookmark + " Contents"

    bookmarks = doc.getBookmarks()
    if bookmarks.hasByName(main_bookmark) or bookmarks.hasByName(toc_bookmark):
        show_message("Bookmark name already exists. Please choose a different parent name or title.",
                     "Bookmark Exists", boxtype=ERRORBOX)
        return

    # --- Step 1: Create the main bookmark covering the clean_title ---
    line_cursor = text.createTextCursorByRange(view_cursor.getStart())
    line_cursor.gotoStartOfParagraph(False)
    if not line_cursor.goRight(len(clean_title), True):
        show_message("Error selecting the title text.", "Error", boxtype=ERRORBOX)
        return

    main_bm_obj = doc.createInstance("com.sun.star.text.Bookmark")
    main_bm_obj.Name = main_bookmark
    text.insertTextContent(line_cursor, main_bm_obj, True)

    # --- Step 2: Apply hyperlink to the colon following the title ---
    colon_cursor = text.createTextCursorByRange(line_cursor.getStart())
    if colon_cursor.goRight(len(clean_title), False) and colon_cursor.goRight(1, True):
        if colon_cursor.getString() == ":":
            full_doc_url = doc.URL
            if not full_doc_url:
                show_message("Please save the document before running this macro.",
                             "Save Required", boxtype=ERRORBOX)
                return
            colon_cursor.HyperLinkURL = full_doc_url + "#" + toc_bookmark
            colon_cursor.HyperLinkName = toc_bookmark
            colon_cursor.HyperLinkTarget = ""

    # --- Step 3: Insert the navigation line above the original paragraph ---
    insert_cursor = text.createTextCursorByRange(view_cursor.getStart())
    insert_cursor.gotoStartOfParagraph(False)
    navigation_line = clean_title + chr(13)
    text.insertString(insert_cursor, navigation_line, False)

    # --- Step 4: Rewind the insertion cursor to select the inserted navigation text ---
    if not insert_cursor.goLeft(len(navigation_line), True):
        show_message("Error selecting the inserted navigation text.",
                     "Error", boxtype=ERRORBOX)
        return
    insert_cursor.collapseToStart()
    if not insert_cursor.goRight(len(clean_title), True):
        show_message("Error refining the selection for the navigation text.",
                     "Error", boxtype=ERRORBOX)
        return

    full_doc_url = doc.URL
    if not full_doc_url:
        show_message("Please save the document before running this macro.",
                     "Save Required", boxtype=ERRORBOX)
        return
    # Apply the hyperlink to the inserted navigation text (linking back to the main bookmark)
    insert_cursor.HyperLinkURL = full_doc_url + "#" + main_bookmark

    # --- Step 5: Apply the TOC bookmark over the inserted navigation text ---
    toc_bm_obj = doc.createInstance("com.sun.star.text.Bookmark")
    toc_bm_obj.Name = toc_bookmark
    text.insertTextContent(insert_cursor, toc_bm_obj, True)

    show_message(f"✅ Bi-directional link created for: {main_bookmark}",
                 "Success", boxtype=INFOBOX)


def get_custom_bookmark_title():
    """
    Prompts the user to enter a custom bookmark title using an input dialog.
    If the user leaves the input blank, returns an empty string.
    """
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.getServiceManager()
    script_provider = smgr.createInstanceWithContext(
        "com.sun.star.script.provider.MasterScriptProviderFactory", ctx
    ).createScriptProvider("")
    # Prompt: Message, Title, and default (empty)
    args = ("Enter Bookmark Title", "Bookmark Title", "")
    result_tuple = script_provider.getScript(
        "vnd.sun.star.script:Standard.Module1.InputBoxWrapper?language=Basic&location=application"
    ).invoke(args, (), ())
    if result_tuple and isinstance(result_tuple, tuple):
        result = result_tuple[0]
        if result and isinstance(result, str):
            return result.strip()
    return ""

def custom_bidirectional_link_for_code():
    """
    Creates bi-directional bookmarks and hyperlinks for a selected heading,
    similar to custom_bidirectional_link but with an up arrow in place of the colon.
    After extracting the title from the selected text (up to the colon),
    the function replaces the colon with an up arrow ("↑") and applies a hyperlink
    to that arrow. All other functionality remains the same.
    """
    doc = XSCRIPTCONTEXT.getDocument()
    text = doc.Text
    view_cursor = doc.CurrentController.getViewCursor()

    # --- Step 0: Get and process the selected text ---
    selected_text = view_cursor.getString()
    colon_index = selected_text.find(":")
    if colon_index == -1:
        show_message("Please include a colon (:) in the selected text.",
                     "Formatting Error", boxtype=ERRORBOX)
        return

    # Extract clean text (without the colon and any trailing spaces)
    clean_title = selected_text[:colon_index].strip()

    # Prompt for the parent bookmark name
    bm = get_input_with_default("Enter the name of your bookmark",
                                "Bookmark", clean_title)
    if not bm:
        show_message("Parent bookmark name cannot be empty.",
                     "Input Error", boxtype=ERRORBOX)
        return

    # Create the main and TOC bookmark names by prepending the parent bookmark
    main_bookmark = bm
    toc_bookmark = main_bookmark + " Contents"

    bookmarks = doc.getBookmarks()
    if bookmarks.hasByName(main_bookmark) or bookmarks.hasByName(toc_bookmark):
        show_message("Bookmark name already exists. Please choose a different parent name or title.",
                     "Bookmark Exists", boxtype=ERRORBOX)
        return

    # --- Step 1: Create the main bookmark covering the clean_title ---
    line_cursor = text.createTextCursorByRange(view_cursor.getStart())
    line_cursor.gotoStartOfParagraph(False)
    if not line_cursor.goRight(len(clean_title), True):
        show_message("Error selecting the title text.", "Error", boxtype=ERRORBOX)
        return

    main_bm_obj = doc.createInstance("com.sun.star.text.Bookmark")
    main_bm_obj.Name = main_bookmark
    text.insertTextContent(line_cursor, main_bm_obj, True)

    # --- Step 2: Replace the colon with an up arrow and apply hyperlink ---
    arrow_cursor = text.createTextCursorByRange(line_cursor.getStart())
    if arrow_cursor.goRight(len(clean_title), False) and arrow_cursor.goRight(1, True):
        if arrow_cursor.getString() == ":":
            # Replace the colon with the up arrow
            arrow_cursor.String = "↑"
            full_doc_url = doc.URL
            if not full_doc_url:
                show_message("Please save the document before running this macro.",
                             "Save Required", boxtype=ERRORBOX)
                return
            arrow_cursor.HyperLinkURL = full_doc_url + "#" + toc_bookmark
            arrow_cursor.HyperLinkName = toc_bookmark
            arrow_cursor.HyperLinkTarget = ""

    # --- Step 3: Insert the navigation line above the original paragraph ---
    insert_cursor = text.createTextCursorByRange(view_cursor.getStart())
    insert_cursor.gotoStartOfParagraph(False)
    navigation_line = clean_title + chr(13)
    text.insertString(insert_cursor, navigation_line, False)

    # --- Step 4: Rewind the insertion cursor to select the inserted navigation text ---
    if not insert_cursor.goLeft(len(navigation_line), True):
        show_message("Error selecting the inserted navigation text.",
                     "Error", boxtype=ERRORBOX)
        return
    insert_cursor.collapseToStart()
    if not insert_cursor.goRight(len(clean_title), True):
        show_message("Error refining the selection for the navigation text.",
                     "Error", boxtype=ERRORBOX)
        return

    full_doc_url = doc.URL
    if not full_doc_url:
        show_message("Please save the document before running this macro.",
                     "Save Required", boxtype=ERRORBOX)
        return
    # Apply the hyperlink to the inserted navigation text (linking back to the main bookmark)
    insert_cursor.HyperLinkURL = full_doc_url + "#" + main_bookmark

    # --- Step 5: Apply the TOC bookmark over the inserted navigation text ---
    toc_bm_obj = doc.createInstance("com.sun.star.text.Bookmark")
    toc_bm_obj.Name = toc_bookmark
    text.insertTextContent(insert_cursor, toc_bm_obj, True)

    show_message(f"✅ Bi-directional link created for: {main_bookmark}",
                 "Success", boxtype=INFOBOX)


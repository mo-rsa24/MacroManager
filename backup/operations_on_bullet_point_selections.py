"""
Summary
Helper Functions:

getParagraphsWithinRange(text_range): Filters and returns paragraph elements from a given text range.

get_parent_bookmark_from_user(default_value): Prompts the user to supply a base parent bookmark name.
--------------------------------------------------------------------------------------------------------
Main Functions for Bullet and Bookmark Operations:

- 1. identifyBulletLevelsInSelection(): Examines the current selection to determine the bullet list structure and prints their bullet levels.

- 2. insert_nested_bookmark_summary(): Creates a summary line (with hyperlinks) and adds bookmarks for each bullet based on its nesting level.

- 3. insert_nested_bookmark_summaries(): Extends the previous functionality by inserting additional summary bookmarks covering the title text, as well as enhanced hyperlinking.

Each of these functions works together to enable dynamic creation of nested bookmarks and hyperlinks within a document, particularly useful for structuring or navigating bullet-pointed lists in a document editing environment.


"""

def identifyBulletLevelsInSelection():
    """
1. Purpose:
This function processes the current selection in the document to identify which paragraphs are part of a bulleted or numbered list and then determines the level (or indentation) of each bullet.

2. How It Works:

2.1 Retrieve the Document and Selection:
It fetches the current document and the current selection. Since the selection could be a single range or multiple ranges (for instance, multiple disjoint selections), it handles both cases.

2.2 Extract Paragraphs:
It calls the helper function getParagraphsWithinRange for each text range in the selection to compile a complete list of paragraphs.

2.3 Determine and Print Bullet Levels:
For every paragraph in the list:

It attempts to access the "NumberingLevel" property.

For the first bullet (index 0), it assigns level 0 by default. For subsequent paragraphs, it adds one to the returned numbering level (i.e., it uses para.getPropertyValue("NumberingLevel") + 1).

It prints out each paragraph's text alongside its computed bullet level.

3. Return Value:
It returns a list of bullet levels corresponding to each paragraph in the selection.
    """
    doc = XSCRIPTCONTEXT.getDocument()
    selection = doc.getCurrentSelection()

    # This list will hold all paragraphs from the selection.
    paragraphs = []

    # The selection might be a collection (multiple ranges) or a single text range.
    try:
        # If the selection supports getCount, it is a collection of ranges.
        count = selection.getCount()
        for i in range(count):
            sel_range = selection.getByIndex(i)
            paragraphs.extend(getParagraphsWithinRange(sel_range))
    except AttributeError:
        # Otherwise, it's a single text range.
        paragraphs = getParagraphsWithinRange(selection)

    if not paragraphs:
        print("No paragraphs found in the selection.")
        return
    levels = []
    for (index, para) in enumerate(paragraphs):
        try:
            # Try to get the 'NumberingLevel' property.
            if index == 0:
                level = 0
            else:
                level = para.getPropertyValue("NumberingLevel") + 1
            levels.append(level)
            print("Paragraph:", para.getString())
            print("Bullet Level:", level)
            print("-----------")
        except Exception:
            # If the property is not available, the paragraph is likely not part of a list.
            pass
    return levels

def getParagraphsWithinRange(text_range):
    """
1. Purpose:
Given a text range, this function returns a list of paragraphs
    (elements that support the "com.sun.star.text.Paragraph" service)
    that are contained within the given range.

    This helper function extracts all paragraphs from a given text range that are recognized as paragraph elements in the document.

2. How It Works:

It creates an enumeration of all elements within the specified text range.

It iterates over each element and checks if it supports the "com.sun.star.text.Paragraph" service.

All matching elements (i.e., paragraphs) are collected in a list and returned.

This function is critical for filtering the selection to only those parts that are valid paragraphs for bullet level assessment.

    """
    paragraphs = []
    enum = text_range.createEnumeration()
    while enum.hasMoreElements():
        element = enum.nextElement()
        if element.supportsService("com.sun.star.text.Paragraph"):
            paragraphs.append(element)
    return paragraphs

def insert_nested_bookmark_summary():
    """
1. Purpose:
This function is designed to work on a selected block of bullet-point text.
It creates a summary line of bullet titles (extracted from the selected lines) and inserts nested bookmarks and hyperlinks into the document based on the bullet nesting hierarchy.

2. How It Works:

2.1 Selection Validation:
It first checks if the current view cursor selection is valid (i.e., not collapsed) and splits the selected text into individual lines.

2.2 Bullet Level Acquisition:
It calls identifyBulletLevelsInSelection() to retrieve the bullet levels for each line. It also ensures there is a one-to-one match between the number of bullet levels and the number of text lines.

2.3 Establishing a Base Bookmark:
It requires that the first (root) line includes a colon “:” to extract the root title. This title is used to create a base parent bookmark. The function then calls get_parent_bookmark_from_user() to allow the user to confirm or modify this base value.

2.4 Building a Nested Bookmark Chain:
It iterates over each line, extracting titles (the text before any colon) and builds nested bookmark names using the bullet level. For each bullet:

Top-level bullets prepend the base parent.

Nested bullets concatenate the parent bookmark with the bullet’s title.

It maintains a dictionary (parent_chain) to track the current parent at each level.

2.5 Inserting the Summary Line and Hyperlinks:
A summary line (a comma-separated list of titles) is inserted right above the selected bullet text.
Then, for each title in the summary, a hyperlink is created linking the title back to its corresponding nested bookmark.

2.6 Inserting Bookmarks into the Document:
It iterates through each bullet line (after the summary) and:

Inserts a bookmark at the start of the line using the computed nested bookmark string.

Optionally, if a colon (":") is detected in the text immediately following the title, it adds a hyperlink on the colon that points to a “Contents” version of the bookmark.

Outcome:
The function prints a success message upon completion, indicating that the nested bookmarks and hierarchy have been inserted.
    """
    import re

    doc = XSCRIPTCONTEXT.getDocument()
    text = doc.Text
    view_cursor = doc.CurrentController.getViewCursor()

    if view_cursor.isCollapsed():
        print("Please select bullet-point text.")
        return

    # Retrieve the selected text and split into individual lines.
    selected_string = view_cursor.getString()
    lines = selected_string.splitlines()

    if not lines:
        print("No text selected.")
        return

    # Obtain bullet levels for each paragraph from the existing function.
    levels = identifyBulletLevelsInSelection()
    if len(levels) != len(lines):
        print("Mismatch between bullet levels and selected lines. Make sure the selection includes only the intended bullet paragraphs.")
        return

    # The first line (the root) must contain a colon.
    if ":" not in lines[0]:
        print("Root title (first line) must contain a colon.")
        return

    # Extract the root title and let the user confirm the base parent bookmark.
    root_title = lines[0].split(":")[0].strip()
    base_parent = get_parent_bookmark_from_user(f"Section 1 {root_title}")

    # Build nested bookmark names based on bullet levels.
    # We use a dictionary to hold the current parent bookmark for each level.
    parent_chain = {}
    bookmarks = []  # full bookmark for each bullet line
    titles = []     # extracted title from each line

    for i, line in enumerate(lines):
        # Get the text before the colon as the title.
        if ":" in line:
            title = line.split(":")[0].strip()
        else:
            title = line.strip()
        titles.append(title)

        current_level = levels[i]
        if current_level == 0:
            # For a top-level bullet, prepend the base parent.
            full_bm = base_parent
        else:
            # For nested bullets, look up the parent's full bookmark.
            parent_bm = parent_chain.get(current_level - 1, base_parent)
            full_bm = parent_bm + " " + title
        # Update the chain so that subsequent lower-level items can use this one.
        parent_chain[current_level] = full_bm
        # If there are any deeper levels already in the chain, clear them
        keys_to_remove = [k for k in parent_chain if k > current_level]
        for k in keys_to_remove:
            del parent_chain[k]
        bookmarks.append(full_bm)

    # Insert a summary line above the selection that lists all titles.
    summary_cursor = text.createTextCursorByRange(view_cursor.getStart())
    summary_cursor.gotoStartOfParagraph(False)
    summary_line = ", ".join(titles)
    text.insertString(summary_cursor, summary_line + "\n", False)
    summary_cursor.goLeft(len(summary_line) + 1, True)
    summary_cursor.collapseToStart()

    full_doc_url = doc.URL

    # Insert hyperlinks in the summary line.
    for i, title in enumerate(titles):
        hyperlink_url = full_doc_url + "#" + bookmarks[i]
        if not summary_cursor.goRight(len(title), True):
            break
        summary_cursor.HyperLinkURL = hyperlink_url
        summary_cursor.HyperLinkName = title
        summary_cursor.HyperLinkTarget = ""
        summary_cursor.collapseToEnd()

        if i < len(titles) - 1:
            summary_cursor.goRight(2, False)

    # For each bullet line, add a bookmark with the computed nested bookmark.
    current_range = text.createTextCursorByRange(view_cursor.getStart())
    current_range.gotoNextParagraph(False)

    for i in range(1,len(titles)):
        title = titles[i]
        full_bm = bookmarks[i]
        line_cursor = text.createTextCursorByRange(current_range)
        line_cursor.gotoStartOfParagraph(False)
        if not line_cursor.goRight(len(title), True):
            continue

        # Create and insert the bookmark.
        bookmark = doc.createInstance("com.sun.star.text.Bookmark")
        bookmark.Name = full_bm
        text.insertTextContent(line_cursor, bookmark, True)

        # Optionally add a hyperlink on the colon to point to a "Contents" version.
        colon_cursor = text.createTextCursorByRange(line_cursor.getStart())
        if colon_cursor.goRight(len(title), False) and colon_cursor.goRight(1, True):
            if colon_cursor.getString() == ":":
                colon_cursor.HyperLinkURL = full_doc_url + "#" + full_bm + " Contents"
                colon_cursor.HyperLinkName = full_bm + " Contents"
                colon_cursor.HyperLinkTarget = ""

        current_range = text.createTextCursorByRange(line_cursor.getEnd())
        current_range.gotoNextParagraph(False)

    print("✅ Nested bookmarks with hierarchy inserted.")


def get_parent_bookmark_from_user(default_value):

    """
1. Purpose:
This helper function interacts with the user to obtain a desired parent bookmark name.
It ensures that the nested bookmark hierarchy has a proper starting point as provided by the user or falls back to a default.

2. How It Works:

It uses the document’s component context and service manager to create a script provider.

It then calls an input dialog (via an external Basic script) that prompts the user with a message including the default bookmark value.

If the user provides a non-empty response, it strips and returns that response; otherwise, it returns the default value.

This function is essential in allowing some degree of user customizability for the base bookmark name used in nested bookmark generation.

    """
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.ServiceManager

    script_provider = smgr.createInstanceWithContext(
        "com.sun.star.script.provider.MasterScriptProviderFactory", ctx
    ).createScriptProvider("")

    script = script_provider.getScript(
        "vnd.sun.star.script:Standard.Module1.InputBoxWrapper?language=Basic&location=application"
    )

    args = (f"Leave blank to use default:\n{default_value}", "Parent Bookmark", "")
    result_tuple = script.invoke(args, (), ())

    # Extract string from tuple
    if result_tuple and isinstance(result_tuple, tuple):
        result = result_tuple[0]
        return result.strip() if result and isinstance(result, str) else default_value

    return default_value

def insert_nested_bookmark_summaries():
    """
    Shortcut: Ctrl + Alt + N
    1. Purpose:
    This function is a variant of insert_nested_bookmark_summary().
    It performs similar operations by creating nested bookmarks based on the bullet hierarchy, but adds additional functionality by inserting extra summary bookmarks along with the summary line.

    2. How It Works:

    2.1 Selection and Bullet Extraction:
    Like its counterpart, it validates that bullet point text is selected and splits the text into lines.
    It then retrieves bullet levels using identifyBulletLevelsInSelection() and checks for the presence of a colon in the root line.

    2.2 Setting the Base Parent:
    It extracts a root title from the first line and again calls get_parent_bookmark_from_user() to determine the base parent bookmark.

    2.3 Building the Nested Bookmark Chain:
    Using the same parent-chain approach, it computes nested bookmark names for each bullet line, ensuring the correct concatenation of parent bookmarks with bullet titles.

    2.3 Inserting a Summary Line with Enhanced Hyperlinks:

    It inserts a summary line above the selection, but uses a different separator ("| ") between titles.
    A fresh cursor is used to process the summary line, and for each title:
        - A hyperlink is applied that links to the corresponding nested bookmark.
        - An additional bookmark is inserted that covers the entire title text. This extra bookmark (named as a “Contents” version) allows for an alternative navigation link.

    2.4 Bookmark Insertion for Each Bullet Item:
    The function then goes through each bullet item (after the root) and:
        - Inserts a primary bookmark for the bullet line.
        = Optionally adds a hyperlink on the colon if present to point to the “Contents” version of the bookmark.

    3. Outcome:
    It prints a confirmation message indicating that nested bookmarks with hierarchy and additional summary bookmarks have been successfully inserted.
    """

    doc = XSCRIPTCONTEXT.getDocument()
    text = doc.Text
    view_cursor = doc.CurrentController.getViewCursor()

    if view_cursor.isCollapsed():
        print("Please select bullet-point text.")
        return

    # Retrieve the selected text and split into individual lines.
    selected_string = view_cursor.getString()
    lines = selected_string.splitlines()

    if not lines:
        print("No text selected.")
        return

    # Obtain bullet levels for each paragraph from the existing function.
    levels = identifyBulletLevelsInSelection()
    if len(levels) != len(lines):
        print(
            "Mismatch between bullet levels and selected lines. Make sure the selection includes only the intended bullet paragraphs.")
        return

    # The first line (the root) must contain a colon.
    if ":" not in lines[0]:
        print("Root title (first line) must contain a colon.")
        return

    # Extract the root title and let the user confirm the base parent bookmark.
    root_title = lines[0].split(":")[0].strip()
    base_parent = get_parent_bookmark_from_user(f"Section 1 {root_title}")

    # Build nested bookmark names based on bullet levels.
    parent_chain = {}
    bookmarks = []  # full bookmark for each bullet line
    titles = []  # extracted title from each line

    for i, line in enumerate(lines):
        # Get the text before the colon as the title.
        if ":" in line:
            title = line.split(":")[0].strip()
        else:
            title = line.strip()
        titles.append(title)

        current_level = levels[i]
        if current_level == 0:
            # For a top-level bullet, prepend the base parent.
            full_bm = base_parent
        else:
            # For nested bullets, look up the parent's full bookmark.
            parent_bm = parent_chain.get(current_level - 1, base_parent)
            full_bm = parent_bm + " " + title
        # Update the chain so that subsequent lower-level items can use this one.
        parent_chain[current_level] = full_bm
        # If there are any deeper levels already in the chain, clear them.
        keys_to_remove = [k for k in parent_chain if k > current_level]
        for k in keys_to_remove:
            del parent_chain[k]
        bookmarks.append(full_bm)

    # Insert a summary line above the selection that lists all titles.
    summary_cursor = text.createTextCursorByRange(view_cursor.getStart())
    summary_cursor.gotoStartOfParagraph(False)
    summary_line = "| ".join(titles)
    text.insertString(summary_cursor, summary_line + "\n", False)
    summary_cursor.goLeft(len(summary_line) + 1, True)
    summary_cursor.collapseToStart()

    full_doc_url = doc.URL

    # Insert hyperlinks in the summary line and add additional bookmarks spanning the entire title text.
    # Create a fresh cursor starting at the beginning of the inserted summary line.
    summary_line_cursor = text.createTextCursorByRange(summary_cursor.getStart())

    for i, title in enumerate(titles):
        # Create a new cursor for this title's text.
        title_cursor = text.createTextCursorByRange(summary_line_cursor.getStart())
        if not title_cursor.goRight(len(title), True):
            continue

        hyperlink_url = full_doc_url + "#" + bookmarks[i]
        # Apply the hyperlink for the title.
        title_cursor.HyperLinkURL = hyperlink_url
        title_cursor.HyperLinkName = title
        title_cursor.HyperLinkTarget = ""

        # Add an additional bookmark that covers the entire title text.
        additional_bookmark = doc.createInstance("com.sun.star.text.Bookmark")
        additional_bookmark.Name = bookmarks[i] + " Contents"
        text.insertTextContent(title_cursor, additional_bookmark, True)

        # Move summary_line_cursor past the current title and the following comma and space if present.
        summary_line_cursor.goRight(len(title), False)
        if i < len(titles) - 1:
            summary_line_cursor.goRight(2, False)

    # For each bullet line (excluding the first root), add a bookmark with the computed nested bookmark.
    current_range = text.createTextCursorByRange(view_cursor.getStart())
    current_range.gotoNextParagraph(False)

    for i in range(1, len(titles)):
        title = titles[i]
        full_bm = bookmarks[i]
        line_cursor = text.createTextCursorByRange(current_range)
        line_cursor.gotoStartOfParagraph(False)
        if not line_cursor.goRight(len(title), True):
            continue

        # Create and insert the primary bookmark for the bullet line.
        bookmark = doc.createInstance("com.sun.star.text.Bookmark")
        bookmark.Name = full_bm
        text.insertTextContent(line_cursor, bookmark, True)

        # Optionally add a hyperlink on the colon to point to a "Contents" version.
        colon_cursor = text.createTextCursorByRange(line_cursor.getStart())
        if colon_cursor.goRight(len(title), False) and colon_cursor.goRight(1, True):
            if colon_cursor.getString() == ":":
                colon_cursor.HyperLinkURL = full_doc_url + "#" + full_bm + " Contents"
                colon_cursor.HyperLinkName = full_bm + " Contents"
                colon_cursor.HyperLinkTarget = ""
        current_range = text.createTextCursorByRange(line_cursor.getEnd())
        current_range.gotoNextParagraph(False)

    print("✅ Nested bookmarks with hierarchy and additional summary bookmarks inserted.")






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
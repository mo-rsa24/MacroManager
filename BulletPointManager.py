import uno
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
from com.sun.star.awt.MessageBoxType import ERRORBOX, INFOBOX
from com.sun.star.awt.FontWeight import BOLD
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

class BulletPointManager:
    def __init__(self, doc=None):
        # Use the provided document or get it from the global XSCRIPTCONTEXT.
        self.doc = doc if doc is not None else XSCRIPTCONTEXT.getDocument()
        self.text = self.doc.Text
        self.controller = self.doc.CurrentController
        self.view_cursor = self.controller.getViewCursor()
        self.ctx = XSCRIPTCONTEXT.getComponentContext()
        self.smgr = self.ctx.ServiceManager

    def show_message(self, message, title="Message", boxtype=INFOBOX):
        """
        Displays a message box.
        """
        frame = self.doc.CurrentController.Frame
        container_window = frame.ContainerWindow
        toolkit = self.smgr.createInstance("com.sun.star.awt.Toolkit")
        box = toolkit.createMessageBox(container_window, boxtype, BUTTONS_OK, title, message)
        box.execute()

    def get_paragraphs_within_range(self, text_range):
        """
        Returns a list of paragraphs (elements that support the
        "com.sun.star.text.Paragraph" service) contained in the given text_range.
        """
        paragraphs = []
        enum = text_range.createEnumeration()
        while enum.hasMoreElements():
            element = enum.nextElement()
            if element.supportsService("com.sun.star.text.Paragraph"):
                paragraphs.append(element)
        return paragraphs

    def identify_bullet_levels(self):
        """
        Iterates through the paragraphs of the current selection and
        returns a list of bullet levels (with the first paragraph level set to 0).
        Also prints out each paragraph and its bullet level.
        """
        selection = self.doc.getCurrentSelection()
        paragraphs = []
        # The selection may be a collection or a single range.
        try:
            count = selection.getCount()
            for i in range(count):
                sel_range = selection.getByIndex(i)
                paragraphs.extend(self.get_paragraphs_within_range(sel_range))
        except AttributeError:
            paragraphs = self.get_paragraphs_within_range(selection)

        if not paragraphs:
            print("No paragraphs found in the selection.")
            return []

        levels = []
        for index, para in enumerate(paragraphs):
            try:
                # Set the first paragraph level to 0; adjust subsequent bullet levels.
                level = 0 if index == 0 else para.getPropertyValue("NumberingLevel") + 1
                levels.append(level)
                print("Paragraph:", para.getString())
                print("Bullet Level:", level)
                print("-----------")
            except Exception:
                # If the property is not available, skip this paragraph.
                pass
        return levels

    def get_selection_lines(self):
        """
        Returns the selected text split into a list of lines.
        """
        selected_string = self.view_cursor.getString()
        return selected_string.splitlines()

    def get_parent_bookmark_from_user(self, default_value):
        """
        Prompts the user for a parent bookmark using an input dialog.
        If the user leaves the input blank, returns the provided default_value.
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
        if result_tuple and isinstance(result_tuple, tuple):
            result = result_tuple[0]
            return result.strip() if result and isinstance(result, str) else default_value
        return default_value

    def build_bookmark_chain(self, lines, levels, base_parent):
        """
        Given selection lines and their corresponding bullet levels,
        builds and returns a tuple (titles, bookmarks) where:
         - titles is a list of extracted titles (text before colon, or the whole line)
         - bookmarks is the corresponding fully qualified nested bookmark name.
        """
        parent_chain = {}
        bookmarks = []
        titles = []
        for i, line in enumerate(lines):
            if ":" not in line:
                continue
            title = line.split(":")[0].strip()
            titles.append(title)
            current_level = levels[i]
            if current_level == 0:
                full_bm = base_parent
            else:
                parent_bm = parent_chain.get(current_level - 1, base_parent)
                full_bm = parent_bm + " " + title
            parent_chain[current_level] = full_bm
            # Clear any deeper levels from the chain.
            for k in list(parent_chain.keys()):
                if k > current_level:
                    del parent_chain[k]
            bookmarks.append(full_bm)
        return titles, bookmarks

    def insert_summary_line(self, titles, bookmarks, separator=", ", add_extra_bookmarks=False):
        """
        Inserts a summary line above the current selection that lists all titles,
        and applies hyperlinks linking to the corresponding bookmarks.
        Optionally, inserts additional bookmarks on the summary text.
        """
        summary_text = separator.join(titles)
        summary_cursor = self.text.createTextCursorByRange(self.view_cursor.getStart())
        summary_cursor.gotoStartOfParagraph(False)
        self.text.insertString(summary_cursor, summary_text + "\n", False)
        summary_cursor.goLeft(len(summary_text) + 1, True)
        summary_cursor.collapseToStart()

        full_doc_url = self.doc.URL

        # Apply hyperlinks to each title within the summary line.
        summary_line_cursor = self.text.createTextCursorByRange(summary_cursor.getStart())
        for i, title in enumerate(titles):
            title_cursor = self.text.createTextCursorByRange(summary_line_cursor.getStart())
            if not title_cursor.goRight(len(title), True):
                continue
            hyperlink_url = full_doc_url + "#" + bookmarks[i]
            title_cursor.HyperLinkURL = hyperlink_url
            title_cursor.HyperLinkName = title
            title_cursor.HyperLinkTarget = ""
            if add_extra_bookmarks:
                # For extended summaries, add an additional bookmark.
                additional_bookmark = self.doc.createInstance("com.sun.star.text.Bookmark")
                additional_bookmark.Name = bookmarks[i] + " Contents"
                self.text.insertTextContent(title_cursor, additional_bookmark, True)
            # Move the cursor past the current title and separator.
            title_length = len(title)
            summary_line_cursor.goRight(title_length, False)
            if i < len(titles) - 1:
                summary_line_cursor.goRight(len(separator), False)

    def insert_bullet_bookmarks(self, titles, bookmarks):
        """
        For each bullet (except the root), this version uses a while loop to search for the correct paragraph
        matching the expected title. Only when the paragraph text (obtained via a temporary cursor)
        begins with the expected title and contains a colon is the bookmark inserted.
        """
        full_doc_url = self.doc.URL
        # Start from the paragraph after the summary (assumed to be immediately after the view cursor position).
        current_cursor = self.text.createTextCursorByRange(self.view_cursor.getStart())
        current_cursor.gotoNextParagraph(False)

        # Process each title starting from index 1 (skip root)
        for i in range(1, len(titles)):
            expected_title = titles[i]
            full_bm = bookmarks[i]
            found = False
            attempts = 0

            # Try up to a maximum of 10 paragraphs to find the expected text
            while attempts < 10:
                # Create a temporary cursor for the current paragraph so we can inspect its full text.
                temp_cursor = self.text.createTextCursorByRange(current_cursor)
                temp_cursor.gotoEndOfParagraph(True)
                para_text = temp_cursor.getString().strip()

                # If no meaningful text exists, move on
                if not para_text:
                    if not current_cursor.gotoNextParagraph(False):
                        break
                    attempts += 1
                    continue

                # If the paragraph does not contain a colon, it’s not eligible for processing,
                # so move to the next paragraph.
                if ":" not in para_text:
                    if not current_cursor.gotoNextParagraph(False):
                        break
                    attempts += 1
                    continue

                # Now, check if the paragraph starts with the expected title.
                if para_text.startswith(expected_title):
                    found = True
                    break
                else:
                    # Otherwise, move to the next paragraph and try again.
                    if not current_cursor.gotoNextParagraph(False):
                        break
                    attempts += 1

            if not found:
                # If after several attempts no matching paragraph is found, skip to the next title.
                continue

            # At this point, current_cursor points to the paragraph matching expected_title.
            # Create a new cursor for that paragraph for precise bookmarking.
            line_cursor = self.text.createTextCursorByRange(current_cursor)
            line_cursor.gotoStartOfParagraph(False)
            if not line_cursor.goRight(len(expected_title), True):
                continue  # If we cannot cover the expected title, skip this bullet.

            # Insert the bookmark over the expected title.
            bookmark = self.doc.createInstance("com.sun.star.text.Bookmark")
            bookmark.Name = full_bm
            self.text.insertTextContent(line_cursor, bookmark, True)

            # Create another cursor to check if a colon immediately follows the expected title.
            colon_cursor = self.text.createTextCursorByRange(line_cursor.getStart())
            if colon_cursor.goRight(len(expected_title), False) and colon_cursor.goRight(1, True):
                if colon_cursor.getString() == ":":
                    colon_cursor.HyperLinkURL = full_doc_url + "#" + full_bm + " Contents"
                    colon_cursor.HyperLinkName = full_bm + " Contents"
                    colon_cursor.HyperLinkTarget = ""

            # Move the main cursor past this paragraph for the next iteration.
            current_cursor.gotoNextParagraph(False)

    def propagate_title_character_style(self):
        """
        Reads the parent's bullet title portion (exact text run) to retrieve its
        character style name, then assigns that style to the child's title text.
        """

        selection = self.doc.getCurrentSelection()
        paragraphs = []
        try:
            count = selection.getCount()
            for i in range(count):
                sel_range = selection.getByIndex(i)
                paragraphs.extend(self.get_paragraphs_within_range(sel_range))
        except AttributeError:
            paragraphs = self.get_paragraphs_within_range(selection)

        if not paragraphs:
            self.show_message("No bullet paragraphs found in the selection.", "Error", boxtype=ERRORBOX)
            return

        style_stack = {}

        for idx, para in enumerate(paragraphs):
            full_text = para.getString()
            if ":" not in full_text:
                # not a bullet with a "title:"
                continue

            # bullet "level" (like your color example)
            try:
                level = 0 if idx == 0 else para.getPropertyValue("NumberingLevel") + 1
            except Exception:
                level = 0

            # Extract text before the colon => "title"
            title = full_text.split(":", 1)[0].strip()

            # Create a cursor for just that "title" portion
            title_cursor = self.text.createTextCursorByRange(para)
            title_cursor.gotoStartOfParagraph(False)
            if not title_cursor.goRight(len(title), True):
                continue

            if level == 0:
                # "Parent" bullet. Grab this bullet's style (exact text run).
                parent_style_name = title_cursor.getPropertyValue("CharStyleName")
                if parent_style_name and parent_style_name != "No Character Style":
                    style_stack[level] = parent_style_name
            else:
                # A nested bullet => find closest parent’s style
                inherited_style = None
                for lvl in range(level - 1, -1, -1):
                    if lvl in style_stack:
                        inherited_style = style_stack[lvl]
                        break

                # If the parent has a real style name, apply it to the child
                if inherited_style and inherited_style != "No Character Style":
                    title_cursor.setPropertyValue("CharStyleName", inherited_style)
                    title_cursor.setPropertyValue("CharWeight", BOLD)
                    # So deeper children can inherit from us
                    style_stack[level] = inherited_style

            # Clear out deeper levels from style_stack if we moved back up
            for deeper_lvl in list(style_stack.keys()):
                if deeper_lvl > level:
                    del style_stack[deeper_lvl]

        self.show_message("Character styles propagated to nested bullet titles.",
                          "Style Propagation Success")

    def insert_parent_bookmark_hyperlink(self, titles, bookmarks):
        """
        Locates the *parent* bullet (titles[0]) in the current selection,
        inserts a bookmark on its text, and hyperlinks the colon.
        """
        if not titles:
            return

        parent_title = titles[0]
        parent_bookmark = bookmarks[0]

        # Start searching from the view cursor (which presumably
        # points to the start of the bullet selection).
        cur = self.text.createTextCursorByRange(self.view_cursor.getStart())

        found = False
        attempts = 0
        while attempts < 10:
            tmp_cursor = self.text.createTextCursorByRange(cur)
            tmp_cursor.gotoEndOfParagraph(True)
            p_text = tmp_cursor.getString().strip()

            if p_text and ":" in p_text and p_text.startswith(parent_title):
                found = True
                break

            if not cur.gotoNextParagraph(False):
                break
            attempts += 1

        if not found:
            # If you like, you can show a warning or just return
            print(f"Could not find parent bullet titled: {parent_title}")
            return

        # Insert the bookmark over the bullet title portion
        line_cursor = self.text.createTextCursorByRange(cur)
        line_cursor.gotoStartOfParagraph(False)
        if line_cursor.goRight(len(parent_title), True):
            bm = self.doc.createInstance("com.sun.star.text.Bookmark")
            bm.Name = parent_bookmark
            self.text.insertTextContent(line_cursor, bm, True)

            # Optional: hyperlink the colon
            colon_cursor = self.text.createTextCursorByRange(line_cursor.getStart())
            if colon_cursor.goRight(len(parent_title), False) and colon_cursor.goRight(1, True):
                if colon_cursor.getString() == ":":
                    doc_url = self.doc.URL
                    colon_cursor.HyperLinkURL = doc_url + "#" + parent_bookmark + " Contents"
                    colon_cursor.HyperLinkName = parent_bookmark + " Contents"
                    colon_cursor.HyperLinkTarget = ""

    def process_nested_bookmark_summary(self, separator=", ", add_extra_bookmarks=False):
        """
        Main template method that performs the following steps:
         1. Validates the selection.
         2. Extracts the selection lines.
         3. Determines bullet levels.
         4. Verifies that the first line contains a colon.
         5. Obtains a base parent bookmark from the user.
         6. Builds the bookmark chain (titles & full bookmark names).
         7. Inserts the summary line with hyperlinks.
         8. Adds bullet bookmarks to each bullet paragraph.
        The separator and add_extra_bookmarks flag allow you to adjust the behavior
        (for example, basic vs. extended summary).
        """
        if self.view_cursor.isCollapsed():
            print("Please select bullet-point text.")
            return

        lines = self.get_selection_lines()
        if not lines:
            print("No text selected.")
            self.show_message(
                "No text selected.",
                "Formatting Error", boxtype=ERRORBOX)
            return

        levels = self.identify_bullet_levels()
        if len(levels) != len(lines):
            print("Mismatch between bullet levels and selected lines. Make sure the selection includes only the intended bullet paragraphs.")
            self.show_message("Mismatch between bullet levels and selected lines. Make sure the selection includes only the intended bullet paragraphs.",
                              "Formatting Error", boxtype=ERRORBOX)
            return

        if ":" not in lines[0]:
            print("Root title (first line) must contain a colon.")
            self.show_message("Root title (first line) must contain a colon.",
                              "Formatting Error", boxtype=ERRORBOX)
            return

        root_title = lines[0].split(":")[0].strip()
        base_parent = self.get_parent_bookmark_from_user(f"Section 1 {root_title}")
        titles, bookmarks = self.build_bookmark_chain(lines, levels, base_parent)
        self.insert_parent_bookmark_hyperlink(titles, bookmarks)
        self.insert_summary_line(titles, bookmarks, separator, add_extra_bookmarks)
        self.insert_bullet_bookmarks(titles, bookmarks)

        if add_extra_bookmarks:
            self.show_message(
                "✅ Nested bookmarks with hierarchy and additional summary bookmarks inserted.",
                "Nested Bookmarks Success", boxtype=INFOBOX)
            print("✅ Nested bookmarks with hierarchy and additional summary bookmarks inserted.")
        else:
            self.show_message(
                "✅ Nested bookmarks with hierarchy inserted.",
                "Nested Bookmarks Success", boxtype=INFOBOX)
            print("✅ Nested bookmarks with hierarchy inserted.")


# Module-level API functions to maintain existing interface.

def identifyBulletLevelsInSelection():
    manager = BulletPointManager()
    return manager.identify_bullet_levels()


def insert_nested_bookmark_summary():
    manager = BulletPointManager()
    manager.process_nested_bookmark_summary(separator=", ", add_extra_bookmarks=False)


def insert_nested_bookmark_summaries():
    """
    Function:
        - Main template method that performs the following steps:
         1. Validates the selection.
         2. Extracts the selection lines.
         3. Determines bullet levels.
         4. Verifies that the first line contains a colon.
         5. Obtains a base parent bookmark from the user.
         6. Builds the bookmark chain (titles & full bookmark names).
         7. Inserts the summary line with hyperlinks.
         8. Adds bullet bookmarks to each bullet paragraph.
        The separator and add_extra_bookmarks flag allow you to adjust the behavior
        (for example, basic vs. extended summary).
    Shortcut: Ctrl + Shift + Alt + N
    """
    manager = BulletPointManager()
    manager.process_nested_bookmark_summary(separator="| ", add_extra_bookmarks=True)

def change_character_style():
    """
    Function:
        - Reads the parent's bullet title portion (exact text run) to retrieve its
        - character style name, then assigns that style to the child's title text.
    Shortcut: Ctrl + Shift + Alt + S
    """
    manager = BulletPointManager()
    manager.propagate_title_character_style()
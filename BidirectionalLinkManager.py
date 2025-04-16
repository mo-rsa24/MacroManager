import uno
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
from com.sun.star.awt.MessageBoxType import ERRORBOX, INFOBOX

# --- Bidirectional Link Manager Class ---
class BidirectionalLinkManager:
    def __init__(self):
        self.doc = XSCRIPTCONTEXT.getDocument()
        self.text = self.doc.Text
        self.view_cursor = self.doc.CurrentController.getViewCursor()
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

    def get_input_with_default(self, prompt, title, default_value):
        """
        Prompts the user for input using a BASIC input box.
        Returns the trimmed answer or default_value if input is blank.
        """
        script_provider = self.smgr.createInstanceWithContext(
            "com.sun.star.script.provider.MasterScriptProviderFactory", self.ctx
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

    def get_selected_clean_title(self):
        """
        Retrieves the current selected text, verifies it contains a colon,
        and returns the text before the colon (the 'clean' title).
        """
        selected_text = self.view_cursor.getString()
        colon_index = selected_text.find(":")
        if colon_index == -1:
            self.show_message("Please include a colon (:) in the selected text.",
                              "Formatting Error", boxtype=ERRORBOX)
            raise Exception("Colon not found in selection")
        return selected_text[:colon_index].strip()

    def bookmark_exists(self, name):
        """
        Checks if a bookmark with the given name already exists.
        """
        return self.doc.getBookmarks().hasByName(name)

    def create_main_bookmark(self, clean_title, main_bookmark):
        """
        Creates the main bookmark covering the clean_title in the current paragraph.
        """
        line_cursor = self.text.createTextCursorByRange(self.view_cursor.getStart())
        line_cursor.gotoStartOfParagraph(False)
        if not line_cursor.goRight(len(clean_title), True):
            self.show_message("Error selecting the title text.", "Error", boxtype=ERRORBOX)
            raise Exception("Error selecting title text")
        main_bm_obj = self.doc.createInstance("com.sun.star.text.Bookmark")
        main_bm_obj.Name = main_bookmark
        self.text.insertTextContent(line_cursor, main_bm_obj, True)

    def apply_marker_hyperlink(self, clean_title, toc_bookmark, replacement_char):
        """
        Applies a hyperlink on the marker immediately following the clean_title.
        If replacement_char is not the default colon, it replaces the marker.
        """
        marker_cursor = self.text.createTextCursorByRange(self.view_cursor.getStart())
        if marker_cursor.goRight(len(clean_title), False) and marker_cursor.goRight(1, True):
            marker_text = marker_cursor.getString()
            if marker_text == ":":
                # Replace colon with alternate character if needed
                if replacement_char != ":":
                    marker_cursor.String = replacement_char
                full_doc_url = self.doc.URL
                # if not full_doc_url:
                #     self.show_message("Please save the document before running this macro.",
                #                       "Save Required", boxtype=ERRORBOX)
                #     raise Exception("Document not saved")
                marker_cursor.HyperLinkURL = full_doc_url + "#" + toc_bookmark
                marker_cursor.HyperLinkName = toc_bookmark
                marker_cursor.HyperLinkTarget = ""
        else:
            self.show_message("Error applying hyperlink to marker.", "Error", boxtype=ERRORBOX)
            raise Exception("Error in marker hyperlink")

    def insert_navigation_line(self, clean_title, main_bookmark, toc_bookmark):
        """
        Inserts a navigation line above the original paragraph that displays the clean title.
        Applies a hyperlink (linking back to the main bookmark) and adds a TOC bookmark.
        """
        insert_cursor = self.text.createTextCursorByRange(self.view_cursor.getStart())
        insert_cursor.gotoStartOfParagraph(False)
        navigation_line = clean_title + chr(13)
        self.text.insertString(insert_cursor, navigation_line, False)
        if not insert_cursor.goLeft(len(navigation_line), True):
            self.show_message("Error selecting the inserted navigation text.", "Error", boxtype=ERRORBOX)
            raise Exception("Error selecting navigation text")
        insert_cursor.collapseToStart()
        if not insert_cursor.goRight(len(clean_title), True):
            self.show_message("Error refining the selection for the navigation text.", "Error", boxtype=ERRORBOX)
            raise Exception("Error refining navigation text")
        full_doc_url = self.doc.URL
        if not full_doc_url:
            self.show_message("Please save the document before running this macro.",
                              "Save Required", boxtype=ERRORBOX)
            raise Exception("Document not saved")
        insert_cursor.HyperLinkURL = full_doc_url + "#" + main_bookmark

        toc_bm_obj = self.doc.createInstance("com.sun.star.text.Bookmark")
        toc_bm_obj.Name = toc_bookmark
        self.text.insertTextContent(insert_cursor, toc_bm_obj, True)

    def process_link(self, naming_strategy, replacement_char=":"):
        """
        Core processing routine to create bi-directional links. Steps:
         1. Validate selection and extract clean title.
         2. Use the provided naming_strategy to generate bookmark names.
         3. Check for existing bookmarks.
         4. Create the main bookmark, apply the marker hyperlink,
            and insert the navigation line with the TOC bookmark.
         5. Notify the user on success.
        The replacement_char parameter lets you customize the marker (e.g. "↑").
        """
        try:
            clean_title = self.get_selected_clean_title()
        except Exception:
            return

        try:
            main_bookmark, toc_bookmark, success_msg = naming_strategy(self, clean_title)
        except Exception:
            return

        if self.bookmark_exists(main_bookmark) or self.bookmark_exists(toc_bookmark):
            self.show_message("Bookmark name already exists. Please choose a different name.",
                              "Bookmark Exists", boxtype=ERRORBOX)
            return

        try:
            self.create_main_bookmark(clean_title, main_bookmark)
            self.apply_marker_hyperlink(clean_title, toc_bookmark, replacement_char)
            self.insert_navigation_line(clean_title, main_bookmark, toc_bookmark)
        except Exception:
            return

        self.show_message(success_msg, "Success", boxtype=INFOBOX)


# --- Naming Strategies (Strategy Pattern) ---
def naming_strategy_section(manager, clean_title):
    """Naming strategy for bidirectional_link:
       Asks for a section number and returns names like:
       "Section {section_number} {clean_title}" and its corresponding TOC bookmark.
    """
    section_number = manager.get_input_with_default("Enter Section Number (e.g., 1)",
                                                     "Section Number", "1")
    main_bookmark = f"Section {section_number} {clean_title}"
    toc_bookmark = main_bookmark + " Contents"
    return main_bookmark, toc_bookmark, f"✅ Bi-directional link created for: {clean_title}"

def naming_strategy_parent(manager, clean_title):
    """Naming strategy for bidirectional_link_with_parent:
       Prompts for a parent bookmark and prepends it to the clean title.
    """
    parent_bm = manager.get_input_with_default("Enter the name of the parent bookmark",
                                                "Parent Bookmark", "")
    if not parent_bm:
        manager.show_message("Parent bookmark name cannot be empty.",
                             "Input Error", boxtype=ERRORBOX)
        raise Exception("Parent bookmark empty")
    main_bookmark = f"{parent_bm} {clean_title}"
    toc_bookmark = main_bookmark + " Contents"
    return main_bookmark, toc_bookmark, f"✅ Bi-directional link created for: {main_bookmark}"

def naming_strategy_custom(manager, clean_title):
    """Naming strategy for custom_bidirectional_link and custom_bidirectional_link_for_code:
       Prompts for a custom bookmark name.
    """
    bm = manager.get_input_with_default("Enter the name of your bookmark",
                                          "Bookmark", clean_title)
    if not bm:
        manager.show_message("Bookmark name cannot be empty.",
                             "Input Error", boxtype=ERRORBOX)
        raise Exception("Bookmark name empty")
    main_bookmark = bm
    toc_bookmark = main_bookmark + " Contents"
    return main_bookmark, toc_bookmark, f"✅ Bi-directional link created for: {main_bookmark}"


# --- Module-level API Functions ---
def bidirectional_link():
    """
    Function:
        - Creates bi-directional bookmarks and hyperlinks for a selected heading using the section-number strategy.
    Shortcut: Ctrl + Shift + Alt + A
    """
    manager = BidirectionalLinkManager()
    manager.process_link(naming_strategy_section, replacement_char=":")

def bidirectional_link_with_parent():
    """
    Function:
        - Creates bi-directional bookmarks and hyperlinks for a selected heading using the parent-bookmark strategy.
    Shortcut: Ctrl + Shift + Alt + H
    """
    manager = BidirectionalLinkManager()
    manager.process_link(naming_strategy_parent, replacement_char=":")

def custom_bidirectional_link():
    """
    Function:
        - Creates bi-directional bookmarks and hyperlinks using a custom bookmark name (with colon marker).
    Shortcut: Ctrl + Shift + Alt + B
    """
    manager = BidirectionalLinkManager()
    manager.process_link(naming_strategy_custom, replacement_char=":")

def custom_bidirectional_link_for_code():
    """
    Function:
        - Creates bi-directional bookmarks and hyperlinks using a custom bookmark name but replaces the colon with an up arrow.
    Shortcut: Ctrl + Shift + Alt + C
    """
    manager = BidirectionalLinkManager()
    manager.process_link(naming_strategy_custom, replacement_char="↑")

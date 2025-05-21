import uno
import unohelper
import os
import glob
import shutil
from com.sun.star.awt import MessageBoxButtons as BUTTONS_OK
from com.sun.star.awt.MessageBoxType import ERRORBOX, INFOBOX


class FileManager:
    def __init__(self, media_dir=None):
        self.ctx = uno.getComponentContext()
        self.smgr = self.ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        self.doc = self.desktop.getCurrentComponent()
        self.view_cursor = self.doc.CurrentController.getViewCursor()
        self.text = self.doc.Text
        self.doc_url = uno.fileUrlToSystemPath(self.doc.URL)
        self.doc_dir = os.path.dirname(self.doc_url)
        self.media_dir = media_dir or os.path.expanduser("~/Pictures/Screenshots")

    # --- 📦 Latest Media File Retrieval ---
    def get_latest_media_file(self):
        """
        Finds the most recent .png, .mp4, or .webm file in the media directory.
        If no such files exist, displays an error and halts the process.
        """
        extensions = ['*.png', '*.svg','*.jpg','*.jpeg','*.mp4', '*.webm', '*.mov', '*.avi', '*.webp']
        files = []
        for ext in extensions:
            files.extend(glob.glob(os.path.join(self.media_dir, ext)))

        if not files:
            self.show_message(
                title="No Media Files Found",
                message=f"No image or video files found in:\n{self.media_dir}",
                boxtype=ERRORBOX
            )
            raise FileNotFoundError("No media files found.")

        return max(files, key=os.path.getmtime)

    def get_latest_document_file(self):
        """
        Looks in the vmshare directory for the most recent PDF or DOCX file.
        """
        doc_dir = os.path.expanduser("~/vmshare")
        extensions = ['*.pdf','*.pptx','*.docx', '*.xlsx']
        # Check if the directory exists     
        files = []
        for ext in extensions:
            files.extend(glob.glob(os.path.join(doc_dir, ext)))

        if not files:
            self.show_message(
                title="No Documents Found",
                message=f"No PDF or DOCX files found in:\n{doc_dir}",
                boxtype=ERRORBOX
            )
            raise FileNotFoundError("No document files found in vmshare.")

        return max(files, key=os.path.getmtime)

    # --- 🧠 Text Selection from Document ---
    def get_selected_text_and_range(self):
        selection = self.doc.getCurrentSelection()
        if not selection or selection.getCount() == 0:
            raise ValueError("No text selected.")
        text_range = selection.getByIndex(0)
        selected_text = text_range.getString().strip()
        if not selected_text:
            raise ValueError("Selected text is empty.")
        return selected_text, text_range

    # --- 📁 File Path Prep & Move ---
    def prepare_target_path(self, folder_name, filename):
        target_dir = os.path.join(self.doc_dir, folder_name)
        os.makedirs(target_dir, exist_ok=True)
        return os.path.join(target_dir, filename)

    def move_and_rename(self, src_path, dest_path):
        shutil.move(src_path, dest_path)
        return dest_path

    # --- 🔗 LibreOffice Hyperlink Injection ---
    def insert_hyperlink(self, text_range, file_path, label):
        text_range.setString(label)
        text_range.HyperLinkURL = uno.systemPathToFileUrl(file_path)
        text_range.HyperLinkName = label
        text_range.HyperLinkTarget = ""

    # --- 🧾 Error / Info Message ---
    def show_message(self, message, title="Message", boxtype=INFOBOX):
        frame = self.doc.CurrentController.Frame
        container_window = frame.ContainerWindow
        toolkit = self.smgr.createInstance("com.sun.star.awt.Toolkit")
        box = toolkit.createMessageBox(container_window, boxtype, 1, title, message)
        box.execute()

    # --- 🔧 Entry Method ---
    def attach_latest_media_to(self, folder_name):
        """
        Core method to:
        - Fetch latest media file (.png, .mp4, .webm)
        - Rename using selected text in the document
        - Move into a subfolder under the document path
        - Insert hyperlink over the selected text
        """
        try:
            media_path = self.get_latest_media_file()
            selected_text, text_range = self.get_selected_text_and_range()

            ext = os.path.splitext(media_path)[-1].lower()
            safe_name = selected_text + ext
            target_path = self.prepare_target_path(folder_name, safe_name)
            final_path = self.move_and_rename(media_path, target_path)

            self.insert_hyperlink(text_range, final_path, selected_text)
            self.show_message("✅ Media Linked", f"Stored in: {folder_name}/")

        except Exception as e:
            self.show_message("❌ Error", str(e), boxtype=ERRORBOX)

    def attach_latest_document_to_pdf_folder(self, folder_name):
        """
        Finds the latest PDF or DOCX from ~/vmshare,
        moves it to a 'PDF' subfolder in the doc directory,
        renames it based on selected text, and hyperlinks it in the document.
        """
        try:
            doc_path = self.get_latest_document_file()
            selected_text, text_range = self.get_selected_text_and_range()

            ext = os.path.splitext(doc_path)[-1].lower()
            safe_name = selected_text + ext
            target_path = self.prepare_target_path(folder_name, safe_name)
            final_path = self.move_and_rename(doc_path, target_path)

            self.insert_hyperlink(text_range, final_path, selected_text)
            self.show_message("✅ Document Linked", "Stored in: PDF/")

        except Exception as e:
            self.show_message("❌ Error", str(e), boxtype=ERRORBOX)

    def show_file_picker(self, title="Select a file", filter_names=None, filter_extensions=None, initial_folder=None):
        """
        Opens a native LibreOffice file picker.
        Returns the selected file path or None if canceled.
        """
        filepicker = self.smgr.createInstanceWithContext(
            "com.sun.star.ui.dialogs.FilePicker", self.ctx
        )
        filepicker.setTitle(title)

        if initial_folder:
            filepicker.setDisplayDirectory(uno.systemPathToFileUrl(initial_folder))

        if filter_names and filter_extensions:
            for name, ext in zip(filter_names, filter_extensions):
                filepicker.appendFilter(name, ext)
            filepicker.setCurrentFilter(filter_names[0])

        result = filepicker.execute()
        if result == 1:  # OK button
            return uno.fileUrlToSystemPath(filepicker.Files[0])
        else:
            return None

    def prompt_dropdown(self, prompt, title, options):
        """
        Presents a dropdown menu (combo box) for the user to choose from options.
        Returns the selected item, or None if cancelled.
        """
        try:
            script_provider = self.smgr.createInstanceWithContext(
                "com.sun.star.script.provider.MasterScriptProviderFactory", self.ctx
            ).createScriptProvider("")
            script = script_provider.getScript(
                "vnd.sun.star.script:Standard.Module1.DropdownDialog?language=Basic&location=application"
            )

            result = script.invoke((prompt, title, tuple(options)), (), ())
            return result[0] if result and isinstance(result, tuple) and result[0] else None
        except Exception as e:
            self.show_message("Dropdown Error", str(e), boxtype=ERRORBOX)
            return None

    def choose_file_and_target_folder(self, file_dir, extensions):
        """
        Allows user to manually choose a file from the file_dir, and choose a destination folder
        ('References' or 'Outputs'). Then inserts hyperlink in the doc using selected text.
        """
        # Gather eligible files
        files = []
        for ext in extensions:
            files.extend(glob.glob(os.path.join(file_dir, ext)))
        if not files:
            self.show_message("No files found.", f"No files in: {file_dir}", boxtype=ERRORBOX)
            return

        # Sort by modified time
        files = sorted(files, key=os.path.getmtime, reverse=True)

        # Step 1: Ask user to choose a file
        file_names = [os.path.basename(f) for f in files]
        selected_file = self.prompt_dropdown("Select a File", "File Selector", file_names)
        if not selected_file:
            return
        selected_path = os.path.join(file_dir, selected_file)

        # Step 2: Ask user to choose destination folder
        folder_choice = self.prompt_dropdown(
            "Select Destination Folder", "Where should the file go?",
            ["References", "Outputs"]
        )
        if not folder_choice:
            return

        # Step 3: Insert hyperlink
        try:
            selected_text, text_range = self.get_selected_text_and_range()
            ext = os.path.splitext(selected_path)[-1].lower()
            safe_name = selected_text.replace(" ", "_") + ext
            target_path = self.prepare_target_path(folder_choice, safe_name)
            final_path = self.move_and_rename(selected_path, target_path)
            self.insert_hyperlink(text_range, final_path, selected_text)
            self.show_message("✅ File Linked", f"Stored in: {folder_choice}/")
        except Exception as e:
            self.show_message("❌ Error", str(e), boxtype=ERRORBOX)


# --- 🧷 LibreOffice Macro-Compatible Entrypoints ---

def attach_media_macro():
    FileManager().attach_latest_media_to("")

def insert_media_into_references_folder():
    """
    Function:
        - Fetch latest media file (.png, .mp4, .webm)
        - Rename using selected text in the document
        - Move into a subfolder called "References"
        - Insert hyperlink over the selected text
    Shortcut: Ctrl + Shift + Alt + R
    """
    FileManager().attach_latest_media_to("References")

def insert_media_into_outputs_folder():
    """
        Function:
            - Fetch latest media file (.png, .mp4, .webm)
            - Rename using selected text in the document
            - Move into a subfolder called "Outputs"
            - Insert hyperlink over the selected text
        Shortcut: Ctrl + Shift + Alt + O
    """
    FileManager().attach_latest_media_to("Outputs")

def insert_latest_pdf_into_document():
    """
        Function:
            -  Finds the latest PDF or DOCX from ~/vmshare,
            -  moves it to a 'PDF' subfolder in the doc directory,
            -  renames it based on selected text, and hyperlinks it in the document.
        Shortcut: Ctrl + Shift + Alt + P
    """
    FileManager().attach_latest_document_to_pdf_folder("PDF")

def select_and_insert_file():
    manager = FileManager()
    manager.choose_file_and_target_folder(
        file_dir=os.path.expanduser(manager.media_dir),
        extensions=["*.pdf", "*.mp4", "*.webm", "*.png"]
    )

# youtube_clip_hyperlink_macro.py
import shutil
import uuid
from urllib.parse import unquote
import re
import uno, unohelper, threading, subprocess, sys, os, logging, pathlib, traceback
from logging.handlers import RotatingFileHandler
from com.sun.star.awt.MessageBoxType import ERRORBOX, INFOBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
from typing import Tuple

# ────────────────────────────────────────────────────
# 0. Logger Setup (unchanged)
# ────────────────────────────────────────────────────
LOG_DIR  = pathlib.Path.home() / ".cache" / "lo_youtube_clip_macro"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "youtube_clip_macro.log"

logger = logging.getLogger("YouTubeClipMacro")
logger.setLevel(logging.DEBUG)
logger.propagate = False

if not any(h.get_name() == "yt_clip_rot" for h in logger.handlers):
    handler = RotatingFileHandler(
        LOG_FILE, maxBytes=1_000_000, backupCount=3, encoding="utf-8"
    )
    handler.set_name("yt_clip_rot")
    handler.setFormatter(logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(threadName)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    ))
    logger.addHandler(handler)

logger.debug("Logger initialized")

def clear_log():
    try:
        with open(LOG_FILE, "w", encoding="utf-8") as f:
            f.write("")  # Overwrite with empty contents
        logger.debug("Log cleared before macro run")
    except Exception as e:
        logger.warning(f"Failed to clear log file: {e}")

# ────────────────────────────────────────────────────
# 1. Locate clipyt.py
# ────────────────────────────────────────────────────
try:
    BASE_DIR = os.path.dirname(__file__) or ""
except NameError:
    BASE_DIR = os.path.expanduser("~/.config/libreoffice/4/user/Scripts/python")
CLIP_SCRIPT_PATH = os.path.join(BASE_DIR, "clipyt.py")
TWTTER_SCRIPT_PATH = os.path.join(BASE_DIR, "twitter_dl.py")
TIK_TOK_SCRIPT_PATH = os.path.join(BASE_DIR, "pyktok_dl.py")
REDDIT_SCRIPT_PATH = os.path.join(BASE_DIR, "reddit_dl.py")
INSTAGRAM_SCRIPT_PATH = os.path.join(BASE_DIR, "instagram_dl.py")

if not os.path.isfile(INSTAGRAM_SCRIPT_PATH):
    raise FileNotFoundError(f"instagram_dl.py not found at {INSTAGRAM_SCRIPT_PATH}")

if not os.path.isfile(CLIP_SCRIPT_PATH):
    raise FileNotFoundError(f"clipyt.py not found at {CLIP_SCRIPT_PATH}")

if not os.path.isfile(TWTTER_SCRIPT_PATH):
    raise FileNotFoundError(f"twitter_dl.py not found at {TWTTER_SCRIPT_PATH}")

if not os.path.isfile(TIK_TOK_SCRIPT_PATH):
    raise FileNotFoundError(f"pyktok_dl.py not found at {TIK_TOK_SCRIPT_PATH}")

if not os.path.isfile(REDDIT_SCRIPT_PATH):
    raise FileNotFoundError(f"reddit_dl.py not found at {REDDIT_SCRIPT_PATH}")


# ────────────────────────────────────────────────────
# 2. UNO Dialog Helpers
# ────────────────────────────────────────────────────
def get_user_input(prompt: str, title: str, default: str = "") -> str:
    logger.debug(f"Prompting user: {title} (default='{default}')")
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    provider = smgr.createInstanceWithContext(
        "com.sun.star.script.provider.MasterScriptProviderFactory", ctx
    ).createScriptProvider("")
    script = provider.getScript(
        "vnd.sun.star.script:Standard.Module1.InputBoxWrapper"
        "?language=Basic&location=application"
    )
    args = (f"{prompt}\n(Leave blank to cancel)", title, default)
    result = script.invoke(args, (), ())
    val = (result[0].strip() if result and isinstance(result, tuple) else "")
    logger.debug(f"User input '{title}': '{val}'")
    return val

def show_message(doc, title: str, msg: str, boxtype=INFOBOX):
    logger.info(f"MessageBox → {title}: {msg}")
    frame = doc.CurrentController.Frame
    win   = frame.ContainerWindow
    toolkit = frame.getToolkit()
    box = toolkit.createMessageBox(win, boxtype, BUTTONS_OK, title, msg)
    box.execute()
# ────────────────────────────────────────────────────
# 3. Clip Manager: select → thread → hyperlink
# ────────────────────────────────────────────────────
class ClipManager:
    def __init__(self, folder_name):
        self.doc = XSCRIPTCONTEXT.getDocument()
        self.folder_name = folder_name

    def run(self):
        logger.info("Macro invoked: run_youtube_clip")
        # 1) Grab the selected range
        sel = self.doc.getCurrentSelection()
        if not sel or sel.getCount() == 0:
            show_message(self.doc, "Error", "Please select a word first.", boxtype=ERRORBOX)
            return
        rng = sel.getByIndex(0)
        label = rng.getString().strip()
        if not label:
            show_message(self.doc, "Error", "Selection is empty.", boxtype=ERRORBOX)
            return

        # 2) Prompt for clip parameters
        url   = get_user_input("YouTube URL:", "Video URL")
        if not url:
            logger.info("Cancelled at URL prompt"); return
        start = get_user_input("Start time (MM:SS or HH:MM:SS):", "Start Time", "00:00:00")
        if not start:
            logger.info("Cancelled at Start Time prompt"); return
        end   = get_user_input("End time   (MM:SS or HH:MM:SS):", "End Time",   "00:00:00")
        if not end:
            logger.info("Cancelled at End Time prompt"); return

        logger.info(f"Parameters: url={url}, start={start}, end={end}")
        # 3) Store for background and launch thread
        self.range = rng
        self.label = label
        self.params = (url, start, end)
        threading.Thread(target=self._worker, daemon=True, name=f"YT-{uuid.uuid4().hex[:6]}").start()
        logger.info("Background thread started")

    def _worker(self):
        thread_name = threading.current_thread().name
        url, start, end = self.params
        logger.info(f"[{thread_name}] Worker begin")
        try:
            # call clipyt.py
            logger.debug(f"[{thread_name}] Running {CLIP_SCRIPT_PATH}")
            proc = subprocess.run(
                [sys.executable, CLIP_SCRIPT_PATH, url, start, end],
                capture_output=True, text=True
            )
            logger.debug(f"[{thread_name}] Exit code: {proc.returncode}")
            logger.debug(f"[{thread_name}] STDOUT:\n{proc.stdout}")
            logger.debug(f"[{thread_name}] STDERR:\n{proc.stderr}")

            if proc.returncode != 0:
                err = proc.stderr.strip() or proc.stdout.strip() or f"Exit {proc.returncode}"
                logger.error(f"[{thread_name}] Download failed: {err} \n {str(err)}")
                return

            saved = None
            for line in proc.stdout.splitlines():
                if line.startswith("✓ Saved to "):
                    saved = line.split("✓ Saved to ",1)[1].strip()
                    break
            if not saved or not os.path.exists(saved):
                msg = "Download succeeded but file not found."
                logger.error(f"[{thread_name}] {msg}")
                return

            logger.info(f"[{thread_name}] Download OK: {saved}")

            # determine the current document directory
            logger.info(f"Attaining the current document directory")
            doc_url = self.doc.URL  # UNO file URL of the document
            sys_doc_path = uno.fileUrlToSystemPath(doc_url)
            doc_dir = os.path.dirname(unquote(sys_doc_path))
            logger.info(f"📁 The current document directory is {doc_dir}")

            # create target directory under document path
            target_dir = os.path.join(doc_dir, self.folder_name)
            logger.info(f"Preparing target folder: {target_dir}")
            try:
                if not os.path.isdir(target_dir):
                    os.makedirs(target_dir, exist_ok=True)
                    logger.info(f"Created target folder: {target_dir}")
                else:
                    logger.info(f"Target folder already exists: {target_dir}")
            except Exception as e:
                logger.error(f"Could not create or access target folder: {e}", exc_info=True)
                return  # bail out

            # move the downloaded file into the target directory
            basename = os.path.basename(saved)
            new_path = os.path.join(target_dir, basename)
            logger.info(f"Moving file:\n  from: {saved}\n    to: {new_path}")
            try:
                shutil.move(saved, new_path)
                logger.info(f"📁 Moved downloaded file from {saved} → {new_path}")
            except Exception as e:
                logger.error(f"Failed to move file to {new_path}: {e}", exc_info=True)
                return
            base, ext = os.path.splitext(new_path)
            # 5) Sanitize & rename
            safe_label = re.sub(r'[^\w\-]+', '_', self.label).strip('_')[:100]
            new_filename = f"{safe_label}{ext}"
            renamed_path = os.path.join(target_dir, new_filename)
            logger.info(f"✏️   Renaming to: {new_filename}")

            if os.path.exists(renamed_path):
                logger.warning(f"⚠️   Overwriting existing: {renamed_path}")
                try:
                    os.remove(renamed_path)
                    logger.info(f"🗑️   Removed old file")
                except Exception as e:
                    logger.error(f"❌  Could not remove old file: {e}", exc_info=True)

            try:
                os.rename(new_path, renamed_path)
                logger.info(f"✅  Renamed to: {renamed_path}")
                saved = renamed_path
            except Exception as e:
                logger.error(f"❌  Rename failed: {e}", exc_info=True)
                return

            # 6) Insert hyperlink
            logger.info(f"🔗  Inserting hyperlink for '{self.label}'")
            try:
                self.range.setString(self.label)
                self.range.HyperLinkURL = uno.systemPathToFileUrl(saved)
                self.range.HyperLinkName = self.label
                self.range.HyperLinkTarget = ""
                logger.info(f"[{thread_name}]✅   Direct hyperlink applied to '{self.label}' → {saved}")
            except Exception as exc:
                logger.exception(f"[{thread_name}] ❌  Direct hyperlink failed: {exc}")


        except Exception as exc:
            tb = traceback.format_exc()
            logger.exception(f"[{thread_name}] Exception:\n{tb} - {str(exc)}")
        finally:
            logger.info(f"[{thread_name}] Worker end")

    def _download_social_clip(self, script_path):
        """
        Runs the given downloader script (e.g. twitter_dl.py or pyktok_dl.py) in a thread,
        then moves the file to the document-relative folder and hyperlinks it.
        """
        self.social_script = script_path
        sel = self.doc.getCurrentSelection()
        if not sel or sel.getCount() == 0:
            show_message(self.doc, "Error", "Please select a word.", boxtype=ERRORBOX)
            return

        rng = sel.getByIndex(0)
        label = rng.getString().strip()
        if not label:
            show_message(self.doc, "Error", "Selected text is empty.", boxtype=ERRORBOX)
            return

        url = get_user_input("Enter video URL:", f"Social Media Download")
        if not url:
            return

        self.range = rng
        self.label = label
        self.params = (url,)
        threading.Thread(target=self._worker_social, daemon=True).start()

    def _worker_social(self):
        url = self.params[0]
        thread_name = threading.current_thread().name
        logger.info(f"[{thread_name}] Downloading via {self.social_script}: {url}")
        try:
            proc = subprocess.run(
                [sys.executable, self.social_script, url],
                capture_output=True, text=True
            )
            logger.debug(f"STDOUT:\n{proc.stdout}")
            logger.debug(f"STDERR:\n{proc.stderr}")

            if proc.returncode != 0:
                msg = proc.stderr.strip() or proc.stdout.strip() or f"Exit {proc.returncode}"
                logger.error(f"[{thread_name}] Download failed: {msg}")
                return

            saved = None
            for line in proc.stdout.splitlines():
                if line.startswith("✓ Saved to "):
                    saved = line.split("✓ Saved to ", 1)[1].strip()
                    break

            if not saved or not os.path.exists(saved):
                logger.error(f"[{thread_name}] File not found: {saved}")
                return

            logger.info(f"[{thread_name}] Downloaded: {saved}")

            # determine the current document directory
            logger.info(f"Attaining the current document directory")
            doc_url = self.doc.URL  # UNO file URL of the document
            sys_doc_path = uno.fileUrlToSystemPath(doc_url)
            doc_dir = os.path.dirname(unquote(sys_doc_path))
            logger.info(f"📁 The current document directory is {doc_dir}")

            # create target directory under document path
            target_dir = os.path.join(doc_dir, self.folder_name)
            logger.info(f"Preparing target folder: {target_dir}")
            try:
                if not os.path.isdir(target_dir):
                    os.makedirs(target_dir, exist_ok=True)
                    logger.info(f"Created target folder: {target_dir}")
                else:
                    logger.info(f"Target folder already exists: {target_dir}")
            except Exception as e:
                logger.error(f"Could not create or access target folder: {e}", exc_info=True)
                return  # bail out

            # move the downloaded file into the target directory
            basename = os.path.basename(saved)
            new_path = os.path.join(target_dir, basename)
            logger.info(f"Moving file:\n  from: {saved}\n    to: {new_path}")
            try:
                shutil.move(saved, new_path)
                logger.info(f"📁 Moved downloaded file from {saved} → {new_path}")
            except Exception as e:
                logger.error(f"Failed to move file to {new_path}: {e}", exc_info=True)
                return
            base, ext = os.path.splitext(new_path)
            # 5) Sanitize & rename
            safe_label = re.sub(r'[^\w\-]+', '_', self.label).strip('_')[:100]
            new_filename = f"{safe_label}{ext}"
            renamed_path = os.path.join(target_dir, new_filename)
            logger.info(f"✏️   Renaming to: {new_filename}")

            if os.path.exists(renamed_path):
                logger.warning(f"⚠️   Overwriting existing: {renamed_path}")
                try:
                    os.remove(renamed_path)
                    logger.info(f"🗑️   Removed old file")
                except Exception as e:
                    logger.error(f"❌  Could not remove old file: {e}", exc_info=True)

            try:
                os.rename(new_path, renamed_path)
                logger.info(f"✅  Renamed to: {renamed_path}")
                saved = renamed_path
            except Exception as e:
                logger.error(f"❌  Rename failed: {e}", exc_info=True)
                return

            # 6) Insert hyperlink
            logger.info(f"🔗  Inserting hyperlink for '{self.label}'")
            try:
                self.range.setString(self.label)
                self.range.HyperLinkURL = uno.systemPathToFileUrl(saved)
                self.range.HyperLinkName = self.label
                self.range.HyperLinkTarget = ""
                logger.info(f"[{thread_name}]✅   Direct hyperlink applied to '{self.label}' → {saved}")
            except Exception as exc:
                logger.exception(f"[{thread_name}] ❌  Direct hyperlink failed: {exc}")

        except Exception as e:
            logger.exception(f"[{thread_name}] Error during social media clip download: {e}")


# ────────────────────────────────────────────────────
# 4. Macro entry‑point
# ────────────────────────────────────────────────────
def insert_clip_into_references_folder():
    """
    Function:
        - Fetch YouTube URL within start_time & end_time
        - Rename using selected text in the document
        - Move into a subfolder called "References"
        - Insert hyperlink over the selected text
    Shortcut: Ctrl + Shift + Alt + V
    """
    clear_log()
    ClipManager("References").run()

def insert_clip_into_outputs_folder():
    """
    Function:
        - Fetch YouTube URL within start_time & end_time
        - Rename using selected text in the document
        - Move into a subfolder called "Outputs"
        - Insert hyperlink over the selected text
    Shortcut: Ctrl + Shift + Alt + D
    """
    clear_log()
    ClipManager("Outputs").run()


def download_twitter_clip_to_references_folder():
    """
    Prompts for Twitter URL and inserts clip into References folder
    """
    clear_log()
    ClipManager("References")._download_social_clip(TWTTER_SCRIPT_PATH)

def download_tik_tok_clip_to_references_folder():
    """
    Prompts for TikTok URL and inserts clip into References folder
    """
    clear_log()
    ClipManager("References")._download_social_clip(TIK_TOK_SCRIPT_PATH)

def download_reddit_clip_to_references_folder():
    """
    Downloads a Reddit clip and inserts it into the References folder.
    """
    clear_log()
    ClipManager("References")._download_social_clip(REDDIT_SCRIPT_PATH)

def download_instagram_clip_to_references_folder():
    """
    Downloads an Instagram clip and inserts it into the References folder.
    """
    clear_log()
    ClipManager("References")._download_social_clip(INSTAGRAM_SCRIPT_PATH)



g_exportedScripts = (
    insert_clip_into_references_folder,
    insert_clip_into_outputs_folder,
    download_twitter_clip_to_references_folder,
    download_tik_tok_clip_to_references_folder,
    download_reddit_clip_to_references_folder,
    download_instagram_clip_to_references_folder
)




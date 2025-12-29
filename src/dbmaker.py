import pyautogui as pag
import time
import os
import subprocess
import configparser
from pathlib import Path
from datetime import datetime
import pyperclip
import psutil  
from pathlib import Path
import pandas as pd
import sqlite3


# -------------------------------------------------------
# GLOBAL SETTINGS
# -------------------------------------------------------
pag.FAILSAFE = True
pag.PAUSE = 0.25

BASE_DIR = Path(__file__).resolve().parent
IMG_DIR = BASE_DIR / "Images"
FILE_DIR = Path(r"C:\Users\Curve System\Desktop\KsyncwAI\Excels")
RAW_FILE = FILE_DIR / "KowsarExport.xlsx"
CLEAN_FILE = FILE_DIR / "InStockClean.xlsx"
DB_FILE = FILE_DIR / "InStock.db"


# -------------------------------------------------------
# Normalize columns
# -------------------------------------------------------
def normalize_persian(text: str) -> str:
    if not isinstance(text, str):
        return text

    mapping = {
        "\u064a": "ی",  # ي -> ی
        "\u0649": "ی",  # ى -> ی
        "\u0626": "ی",  # ئ -> ی
        "\u0643": "ک",  # ك -> ک
        "\u0629": "ه",  # ة -> ه
        "\u200c": "",   # ZWNJ
        "\u0640": "",   # ـ (kashida)
    }

    for src, dst in mapping.items():
        text = text.replace(src, dst)
    return text


def build_clean_excel_db():
    # 1) Load raw export from Kowsar
    if not RAW_FILE.exists():
        print(f"RAW export not found: {RAW_FILE}")
        return

    df = pd.read_excel(RAW_FILE)

    # 2) Normalize the book name column
    col_name = "نام كتاب"
    if col_name in df.columns:
        df[col_name] = df[col_name].astype(str).map(normalize_persian)
    else:
        print(f"Column '{col_name}' not found in export!")
        print("Columns are:", list(df.columns))
        return

    # 3) (optional) normalize other Persian text columns too, e.g. publisher, author:
    # for c in ["نام نويسنده", "نام مترجم", "نام ناشر"]:
    #     if c in df.columns:
    #         df[c] = df[c].astype(str).map(normalize_persian)

    # 4) Save clean DB for the AI to use
    df.to_excel(CLEAN_FILE, index=False)
    print(f"Clean DB updated: {CLEAN_FILE}")

def excel_to_sqlite(excel_path: Path, db_path: Path):
    """
    Convert the cleaned Excel file into a SQLite database.
    Table name: products
    """
    if not excel_path.exists():
        print(f"Clean Excel not found, skipping SQLite update: {excel_path}")
        return

    print(f"Loading clean Excel: {excel_path}")
    df = pd.read_excel(excel_path)

    # You can drop unused columns here if you want:
    # df = df[["نام كتاب", "كد كتاب", "تعداد", "قيمت", ...]]

    print(f"Writing {len(df)} rows to SQLite DB: {db_path}")
    conn = sqlite3.connect(db_path)
    df.to_sql("products", conn, if_exists="replace", index=False)
    conn.close()
    print(f"SQLite DB updated: {db_path}")


# -------------------------------------------------------
# CONFIG
# -------------------------------------------------------
config = configparser.ConfigParser()
config.read(BASE_DIR / "config.ini", encoding="utf-8")

INTERVAL_MINUTES = int(
    config.get("general", "interval_minutes", fallback="15").strip()
)

KOWSAR_EXE  = config.get("kowsar", "exe_path").strip()
USERNAME    = config.get("kowsar", "username").strip()
PASSWORD    = config.get("kowsar", "password").strip()
FILTER_FILE = config.get("kowsar", "filter_file").strip()

EXPORT_DIR = Path(
    config.get(
        "kowsar",
        "export_dir",
        fallback=r"C:\Users\Curve System\Desktop\KsyncwAI\Excels",
    ).strip()
).resolve()

EXPORT_DIR.mkdir(parents=True, exist_ok=True)

# -------------------------------------------------------
# HELPERS
# -------------------------------------------------------
def log(msg: str):
    t = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{t}] {msg}")


def is_kowsar_running() -> bool:
    """Check if ACC.exe is already running."""
    for proc in psutil.process_iter(["name"]):
        try:
            if proc.info["name"] and "ACC.exe" in proc.info["name"]:
                return True
        except Exception:
            pass
    return False


def wait_image(name, timeout=20, confidence=0.88):
    path = str(IMG_DIR / name)
    log(f"Waiting for {name}...")
    start = time.time()
    while time.time() - start < timeout:
        loc = pag.locateCenterOnScreen(path, confidence=confidence)
        if loc:
            log(f"Found {name}")
            return loc
        time.sleep(0.4)
    raise RuntimeError(f"Image not found: {name}")


def click_image(name, clicks=1, button="left", timeout=20, confidence=0.88):
    loc = wait_image(name, timeout, confidence)
    pag.click(loc.x, loc.y, clicks=clicks, button=button)
    time.sleep(0.3)


def safe_paste(text: str):
    pyperclip.copy(text)
    time.sleep(0.25)
    pag.hotkey("ctrl", "a")
    time.sleep(0.2)
    pag.hotkey("ctrl", "v")
    time.sleep(0.4)


def is_report_window_open() -> bool:
    """Detect if we are already on the 'موجودی تمام کالاها' report window."""
    try:
        path = str(IMG_DIR / "refresh_button.jpg")
        loc = pag.locateCenterOnScreen(path, confidence=0.85)
        return loc is not None
    except Exception:
        return False
    
def clear_export_popup_if_exists():
    """Close the export 'اطلاع / تایید' popup from a previous cycle if it is visible."""
    try:
        click_image("save_ok_button.jpg", timeout=2, confidence=0.85)
        log("Found leftover export popup → clicked تایید.")
        time.sleep(0.5)
    except:
        log("No leftover export popup found.")



# -------------------------------------------------------
# LOGIN / STARTUP
# -------------------------------------------------------
def start_kowsar():
    log("Checking if Kowsar (ACC.exe) is already running...")

    if is_kowsar_running():
        log("✔ ACC.exe is already running. Using existing instance.")
        return

    log("ACC.exe not running. Starting a new instance...")
    exe_dir = os.path.dirname(KOWSAR_EXE)
    subprocess.Popen(f'"{KOWSAR_EXE}"', cwd=exe_dir, shell=True)
    time.sleep(5)


def login_if_needed():
    """If login window is visible, perform login. Otherwise do nothing."""
    try:
        path = str(IMG_DIR / "login_button.jpg")
        loc = pag.locateCenterOnScreen(path, confidence=0.85)
    except Exception:
        loc = None

    if not loc:
        log("Login window not visible – assuming already logged in.")
        return

    log("Logging into Kowsar...")

    # Username
    click_image("username_field.jpg")
    safe_paste(USERNAME)

    # Password
    click_image("password_field.jpg")
    safe_paste(PASSWORD)

    # Login
    click_image("login_button.jpg")
    time.sleep(5)


# -------------------------------------------------------
# NAVIGATE TO STOCK REPORT + APPLY FILTER (first time only)
# -------------------------------------------------------
def open_stock_report():
    log("Navigating to stock report window...")

    click_image("reports_menu.jpg", clicks=1, timeout=20, confidence=0.85)
    time.sleep(1)

    click_image("system_reports.jpg", clicks=1, timeout=20, confidence=0.85)
    time.sleep(1.5)

    click_image("stock_reports.jpg", clicks=2, timeout=20, confidence=0.85)  # گزارشات انبار
    time.sleep(1.2)

    click_image("all_products.jpg", clicks=2, timeout=20, confidence=0.85)   # موجودی تمام کالاها
    time.sleep(4)


def apply_saved_filter_once():
    log("Applying saved filter (.flt)...")

    if not os.path.exists(FILTER_FILE):
        raise RuntimeError(f"Filter file missing: {FILTER_FILE}")

    # Open filter window (bottom-right red icon)
    click_image("filter_button.jpg")
    time.sleep(1)

    # Load -> open file dialog
    click_image("filter_load.jpg")
    time.sleep(1.2)

    # Type full filter path in dialog
    pag.hotkey("ctrl", "l")
    time.sleep(0.3)
    safe_paste(FILTER_FILE)
    pag.press("enter")
    time.sleep(1.5)

    # If you use a saved_filter2.jpg to click the row, keep this:
    try:
        click_image("saved_filter2.jpg", clicks=1, timeout=10, confidence=0.85)
        time.sleep(0.8)
        pag.press("enter")
        time.sleep(1.0)
    except Exception:
        log("saved_filter2.jpg not found – assuming file is already selected.")

    # Confirm filter
    click_image("filter_confirm.jpg")
    time.sleep(1.5)

    log("Filter applied.")

    # Initial refresh after applying filter
    click_image("refresh_button.jpg", clicks=1, timeout=15, confidence=0.85)
    log("Waiting for initial filtered results...")
    time.sleep(5)


def ensure_report_ready():
    """
    Make sure we are on the stock report window and filter is applied.
    If report window is already open -> do nothing.
    If not -> start/login/navigate/filter.
    """
    if is_report_window_open():
        log("✔ Report window already open. No need to navigate/filter.")
        return

    log("Report window not open – performing full setup.")
    start_kowsar()
    login_if_needed()
    open_stock_report()
    apply_saved_filter_once()


# -------------------------------------------------------
# EXPORT EXCEL (with confirmation)
# -------------------------------------------------------
def export_excel():
    log("Exporting Excel...")

    screen_w, screen_h = pag.size()

    # Focus grid
    pag.click(screen_w // 2, screen_h // 2)
    time.sleep(0.3)

    # Right-click -> context menu
    pag.click(screen_w // 2, screen_h // 2, button="right")
    time.sleep(0.4)

    # Click "انتقال به Excel"
    click_image("export_excel.jpg", timeout=10, confidence=0.85)
    time.sleep(2.5)

    # FIXED FILE NAME (always overwrite this one)
    pag.hotkey("ctrl", "l")
    time.sleep(0.3)
    safe_paste(EXPORT_DIR)
    pag.press("enter")
    time.sleep(1.5)
    full_path = str(EXPORT_DIR)
    # log(f"Saving as {full_path}")

    click_image("Master_excel.jpg", timeout=10, confidence=0.85)
    time.sleep(2.5)
    pag.press("enter")
    time.sleep(1.0)

    # Save dialog: paste full path
    # safe_paste(full_path)
    # pag.press("enter")
    # log("Waiting for export to complete...")

    # Wait for "اطلاع / تایید" popup and click تایید
    try:
        wait_image("save_ok_button.jpg", timeout=120, confidence=0.85)
        click_image("save_ok_button.jpg", timeout=5, confidence=0.85)
        log("Export confirmed by popup.")
    except Exception as e:
        log(f"Warning: export confirmation popup not found: {e}")

    log("Excel export finished.")
    return full_path



# -------------------------------------------------------
# ONE CYCLE = REFRESH + EXPORT
# -------------------------------------------------------
def run_stock_cycle():
    log("=== STOCK CYCLE START ===")

    try:
        clear_export_popup_if_exists()  # NEW STEP

        ensure_report_ready()
        click_image("refresh_button.jpg")
        time.sleep(5)
        export_excel()

    except Exception as e:
        log(f"❌ ERROR during cycle: {e}")

    log("=== STOCK CYCLE END ===")



# -------------------------------------------------------
# MAIN LOOP (EVERY INTERVAL_MINUTES)
# -------------------------------------------------------
def main_loop():
    log(
        f"Starting Kowsar In-Stock updater. "
        f"Interval = {INTERVAL_MINUTES} minutes."
    )
    while True:
        # 1) Do refresh + export from ACC.exe
        run_stock_cycle()

        # 2) Give Kowsar a bit of extra time, just in case
        time.sleep(15)

        # 3) Make sure export popup is closed (safety, in case timing shifts)
        try:
            wait_image("save_ok_button.jpg", timeout=10, confidence=0.85)
            click_image("save_ok_button.jpg", timeout=5, confidence=0.85)
            log("Confirmed export popup in main loop.")
        except Exception:
            log("No export popup found in main loop (probably already handled).")

        # 4) Build the cleaned Excel DB (fix Arabic/Persian)
        build_clean_excel_db()

        # 5) Convert cleaned Excel to SQLite DB
        excel_to_sqlite(CLEAN_FILE, DB_FILE)

        # 6) Sleep until next cycle
        log(f"Sleeping for {INTERVAL_MINUTES} minutes...")
        time.sleep(INTERVAL_MINUTES * 60)






if __name__ == "__main__":
    main_loop()

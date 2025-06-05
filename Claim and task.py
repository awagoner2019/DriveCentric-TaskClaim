from __future__ import annotations
#!/usr/bin/env python

"""
DriveCentric-TaskClaim

Automates DriveCentric "claim customer", task-date adjustment and standard /
custom text & e-mail outreach - with logging, hot-keys and auto-update support.
"""

import os, sys, time, json, shutil, zipfile, io, logging, datetime, traceback
import threading, socket, subprocess, stat
from pathlib import Path

try:
    import requests, keyboard
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
except Exception as _imp_err:
    import tkinter as _tk
    from tkinter import messagebox as _mb
    _tk.Tk().withdraw()
    _mb.showerror("Missing Python package",
        f"Required module missing:\n{_imp_err}\n\n"
        "Install the package (e.g. pip install selenium requests keyboard) "
        "and try again.")
    sys.exit(1)

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, scrolledtext

APP_NAME = "DriveCentricTaskClaim"

def resource_path(rel: str | Path) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / rel


def get_user_data_dir() -> Path:
    if os.name == "nt":
        root = Path(os.environ.get("LOCALAPPDATA", Path.home()))
    else:
        root = Path.home()
    return root / APP_NAME

USER_DATA_DIR = get_user_data_dir()
USER_DATA_DIR.mkdir(parents=True, exist_ok=True)

DEFAULT_VERSION = "1.0"
REPO_OWNER = "awagoner2019"
REPO_NAME = "DriveCentric-TaskClaim"

ALLOWED_USERS = {"Aaron Wagoner", "Nate Floyd", "Jean Luc", "Paige Craig", "Daulton Gentry"}
ALLOWED_PASSWORDS = {"Acura2025", "ADMIN"}

PIN_FILE = USER_DATA_DIR / "user_pins.json"
TEMPLATE_FILE = USER_DATA_DIR / "templates.json"
LOG_FILENAME = USER_DATA_DIR / "drivecentric_log.txt"

WATERMARK_ICON = "💠"
WATERMARK_TEXT = "© 2024 • Developed by Aaron Wagoner"

sender_name = ""
auto_stop_event = threading.Event()

# Default templates dictionary with placeholders for personalization
# It includes text and email templates for standard and custom outreach

templates: dict[str, str] = {
    "standard_text":
        "Hello, {customer_name}! This is {sender_name} with Acura of Springfield. "
        "Just wanted to check in - we've got strong offers and fresh inventory "
        "rolling in, and I didn't want you to miss out if you've been thinking "
        "about upgrading. Let me know if you're open to a quick chat!",

    "custom_text_A":
        "Hello, {customer_name}! We've got new offers and inventory updates at "
        "Acura of Springfield. Let me know if you'd like more details or a "
        "personalized overview.",

    "custom_text_B":
        "Hi, {customer_name}! I'm following up on your recent inquiry. "
        "How can I assist you further? Are you interested in a test drive or "
        "more details?",

    "custom_text_C":
        "Urgent: Limited-time offers are available now, {customer_name}. "
        "Please contact us immediately for details.",

    "email_subject":
        "Your Inquiry from Acura of Springfield",

    "email_body":
        "Hello, {customer_name},\n\n"
        "Thank you for reaching out to us at Acura of Springfield. We offer a "
        "diverse range of vehicles along with competitive financing and "
        "exclusive promotions designed to meet your needs.\n\n"
        "Please review our latest inventory and feel free to contact me if you "
        "have any questions or wish to schedule a visit.\n\n"
        "Best regards,\n{sender_name}",

    "custom_email_subject_A":
        "Big News at Acura of Springfield - Let's Talk!",

    "custom_email_body_A":
        "Hello, {customer_name},\n\n"
        "Great news: fresh inventory has just arrived, and I'd love to give you "
        "a first look before it's advertised publicly.\n\n"
        "If you'd like a personalised walk-around video or to schedule a test "
        "drive, just reply to this email.\n\n"
        "Kind regards,\n{sender_name}",

    "custom_email_subject_B":
        "Your Perfect Acura Could Be Waiting - Quick Update",

    "custom_email_body_B":
        "Hi, {customer_name},\n\n"
        "I wanted to follow up on your interest in our vehicles. Based on what "
        "you shared, I have a few options in mind that match your needs.\n\n"
        "Would you have a moment for a short call today?\n\n"
        "Thanks!\n{sender_name}",

    "custom_email_subject_C":
        "Limited-Time Savings Event - Exclusive Offers Inside",

    "custom_email_body_C":
        "Hello, {customer_name},\n\n"
        "We're running a limited-time savings event, and I didn't want you to "
        "miss out on the additional incentives we can provide right now.\n\n"
        "Please let me know if you'd like full details or personalised payment "
        "options.\n\n"
        "Best,\n{sender_name}"
}

logging.basicConfig(
    filename=str(LOG_FILENAME),
    level=logging.INFO,
    filemode="a",
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger("drivecentric")
logger.info("Program started.")

root: tk.Tk | None = None
log_text: scrolledtext.ScrolledText | None = None
status_var: tk.StringVar | None = None

def fatal_popup(msg: str):
    tk.Tk().withdraw()
    messagebox.showerror("Fatal Error", msg)
    sys.exit(1)

def gui_print(message: str, status: str | None = None):
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {message}\n"
    logger.info(message)
    if status and status_var is not None:
        status_var.set(status)
    if root and log_text and log_text.winfo_exists():
        def _append():
            log_text.configure(state="normal")
            log_text.insert(tk.END, line)
            log_text.see(tk.END)
            log_text.configure(state="disabled")
        try:
            root.after(0, _append)
        except Exception:
            pass

def load_templates():
    global templates
    if not TEMPLATE_FILE.is_file():
        save_templates()
        return
    try:
        with TEMPLATE_FILE.open("r", encoding="utf-8") as fp:
            disk = json.load(fp)
        if isinstance(disk, dict):
            templates.update(disk)
            logger.info("Templates loaded & merged.")
        else:
            logger.warning("Template file invalid format; not a dict, using defaults.")
    except Exception as exc:
        gui_print(f"Template load error: {exc}", status="Template load error")

def save_templates():
    try:
        with TEMPLATE_FILE.open("w", encoding="utf-8") as fp:
            json.dump(templates, fp, indent=4, ensure_ascii=False)
        gui_print("Templates saved.", status="Templates saved")
    except Exception as exc:
        gui_print(f"Could not save templates: {exc}", status="Template save error")

def load_pins():
    try:
        with PIN_FILE.open("r", encoding="utf-8") as fp:
            return json.load(fp)
    except Exception:
        return {}

def save_pins(pins: dict):
    try:
        with PIN_FILE.open("w", encoding="utf-8") as fp:
            json.dump(pins, fp, indent=2)
    except Exception:
        pass

def is_port_in_use(port: int, host: str = "127.0.0.1") -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.settimeout(1)
            s.connect((host, port))
            return True
        except Exception:
            return False

def forcibly_remove_folder(path: Path | str):
    path = Path(path)
    if not path.exists():
        return
    for rd, _, files in os.walk(path):
        for f in files:
            try:
                os.chmod(Path(rd) / f, stat.S_IWRITE)
            except Exception:
                pass
    shutil.rmtree(path, ignore_errors=True)

def get_windows_date() -> str:
    return datetime.datetime.now().strftime("%m/%d/%y")

def get_chrome_path() -> str:
    default_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ]
    for p in default_paths:
        if os.path.isfile(p):
            return p
    tk.Tk().withdraw()
    custom = simpledialog.askstring("Chrome not found",
                                   "Cannot locate chrome.exe.\nPlease enter the full path:")
    if custom and os.path.isfile(custom):
        return custom
    fatal_popup("Google Chrome is required. Exiting.")

def launch_chrome():
    chrome_path = get_chrome_path()
    user_data_dir = r"C:\TempChromeProfile"
    cmd = [
        chrome_path,
        "--remote-debugging-port=9222",
        f"--user-data-dir={user_data_dir}"
    ]
    try:
        subprocess.Popen(cmd)
        gui_print(
            "Chrome launched. Log into DriveCentric, then open a customer in a new tab/window if needed.",
            status="Chrome launched")
        time.sleep(1)
    except Exception as exc:
        gui_print(f"Could not launch Chrome: {exc}", status="Chrome error")

def _find_and_switch_to_drivecentric_tab(driver):
    keywords = ["drivecentric", "dealer", "crm"]
    for handle in driver.window_handles:
        try:
            driver.switch_to.window(handle)
            url = driver.current_url
            for kw in keywords:
                if kw in url.lower():
                    return True
        except Exception:
            continue
    return False

def get_chrome_driver():
    if not is_port_in_use(9222):
        gui_print("Remote debugging port 9222 not open. Click 'Launch Chrome' first.",
                  status="Chrome not attached")
        return None
    opts = Options()
    opts.debugger_address = "127.0.0.1:9222"
    try:
        driver = webdriver.Chrome(options=opts)
    except Exception as exc:
        gui_print(f"Cannot attach to Chrome: {exc}", status="Chrome attach error")
        return None
    found = _find_and_switch_to_drivecentric_tab(driver)
    if not found:
        gui_print(
            "DriveCentric tab not found. "
            "Please make sure you have DriveCentric open in one of the tabs/windows in Chrome "
            "(with --remote-debugging-port=9222 enabled). "
            "Then re-try your action after opening/selecting the correct customer tab.",
            status="Open customer tab in Chrome"
        )
        driver.quit()
        return None
    return driver

def safe_click(driver, elem):
    try:
        elem.click()
    except Exception:
        driver.execute_script("arguments[0].click();", elem)

def gui_login() -> str:
    pins = load_pins()
    while True:
        pin = simpledialog.askstring("Login", "Quick-PIN (blank for username):", show="*")
        if pin:
            matches = [u for u, p in pins.items() if p == pin]
            if len(matches) == 1:
                gui_print(f"User '{matches[0]}' logged in via PIN.", status="Logged in")
                return matches[0]
            messagebox.showerror("Login", "Incorrect PIN.")
            continue
        user = simpledialog.askstring("Login", "Username:")
        if not user or user not in ALLOWED_USERS:
            messagebox.showerror("Login", "Not authorised.")
            continue
        pwd = simpledialog.askstring("Login", f"Password for {user}:", show="*")
        if pwd not in ALLOWED_PASSWORDS:
            messagebox.showerror("Login", "Incorrect password.")
            continue
        if user not in pins and messagebox.askyesno("Quick-PIN", "Create quick-PIN for next time?"):
            p1 = simpledialog.askstring("Quick-PIN", "New PIN:", show="*")
            p2 = simpledialog.askstring("Quick-PIN", "Confirm PIN:", show="*")
            if p1 and p1 == p2:
                pins[user] = p1
                save_pins(pins)
        gui_print(f"User '{user}' logged in.", status="Logged in")
        return user

def is_customer_claimed(driver) -> bool:
    try:
        if not driver:
            return False
        new_deal_btns = driver.find_elements(By.XPATH, "//div[contains(@class,'act-button')]//div[contains(@class,'actionvalue') and contains(text(),'New Deal')]")
        if new_deal_btns:
            return True
        vehicles = driver.find_elements(By.XPATH, "//drc-add-vehicle")
        trades = driver.find_elements(By.XPATH, "//drc-add-trade")
        if vehicles or trades:
            return True
        claimed = driver.find_elements(By.XPATH, "//div[@analyticsdetect='Sidebar|Open|NewDeal']")
        if claimed:
            return True
        claim_btns = driver.find_elements(By.XPATH, "//*[contains(@analyticsdetect,'ClaimCustomer')]")
        if claim_btns:
            return False
        return False
    except Exception:
        return False

def customer_has_email(driver) -> bool:
    try:
        no_email_warning = driver.find_elements(
            By.XPATH, "//div[contains(@class,'cust-act-cnt-eml')]//div[contains(@class,'msg')]//h4[contains(text(),'no valid email specified')]"
        )
        return not bool(no_email_warning)
    except Exception:
        return True

def click_claim_and_replace(driver):
    try:
        claim_btn = None
        claim_btns = driver.find_elements(
            By.XPATH,
            "//*[contains(@analyticsdetect,'ClaimCustomer') and (self::button or self::div or self::a)]"
        )
        claim_btns = [b for b in claim_btns if b.is_displayed()] or claim_btns
        if claim_btns:
            claim_btn = claim_btns[0]
        else:
            btns = driver.find_elements(By.XPATH, "//button | //a | //div")
            for btn in btns:
                if btn.text.strip().lower().find("claim customer") >= 0:
                    claim_btn = btn
                    break
        if not claim_btn:
            html = driver.page_source
            logger.error("Claim button not found! DriveCentric DOM dumped for debugging.")
            gui_print("❌ Claim button not found. (See log for DOM html).")
            with open(USER_DATA_DIR / "last_dom.html", "w", encoding="utf-8") as f:
                f.write(html)
            return False
        driver.execute_script("arguments[0].scrollIntoView(true);", claim_btn)
        time.sleep(0.2)
        safe_click(driver, claim_btn)
        gui_print("✅ 'Claim Customer' button clicked. Waiting for modal...")
        try:
            radio_inputs = WebDriverWait(driver, 7).until(
                EC.presence_of_all_elements_located((By.XPATH, "//input[@type='radio']"))
            )
        except Exception:
            radio_inputs = []
        found = False
        for radio in radio_inputs:
            try:
                label_el = radio.find_element(By.XPATH, ".//following-sibling::label | ./parent::label")
                label_text = label_el.text.lower()
                if "remove" in label_text or "replace" in label_text or "you" in label_text:
                    driver.execute_script("arguments[0].scrollIntoView(true);", radio)
                    safe_click(driver, radio)
                    found = True
                    gui_print("Selected best salesperson radio option.")
                    break
            except Exception:
                continue
        if not found and radio_inputs:
            safe_click(driver, radio_inputs[0])
            gui_print("Default salesperson selected.")
        claim_btn_modal = None
        try:
            claim_btn_modal = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[.//span[normalize-space(text())='Claim']]")
                )
            )
        except Exception:
            btns = driver.find_elements(By.XPATH, "//button | //a | //div")
            for b in btns:
                if b.text.strip().lower() == "claim":
                    claim_btn_modal = b
                    break
        if claim_btn_modal is not None:
            safe_click(driver, claim_btn_modal)
            gui_print("🎯 Final 'Claim' confirmed.")
        else:
            gui_print("❌ Could not find final 'Claim' confirmation button.")
            return False
        time.sleep(1)
        return True
    except Exception as exc:
        logger.error(f"Error in claim process: {exc}\n{traceback.format_exc()}")
        gui_print(f"Error in claim process: {exc}")
        return False

def claim_only_customer():
    gui_print("\n--- Claim Only ---", status="Claim Only")
    def _work():
        drv = get_chrome_driver()
        if drv:
            try:
                if not is_customer_claimed(drv):
                    claimed = click_claim_and_replace(drv)
                    if claimed:
                        gui_print("Claimed customer.")
                    else:
                        gui_print("Claim not performed.")
                else:
                    gui_print("Customer already claimed.")
            except Exception as exc:
                gui_print(f"Claim only error: {exc}")
                logger.debug(traceback.format_exc())
        status_var.set("Ready")
    threading.Thread(target=_work, daemon=True).start()

def clear_input_fast(elem):
    elem.click()
    elem.send_keys(Keys.CONTROL, "a")
    elem.send_keys(Keys.BACKSPACE)
    elem.send_keys(Keys.CONTROL, "a")
    elem.send_keys(Keys.BACKSPACE)
    time.sleep(0.1)

def edit_task_after_claim(driver):
    try:
        WebDriverWait(driver, 12).until(
            EC.presence_of_element_located((By.XPATH, "//li[@analyticsdetect='Timeline|PerformAction|TaskToDo']"))
        )
        all_edit = driver.find_elements(By.XPATH,
            "//li[@analyticsdetect='Timeline|PerformAction|TaskToDo' and contains(.,'Edit')]")
        edit_btn = None
        for btn in all_edit:
            if btn.is_displayed():
                edit_btn = btn
                break
        if edit_btn is None:
            raise Exception("Edit button not found for task.")
        safe_click(driver, edit_btn)
        gui_print("Task 'Edit' opened.")
        time.sleep(0.5)
        act_btns = driver.find_elements(By.XPATH,
            "//div[contains(@class,'action-list__button') and (.//span[contains(text(),'Phone')] or .//span[contains(text(),'Text')])]")
        if act_btns:
            safe_click(driver, act_btns[0])
        tp_btns = driver.find_elements(By.XPATH,
            "//div[contains(@class,'drc-action-list-item') and .//span[contains(text(),'Touchpoint')]]")
        if tp_btns:
            safe_click(driver, tp_btns[0])
        date_input = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Select a date']"))
        )
        today = get_windows_date()
        clear_input_fast(date_input)
        date_input.send_keys(today)
        gui_print(f"Task date set to {today}.")
        save_btns = driver.find_elements(By.CSS_SELECTOR,
            "button.drc-button.kind-filled.type-primary.size-medium.state-default")
        for btn in save_btns:
            if btn.is_displayed() and btn.is_enabled():
                safe_click(driver, btn)
                break
        gui_print("Task saved.")
        time.sleep(0.5)
    except Exception as exc:
        gui_print(f"Task edit error: {exc}")
        logger.debug(traceback.format_exc())

def set_task_to_touchpoint(driver):
    try:
        WebDriverWait(driver, 12).until(
            EC.presence_of_element_located((By.XPATH, "//li[@analyticsdetect='Timeline|PerformAction|TaskToDo']"))
        )
        all_edit = driver.find_elements(By.XPATH,
            "//li[@analyticsdetect='Timeline|PerformAction|TaskToDo' and contains(.,'Edit')]")
        edit_btn = None
        for btn in all_edit:
            if btn.is_displayed():
                edit_btn = btn
                break
        if edit_btn is None:
            raise Exception("Edit button not found for task.")
        safe_click(driver, edit_btn)
        gui_print("Task 'Edit' opened.")
        time.sleep(0.5)
        act_btns = driver.find_elements(By.XPATH,
            "//div[contains(@class,'action-list__button') and (.//span[contains(text(),'Phone')] or .//span[contains(text(),'Text')])]")
        if act_btns:
            safe_click(driver, act_btns[0])
        tp_btns = driver.find_elements(By.XPATH,
            "//div[contains(@class,'drc-action-list-item') and .//span[contains(text(),'Touchpoint')]]")
        if tp_btns:
            safe_click(driver, tp_btns[0])
        else:
            gui_print("Touchpoint option NOT found in task edit window.")
        date_input = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Select a date']"))
        )
        today = get_windows_date()
        clear_input_fast(date_input)
        date_input.send_keys(today)
        gui_print(f"Task date set to {today}.")
        save_btns = driver.find_elements(By.CSS_SELECTOR,
            "button.drc-button.kind-filled.type-primary.size-medium.state-default")
        for btn in save_btns:
            if btn.is_displayed() and btn.is_enabled():
                safe_click(driver, btn)
                break
        gui_print("Task saved (set as Touchpoint).")
        time.sleep(0.5)
        return True
    except Exception as exc:
        gui_print(f"Touchpoint task error: {exc}")
        logger.debug(traceback.format_exc())
        return False

def ensure_claimed_and_update_task(driver):
    if is_customer_claimed(driver):
        gui_print("Customer already claimed, updating task ...")
        edit_task_after_claim(driver)
    else:
        gui_print("Customer not claimed. Claiming and updating task ...")
        if click_claim_and_replace(driver):
            edit_task_after_claim(driver)

def ensure_claimed_only(driver):
    if not is_customer_claimed(driver):
        gui_print("Customer not claimed; claiming ...")
        click_claim_and_replace(driver)

def send_email_message(driver):
    try:
        nav_email_tabs = driver.find_elements(By.XPATH, "//li[contains(@analyticsdetect,'CustomerAction|Navigate|Email') and not(contains(@class,'active'))]")
        if nav_email_tabs:
            safe_click(driver, nav_email_tabs[0])
            time.sleep(0.5)
        else:
            email_btn_alts = driver.find_elements(By.XPATH, "//button[.//span[contains(text(),'Email')]] | //a[.//span[contains(text(),'Email')]]")
            if email_btn_alts:
                safe_click(driver, email_btn_alts[0])
                time.sleep(0.5)
    except Exception:
        pass
    if not customer_has_email(driver):
        gui_print("No customer email found, skipping email step for this customer.", status="Skipped email")
        return False
    try:
        first_name = driver.find_element(
            By.XPATH,
            "//div[contains(@class,'deal-customer')]//span[contains(@class,'cust-name')]"
        ).text.strip().title().split()[0]
    except Exception:
        gui_print("Could not read customer name for e-mail.")
        return False
    subject = templates["email_subject"].format(
        customer_name=first_name,
        sender_name=sender_name
    )
    body = templates["email_body"].format(
        customer_name=first_name,
        sender_name=sender_name
    )
    try:
        _compose_email(driver, subject, body)
        gui_print("📧 Standard e-mail sent.")
        return True
    except Exception as exc:
        gui_print(f"Email send error: {exc}")
        logger.debug(traceback.format_exc())
        return False

def send_custom_email_message(driver):
    try:
        nav_email_tabs = driver.find_elements(By.XPATH, "//li[contains(@analyticsdetect,'CustomerAction|Navigate|Email') and not(contains(@class,'active'))]")
        if nav_email_tabs:
            safe_click(driver, nav_email_tabs[0])
            time.sleep(0.5)
        else:
            email_btn_alts = driver.find_elements(By.XPATH, "//button[.//span[contains(text(),'Email')]] | //a[.//span[contains(text(),'Email')]]")
            if email_btn_alts:
                safe_click(driver, email_btn_alts[0])
                time.sleep(0.5)
    except Exception:
        pass
    if not customer_has_email(driver):
        gui_print("No customer email found, skipping custom email step.", status="Skipped email")
        return False
    variant = choose_custom_email_template()
    if not variant:
        gui_print("Custom e-mail cancelled by user.")
        return False
    subj_key = f"custom_email_subject_{variant}"
    body_key = f"custom_email_body_{variant}"
    if not templates.get(subj_key) or not templates.get(body_key):
        gui_print("Selected e-mail template is empty – edit templates first.")
        return False
    try:
        first_name = driver.find_element(
            By.XPATH,
            "//div[contains(@class,'deal-customer')]//span[contains(@class,'cust-name')]"
        ).text.strip().title().split()[0]
    except Exception:
        gui_print("Could not read customer name for e-mail.")
        return False
    subject = templates[subj_key].format(
        customer_name=first_name,
        sender_name=sender_name
    )
    body = templates[body_key].format(
        customer_name=first_name,
        sender_name=sender_name
    )
    try:
        _compose_email(driver, subject, body)
        gui_print(f"📧 Custom e-mail ({variant}) sent.")
        return True
    except Exception as exc:
        gui_print(f"Custom e-mail error: {exc}")
        logger.debug(traceback.format_exc())
        return False

def _compose_email(driver, subject: str, body: str):
    try:
        nav_email_tabs = driver.find_elements(By.XPATH, "//li[contains(@analyticsdetect,'CustomerAction|Navigate|Email') and not(contains(@class,'active'))]")
        if nav_email_tabs:
            safe_click(driver, nav_email_tabs[0])
            time.sleep(0.3)
        else:
            email_btn_alts = driver.find_elements(By.XPATH, "//button[.//span[contains(text(),'Email')]] | //a[.//span[contains(text(),'Email')]]")
            if email_btn_alts:
                safe_click(driver, email_btn_alts[0])
                time.sleep(0.3)
    except Exception:
        pass
    try:
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH,
                "//li[@analyticsdetect='CustomerAction|Navigate|Email']"))
        ).click()
    except Exception:
        pass
    if not customer_has_email(driver):
        raise Exception("No valid email specified for this contact.")
    time.sleep(0.3)
    subj_box = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Subject']"))
    )
    subj_box.clear()
    subj_box.send_keys(subject)
    try:
        iframe = driver.find_element(By.XPATH, "//iframe[contains(@id,'_ifr')]")
        driver.switch_to.frame(iframe)
    except Exception:
        pass
    body_elem = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, "//body[@contenteditable='true']"))
    )
    body_elem.click()
    body_elem.send_keys(Keys.CONTROL + "a")
    body_elem.send_keys(Keys.BACKSPACE)
    body_elem.send_keys(body)
    driver.switch_to.default_content()
    send_btn = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH,
            "//button[@analyticsdetect='ComposeEmail|Send|Email']"))
    )
    safe_click(driver, send_btn)

def send_custom_text_message(driver):
    template_key = choose_custom_text_template()
    if not template_key:
        gui_print("Custom text cancelled by user.")
        return
    if not templates.get(template_key):
        gui_print("Selected text template is empty – edit templates first.")
        return
    WebDriverWait(driver, 7).until(
        EC.element_to_be_clickable((By.XPATH,
            "//li[@analyticsdetect='CustomerAction|Navigate|Text']"))
    ).click()
    time.sleep(0.4)
    if driver.find_elements(By.XPATH, "//h4[contains(text(),'Status: Opted out')]"):
        gui_print("Customer is opted-out of texts.")
        return
    try:
        opt_in = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((
                By.XPATH, "//button[@analyticsdetect='CustomerActions|OptIn|Text']"))
        )
        if opt_in:
            safe_click(driver, opt_in)
            gui_print("Customer opted-in for texting.")
            time.sleep(0.4)
    except Exception:
        pass
    send_btn = WebDriverWait(driver, 4).until(
        EC.presence_of_element_located((
            By.XPATH, "//button[@analyticsdetect='CustomerActions|Send|Text']"))
    )
    try:
        first_name = driver.find_element(
            By.XPATH,
            "//div[contains(@class,'deal-customer')]//span[contains(@class,'cust-name')]"
        ).text.strip().title().split()[0]
    except Exception:
        first_name = ""
    textarea = WebDriverWait(driver, 5).until(
        EC.visibility_of_element_located((
            By.XPATH, "//textarea[contains(@class,'emoji-input-action-text')]"))
    )
    if not textarea.get_attribute("value").strip():
        textarea.send_keys(
            templates[template_key].format(
                customer_name=first_name,
                sender_name=sender_name
            )
        )
    safe_click(driver, send_btn)
    gui_print(f"📲 Custom text ({template_key[-1]}) sent.")

def send_text_message(driver):
    WebDriverWait(driver, 7).until(
        EC.element_to_be_clickable((By.XPATH,
            "//li[@analyticsdetect='CustomerAction|Navigate|Text']"))
    ).click()
    time.sleep(0.4)
    if driver.find_elements(By.XPATH, "//h4[contains(text(),'Status: Opted out')]"):
        gui_print("Customer is opted-out of texts.")
        return
    try:
        opt_in = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((
                By.XPATH, "//button[@analyticsdetect='CustomerActions|OptIn|Text']"))
        )
        if opt_in:
            safe_click(driver, opt_in)
            gui_print("Customer opted-in for texting.")
            time.sleep(0.4)
    except Exception:
        pass
    send_btn = WebDriverWait(driver, 4).until(
        EC.presence_of_element_located((
            By.XPATH, "//button[@analyticsdetect='CustomerActions|Send|Text']"))
    )
    try:
        first_name = driver.find_element(
            By.XPATH,
            "//div[contains(@class,'deal-customer')]//span[contains(@class,'cust-name')]"
        ).text.strip().title().split()[0]
    except Exception:
        first_name = ""
    textarea = WebDriverWait(driver, 5).until(
        EC.visibility_of_element_located((
            By.XPATH, "//textarea[contains(@class,'emoji-input-action-text')]"))
    )
    if not textarea.get_attribute("value").strip():
        textarea.send_keys(
            templates["standard_text"].format(
                customer_name=first_name,
                sender_name=sender_name
            )
        )
    safe_click(driver, send_btn)
    gui_print("📲 Standard text sent.")

def choose_radio_dialog(title: str, prompt: str, options: list[tuple[str, str]]) -> str | None:
    top = tk.Toplevel(root)
    top.title(title)
    top.geometry("340x200")
    top.grab_set()
    choice = tk.StringVar(value=options[0][0])
    tk.Label(top, text=prompt, pady=6).pack(anchor="w", padx=10)
    for key, label in options:
        tk.Radiobutton(top, text=label, variable=choice, value=key).pack(anchor="w", padx=22)
    def ok():
        top.destroy()
    def cancel():
        choice.set("")
        top.destroy()
    btn_frm = ttk.Frame(top); btn_frm.pack(pady=10)
    ttk.Button(btn_frm, text="Send", width=8, command=ok).pack(side=tk.LEFT, padx=6)
    ttk.Button(btn_frm, text="Cancel", width=8, command=cancel).pack(side=tk.LEFT, padx=6)
    root.wait_window(top)
    return choice.get() or None

def choose_custom_text_template() -> str | None:
    opts = [("custom_text_A", "Template A"),
            ("custom_text_B", "Template B"),
            ("custom_text_C", "Template C")]
    return choose_radio_dialog("Choose Custom Text",
                              "Select the custom TEXT template to send:",
                              opts)

def choose_custom_email_template() -> str | None:
    opts = [("A", "Template A"),
            ("B", "Template B"),
            ("C", "Template C")]
    return choose_radio_dialog("Choose Custom Email",
                              "Select the custom E-MAIL template to send:",
                              opts)

def threaded(fn):
    return lambda *a, **kw: threading.Thread(target=fn, args=a, kwargs=kw,
        daemon=True).start()

@threaded
def claim_customer():
    gui_print("\n--- Claim + Edit Task ---", status="Claim+Edit")
    drv = get_chrome_driver()
    if drv:
        try:
            ensure_claimed_and_update_task(drv)
        except Exception as exc:
            gui_print(f"Claim+Edit error: {exc}")
            logger.debug(traceback.format_exc())
        status_var.set("Ready")

@threaded
def send_text_wrapper():
    gui_print("\n--- Send Standard Text ---", status="Send Text")
    drv = get_chrome_driver()
    if drv:
        try:
            ensure_claimed_only(drv)
            send_text_message(drv)
        except Exception as exc:
            gui_print(f"Text flow error: {exc}")
            logger.debug(traceback.format_exc())
        status_var.set("Ready")

@threaded
def send_custom_text_wrapper():
    gui_print("\n--- Send Custom Text ---", status="Custom Text")
    drv = get_chrome_driver()
    if drv:
        try:
            ensure_claimed_only(drv)
            send_custom_text_message(drv)
        except Exception as exc:
            gui_print(f"Custom text flow error: {exc}")
            logger.debug(traceback.format_exc())
        status_var.set("Ready")

@threaded
def send_email_wrapper():
    gui_print("\n--- Send Standard Email ---", status="Send Email")
    drv = get_chrome_driver()
    if drv:
        try:
            ensure_claimed_only(drv)
            res = send_email_message(drv)
            if res is False:
                gui_print("No email found. Email step skipped.", status="No Email")
        except Exception as exc:
            gui_print(f"Email flow error: {exc}")
            logger.debug(traceback.format_exc())
        status_var.set("Ready")

@threaded
def send_custom_email_wrapper():
    gui_print("\n--- Send Custom Email ---", status="Custom Email")
    drv = get_chrome_driver()
    if drv:
        try:
            ensure_claimed_only(drv)
            res = send_custom_email_message(drv)
            if res is False:
                gui_print("No email found. Custom email step skipped.", status="No Email")
        except Exception as exc:
            gui_print(f"Custom e-mail flow error: {exc}")
            logger.debug(traceback.format_exc())
        status_var.set("Ready")

@threaded
def full_outreach_wrapper():
    gui_print("\n--- Full Outreach (Claim + Email + Text) ---", status="Full Outreach")
    drv = get_chrome_driver()
    if drv:
        try:
            ensure_claimed_only(drv)
            res = send_email_message(drv)
            if res is False:
                gui_print("No email found. Skipping to text.", status="No Email")
            send_text_message(drv)
            gui_print("🏁 Outreach finished.")
        except Exception as exc:
            gui_print(f"Outreach error: {exc}")
            logger.debug(traceback.format_exc())
        status_var.set("Ready")

# ---- AUTO Touchpoint+Email+Text+Next mode: will process all customers until carousel ends or stopped ----
@threaded
def auto_touchpoint_email_text_next():
    gui_print("Auto Touchpoint+Email+Text+Next started (STOP/ctrl+alt+Q to halt).", status="Auto Touchpoint+Email+Text+Next")
    drv = get_chrome_driver()
    if not drv:
        status_var.set("Ready")
        return
    auto_stop_event.clear()
    while not auto_stop_event.is_set():
        try:
            # Ensure claimed for each customer
            if not is_customer_claimed(drv):
                gui_print("Customer not claimed; claiming ...")
                if not click_claim_and_replace(drv):
                    gui_print("Could not claim. Skipping this customer.", status="Claim failed")
                    # Move to next and continue loop!
                    next_btns = drv.find_elements(By.XPATH, "//*[contains(@analyticsdetect,'Carousel|Navigate|Right')]")
                    if next_btns:
                        safe_click(drv, next_btns[0])
                        gui_print("➡️ Moved to next customer after claim fail.", status="Next customer")
                        time.sleep(1.0)
                        continue
                    else:
                        gui_print("No more customers in carousel/list. Stopping.", status="No more customers")
                        break
            else:
                gui_print("Customer already claimed.")

            # Edit task to Touchpoint (same as F8 logic)
            ok = set_task_to_touchpoint(drv)
            if not ok:
                gui_print("Could not set task to Touchpoint (see log).")

            # Send e-mail if available
            sent_email = False
            if customer_has_email(drv):
                try:
                    sent_email = send_email_message(drv)
                    if sent_email:
                        gui_print("Standard e-mail sent.")
                except Exception as ex:
                    gui_print(f"Email error: {ex}")
            else:
                gui_print("No email for this customer.")

            # Send text if possible
            try:
                send_text_message(drv)
            except Exception as ex:
                gui_print(f"Text send error: {ex}")

            # Advance to next customer (carousel)
            next_btns = drv.find_elements(By.XPATH, "//*[contains(@analyticsdetect,'Carousel|Navigate|Right')]")
            if next_btns:
                safe_click(drv, next_btns[0])
                gui_print("➡️ Moved to next customer via carousel.", status="Next customer")
                time.sleep(1.0)
            else:
                gui_print("No more customers in carousel/list. Stopping.", status="No more customers")
                break
        except Exception as exc:
            gui_print(f"Auto Touchpoint+Email+Text+Next error: {exc}")
            logger.debug(traceback.format_exc())
            time.sleep(2)
    gui_print("Auto Touchpoint+Email+Text+Next stopped.")
    status_var.set("Ready")

@threaded
def auto_process_customers():
    gui_print("Auto-process started (Ctrl+Alt+Q or STOP button to stop).", status="Auto-process")
    drv = get_chrome_driver()
    if not drv:
        status_var.set("Ready")
        return
    auto_stop_event.clear()
    while not auto_stop_event.is_set():
        try:
            claimed = is_customer_claimed(drv)
            if not claimed:
                try:
                    claimed_now = click_claim_and_replace(drv)
                    if claimed_now:
                        edit_task_after_claim(drv)
                except Exception as exc:
                    gui_print(f"Claiming error: {exc}")
            else:
                try:
                    edit_task_after_claim(drv)
                except Exception as exc:
                    gui_print(f"Task edit error: {exc}")
            email_sent = False
            email_available = customer_has_email(drv)
            if email_available:
                try:
                    email_sent = send_email_message(drv)
                    if email_sent:
                        gui_print("Email sent for this customer.")
                except Exception as exc:
                    gui_print(f"Email error: {exc}")
                    email_sent = False
            else:
                gui_print("No email found for customer. Skipping email.")
            try:
                send_text_message(drv)
            except Exception as exc:
                gui_print(f"Text error: {exc}")
            next_btns = drv.find_elements(By.XPATH, "//*[contains(@analyticsdetect,'Carousel|Navigate|Right')]")
            if next_btns:
                safe_click(drv, next_btns[0])
                gui_print("➡️ Moved to next customer via carousel.")
                time.sleep(1.2)
            else:
                gui_print("No more customers in carousel/list. Halting auto-process.", status="Auto-process stopped")
                break
        except Exception as exc:
            gui_print(f"Auto-process error (outer loop): {exc}")
            logger.debug(traceback.format_exc())
            time.sleep(2)
    gui_print("Auto-process stopped.")
    status_var.set("Ready")

@threaded
def auto_text_only_customers():
    gui_print("Auto-text-only started (Ctrl+Alt+Q or STOP button to stop).", status="Auto-text")
    drv = get_chrome_driver()
    if not drv:
        status_var.set("Ready")
        return
    auto_stop_event.clear()
    while not auto_stop_event.is_set():
        try:
            if not is_customer_claimed(drv):
                gui_print("Auto: Not claimed, claiming first.")
                click_claim_and_replace(drv)
            else:
                gui_print("Auto: Already claimed.")
            WebDriverWait(drv, 7).until(
                EC.element_to_be_clickable((By.XPATH,
                    "//li[@analyticsdetect='CustomerAction|Navigate|Text']"))
            ).click()
            time.sleep(0.3)
            if drv.find_elements(By.XPATH, "//h4[contains(text(),'Status: Opted out')]"):
                gui_print("Auto: Customer is opted-out of texts. Skipping this customer.")
                next_btns = drv.find_elements(By.XPATH,
                    "//*[contains(@analyticsdetect,'Carousel|Navigate|Right')]")
                if next_btns:
                    safe_click(drv, next_btns[0])
                    gui_print("Auto: ➡️ Moved to next customer via carousel.")
                    time.sleep(1.0)
                else:
                    gui_print("No more customers in carousel/list. Halting auto-text-only.", status="Auto-text-only stopped")
                    break
                continue
            opt_in_sent = False
            try:
                optin_btns = drv.find_elements(By.XPATH, "//button[@analyticsdetect='CustomerActions|OptIn|Text'] | //button[contains(.,'RESEND')] | //button[contains(.,'Resend')]")
                for btn in optin_btns:
                    if btn.is_displayed() and btn.is_enabled():
                        safe_click(drv, btn)
                        gui_print("Auto: Opt-in or RESEND clicked for texting.")
                        opt_in_sent = True
                        time.sleep(0.5)
                        break
            except Exception:
                pass
            if opt_in_sent:
                next_btns = drv.find_elements(By.XPATH, "//*[contains(@analyticsdetect,'Carousel|Navigate|Right')]")
                if next_btns:
                    safe_click(drv, next_btns[0])
                    gui_print("Auto: ➡️ Moved to next customer via carousel.")
                    time.sleep(1.0)
                else:
                    gui_print("No more customers in carousel/list. Halting auto-text-only.", status="Auto-text-only stopped")
                    break
                continue
            send_btn = WebDriverWait(drv, 2.5).until(
                EC.presence_of_element_located((
                    By.XPATH, "//button[@analyticsdetect='CustomerActions|Send|Text']"))
            )
            try:
                first_name = drv.find_element(
                    By.XPATH,
                    "//div[contains(@class,'deal-customer')]//span[contains(@class,'cust-name')]"
                ).text.strip().title().split()[0]
            except Exception:
                first_name = ""
            textarea = WebDriverWait(drv, 3).until(
                EC.visibility_of_element_located((
                    By.XPATH, "//textarea[contains(@class,'emoji-input-action-text')]"))
            )
            if not textarea.get_attribute("value").strip():
                textarea.send_keys(
                    templates["standard_text"].format(
                        customer_name=first_name,
                        sender_name=sender_name
                    )
                )
            safe_click(drv, send_btn)
            gui_print("Auto: 📲 Standard text sent.")
            next_btns = drv.find_elements(By.XPATH,
                "//*[contains(@analyticsdetect,'Carousel|Navigate|Right')]")
            if next_btns:
                safe_click(drv, next_btns[0])
                gui_print("Auto: ➡️ Moved to next customer via carousel.")
                time.sleep(1.0)
            else:
                gui_print("No more customers in carousel/list. Halting auto-text-only.", status="Auto-text-only stopped")
                break
        except Exception as exc:
            gui_print(f"Auto-text-only error: {exc}")
            time.sleep(1)
    gui_print("Auto-text-only stopped.")
    status_var.set("Ready")

@threaded
def auto_email_only_customers():
    gui_print("Auto-email-only started (Ctrl+Alt+Q or STOP button to stop).", status="Auto-email")
    drv = get_chrome_driver()
    if not drv:
        status_var.set("Ready")
        return
    auto_stop_event.clear()
    while not auto_stop_event.is_set():
        try:
            if not is_customer_claimed(drv):
                gui_print("Auto: Not claimed, claiming first.")
                click_claim_and_replace(drv)
            else:
                gui_print("Auto: Already claimed.")
            if not customer_has_email(drv):
                gui_print("Auto: No email found for customer. Skipping to next.", status="No Email")
            else:
                send_email_message(drv)
            next_btns = drv.find_elements(By.XPATH,
                "//*[contains(@analyticsdetect,'Carousel|Navigate|Right')]")
            if next_btns:
                safe_click(drv, next_btns[0])
                gui_print("Auto: ➡️ Moved to next customer via carousel.")
                time.sleep(1.0)
            else:
                gui_print("No more customers in carousel/list. Halting auto-email-only.", status="Auto-email-only stopped")
                break
        except Exception as exc:
            gui_print(f"Auto-email-only error: {exc}")
            time.sleep(1)
    gui_print("Auto-email-only stopped.")
    status_var.set("Ready")

@threaded
def auto_outreach_claimed_only():
    gui_print("Auto Claimed-Only Outreach started (STOP to halt).", status="Auto-Claimed-Outreach")
    drv = get_chrome_driver()
    if not drv:
        status_var.set("Ready")
        return
    auto_stop_event.clear()
    while not auto_stop_event.is_set():
        try:
            if not is_customer_claimed(drv):
                gui_print("Not claimed, skipping (this auto mode only processes claimed).")
                next_btns = drv.find_elements(By.XPATH, "//*[contains(@analyticsdetect,'Carousel|Navigate|Right')]")
                if next_btns:
                    safe_click(drv, next_btns[0])
                    gui_print("➡️ Moved to next customer via carousel.")
                    time.sleep(1.2)
                else:
                    gui_print("No more customers in carousel/list. Halting.", status="Auto-Claimed-Outreach stopped")
                    break
                continue
            email_sent = False
            if customer_has_email(drv):
                try:
                    email_sent = send_email_message(drv)
                    if email_sent:
                        gui_print("Email sent for claimed customer.")
                except Exception as exc:
                    gui_print(f"Email error: {exc}")
            try:
                send_text_message(drv)
            except Exception as exc:
                gui_print(f"Text error: {exc}")
            next_btns = drv.find_elements(By.XPATH, "//*[contains(@analyticsdetect,'Carousel|Navigate|Right')]")
            if next_btns:
                safe_click(drv, next_btns[0])
                gui_print("➡️ Moved to next customer via carousel.")
                time.sleep(1.2)
            else:
                gui_print("No more customers in carousel/list. Halting.", status="Auto-Claimed-Outreach stopped")
                break
        except Exception as exc:
            gui_print(f"Error in auto-claimed outreach: {exc}")
            logger.debug(traceback.format_exc())
            time.sleep(1)
    gui_print("Auto Claimed-Only Outreach stopped.")
    status_var.set("Ready")

def get_numeric_version(v: str) -> float:
    try: return float(v)
    except Exception: return 0.0

def get_current_version() -> str:
    try:
        return resource_path("version.txt").read_text(encoding="utf-8").strip()
    except Exception:
        return DEFAULT_VERSION

def check_for_update():
    api = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases/latest"
    gui_print(f"Checking updates at {api} ...", status="Checking update")
    try:
        data = requests.get(api, timeout=10).json()
    except Exception as exc:
        gui_print(f"Update check failed: {exc}")
        status_var.set("Ready")
        return False
    remote = data.get("tag_name", "").lstrip("v")
    local = get_current_version()
    if get_numeric_version(remote) <= get_numeric_version(local):
        gui_print("No update available.")
        status_var.set("Ready")
        return False
    if not messagebox.askyesno("Update", f"Update {remote} available. Download?"):
        status_var.set("Ready")
        return False
    zip_url = next((a["browser_download_url"]
        for a in data.get("assets", [])
        if a["name"].endswith(".zip")), None)
    if not zip_url:
        gui_print("Release has no .zip asset.")
        status_var.set("Ready")
        return False
    gui_print(f"Downloading {zip_url} ...", status="Downloading update")
    try:
        zdata = requests.get(zip_url, timeout=30).content
        zf = zipfile.ZipFile(io.BytesIO(zdata))
    except Exception as exc:
        gui_print(f"Download failed: {exc}")
        status_var.set("Ready")
        return False
    tmp = USER_DATA_DIR / "update_tmp"
    forcibly_remove_folder(tmp)
    tmp.mkdir(exist_ok=True)
    zf.extractall(tmp)
    for src in tmp.rglob("*"):
        if src.is_file():
            dst = Path.cwd() / src.relative_to(tmp)
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
    gui_print("Update applied - restart program.")
    forcibly_remove_folder(tmp)
    status_var.set("Ready")
    return True

def manual_update_check():
    if check_for_update():
        root.quit()

def edit_templates_wrapper():
    threading.Thread(target=_edit_templates_worker, daemon=True).start()

def _edit_templates_worker():
    def do_save():
        for k, w in widgets.items():
            templates[k] = w.get("1.0", tk.END).strip() if isinstance(w, tk.Text) else w.get().strip()
        save_templates()
        top.destroy()
    top = tk.Toplevel(root)
    top.title("Template Editor")
    top.geometry("780x600")
    canvas = tk.Canvas(top)
    scrollbar = tk.Scrollbar(top, orient="vertical", command=canvas.yview)
    scroll_frame = ttk.Frame(canvas)
    scroll_frame.bind("<Configure>",
                     lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    widgets = {}
    for key in sorted(templates.keys()):
        frm = ttk.Frame(scroll_frame); frm.pack(fill='x', pady=2, padx=6)
        ttk.Label(frm, text=key, width=30, anchor='w').pack(side=tk.LEFT)
        if key.endswith("_body") or key.endswith("_text"):
            txt = tk.Text(frm, width=80, height=6); txt.insert(tk.END, templates[key])
        else:
            txt = ttk.Entry(frm, width=80); txt.insert(0, templates[key])
        txt.pack(side=tk.LEFT, expand=True, fill='x')
        widgets[key] = txt
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    ttk.Button(top, text="Save", command=do_save).pack(pady=8)
    top.grab_set()
    root.wait_window(top)

def stop_auto_process_gui():
    auto_stop_event.set()
    gui_print("Auto-process stop requested (via button).", status="Auto-process stop requested")

def start_hotkey_thread():
    threading.Thread(target=_register_hotkeys, daemon=True).start()

def _register_hotkeys():
    keyboard.add_hotkey("F8", claim_customer)
    keyboard.add_hotkey("ctrl+shift+F8", claim_customer)
    keyboard.add_hotkey("F9", send_text_wrapper)
    keyboard.add_hotkey("ctrl+alt+t", send_custom_text_wrapper)
    keyboard.add_hotkey("F10", full_outreach_wrapper)
    keyboard.add_hotkey("ctrl+alt+c", full_outreach_wrapper)
    keyboard.add_hotkey("ctrl+alt+e", send_email_wrapper)
    keyboard.add_hotkey("F11", send_custom_email_wrapper)
    keyboard.add_hotkey("ctrl+alt+shift+e", send_custom_email_wrapper)
    keyboard.add_hotkey("ctrl+alt+s", auto_process_customers)
    keyboard.add_hotkey("ctrl+alt+x", auto_text_only_customers)
    keyboard.add_hotkey("ctrl+alt+m", auto_email_only_customers)
    keyboard.add_hotkey("ctrl+alt+z", auto_outreach_claimed_only) # new hotkey
    keyboard.add_hotkey("ctrl+alt+q", auto_stop_event.set)
    keyboard.add_hotkey("ctrl+alt+p", edit_templates_wrapper)
    keyboard.add_hotkey("ctrl+alt*u", manual_update_check)
    keyboard.add_hotkey("ctrl+alt+n", auto_touchpoint_email_text_next)
    while root and root.winfo_exists():
        time.sleep(0.1)

def build_gui():
    global root, log_text, status_var
    root = tk.Tk()
    root.title("DriveCentric TaskClaim")
    root.geometry("1040x720")
    style = ttk.Style(root)
    if style.theme_use() == "vista":
        style.configure("TButton", padding=6)
    top = ttk.Frame(root); top.pack(side=tk.TOP, fill=tk.X, pady=4)
    top2 = ttk.Frame(root); top2.pack(side=tk.TOP, fill=tk.X)

    import platform
    def open_templates_file():
        file = TEMPLATE_FILE.resolve()
        if not file.exists():
            messagebox.showerror("File Not Found", f"{file} does not exist.")
            return
        if platform.system() == "Windows":
            os.startfile(str(file))
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", str(file)])
        else:
            subprocess.Popen(["xdg-open", str(file)])

    def add_btn(parent, lbl, cmd, w=20):
        ttk.Button(parent, text=lbl, width=w, command=cmd).pack(side=tk.LEFT, padx=3, pady=3)
    add_btn(top, "Launch Chrome", lambda: threading.Thread(
        target=launch_chrome, daemon=True).start(), 16)
    add_btn(top, "Claim Only", claim_only_customer)
    add_btn(top, "Claim+Edit (F8)", claim_customer)
    add_btn(top, "Std Text (F9)", send_text_wrapper)
    add_btn(top, "Custom Text (Ctrl+Alt+T)", send_custom_text_wrapper)
    add_btn(top, "Std Email (Ctrl+Alt+E)", send_email_wrapper)
    add_btn(top, "Custom Email (F11)", send_custom_email_wrapper)
    ttk.Button(top, text="Open Templates File", width=22, command=open_templates_file).pack(side=tk.LEFT, padx=3, pady=3)
    add_btn(top2, "Auto Touchpoint+Email+Text+Next", auto_touchpoint_email_text_next, 34)
    add_btn(top2, "Full Outreach (F10)", full_outreach_wrapper)
    add_btn(top2, "Templates (Ctrl+Alt+P)", edit_templates_wrapper)
    add_btn(top2, "Auto Process (Ctrl+Alt+S)", auto_process_customers)
    add_btn(top2, "Auto Text Only (Ctrl+Alt+X)", auto_text_only_customers)
    add_btn(top2, "Auto Email Only (Ctrl+Alt+M)", auto_email_only_customers)
    add_btn(top2, "Auto Outreach (claimed only)", auto_outreach_claimed_only, 24)
    add_btn(top2, "Check Updates", manual_update_check)
    add_btn(top2, "STOP Auto Process", stop_auto_process_gui, 16)
    add_btn(top2, "Exit", root.quit, 10)
    log_text = scrolledtext.ScrolledText(root, state='disabled', wrap='word', height=25)
    log_text.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
    status_var = tk.StringVar(value="Ready")
    status = ttk.Label(root, textvariable=status_var, relief=tk.SUNKEN, anchor='w')
    status.pack(fill=tk.X, side=tk.BOTTOM)
    ttk.Label(root, text=f"{WATERMARK_ICON} {WATERMARK_TEXT} {WATERMARK_ICON}",
        foreground="gray50", font=("Segoe UI", 9, "italic"))\
        .pack(side=tk.BOTTOM, pady=2)

def main():
    global sender_name
    build_gui()
    load_templates()
    sender_name = gui_login()
    if not sender_name:
        root.quit(); return
    gui_print(f"Logged in as {sender_name}.", status="Ready")
    start_hotkey_thread()
    gui_print("Ready. Use buttons or hot-keys. Full Outreach -> F10.")
    root.mainloop()

if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        fatal_popup(f"{exc}\n\n{traceback.format_exc()}")

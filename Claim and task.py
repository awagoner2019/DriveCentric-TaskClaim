import subprocess
import time
import keyboard
import traceback
import os
import datetime
import json
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Allowed credentials: any of these users can use either password.
ALLOWED_USERS = {"Aaron Wagoner", "Nate Floyd", "Jean Luc", "Daulton Gentry"}
ALLOWED_PASSWORDS = {"Acura2025", "ADMIN"}

# Global sender name (set during login)
sender_name = ""

# Global templates dictionary
templates = {
    "standard_text": "Hello, {customer_name}! This is {sender_name} with Acura of Springfield. Just wanted to check in — we’ve got some strong offers and fresh inventory rolling in, and I didn’t want you to miss out if you’ve been thinking about upgrading. Let me know if you're open to a quick chat!",
    "custom_text_A": "Hello, {customer_name}! We've got new offers and inventory updates at Acura of Springfield. Let me know if you'd like more details or a personalized overview.",
    "custom_text_B": "Hi, {customer_name}! I'm following up on your recent inquiry. How can I assist you further? Are you interested in a test drive or more details?",
    "custom_text_C": "Urgent: Limited-time offers are available now, {customer_name}. Please contact us immediately for details.",
    "email_subject": "Your Inquiry from Acura of Springfield",
    "email_body": "Hello, {customer_name},\n\nThank you for reaching out to us at Acura of Springfield. We offer a diverse range of vehicles along with competitive financing and exclusive promotions designed to meet your needs. Please review our latest inventory and feel free to contact us if you have any questions or wish to schedule a visit. We look forward to assisting you further.\n\nBest regards,"
}

# File to store templates
TEMPLATE_FILE = "templates.json"

# Set up logging
LOG_FILENAME = "drivecentric_log.txt"
logging.basicConfig(
    filename=LOG_FILENAME,
    filemode='a',
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.DEBUG
)
logger = logging.getLogger("drivecentric")
logger.info("Program started.")

# ------------------------------
# Utility Functions
# ------------------------------

def get_windows_date():
    """
    Uses Windows 'date /t' to return the current date in MM/DD/YY format.
    """
    try:
        date_str = os.popen("date /t").read().strip()
        if " " in date_str:
            date_str = date_str.split()[-1]
        parts = date_str.split('/')
        if len(parts) == 3:
            mm, dd, yyyy = parts
            yy = yyyy[-2:]
            result = f"{mm}/{dd}/{yy}"
            logger.debug(f"Obtained Windows date: {result}")
            return result
        return date_str
    except Exception as e:
        logger.error("Error obtaining Windows date: " + str(e))
        return datetime.datetime.now().strftime("%m/%d/%y")

def safe_click(driver, element):
    """
    Attempts to click an element. If a normal click fails, uses JavaScript.
    """
    try:
        element.click()
    except Exception:
        driver.execute_script("arguments[0].click();", element)

# ------------------------------
# Template Loading and Editing Functions
# ------------------------------

def load_templates():
    global templates
    if os.path.exists(TEMPLATE_FILE):
        try:
            with open(TEMPLATE_FILE, "r") as f:
                templates = json.load(f)
            logger.info("Templates loaded from file.")
        except Exception as e:
            logger.error("Error loading templates: " + str(e))
    else:
        logger.info("Template file not found; using default templates.")
        ans = input("No template file found. Would you like to edit the default templates? (Y/N): ").strip().lower()
        if ans == "y":
            manual_edit_templates(save_after=True)
        else:
            save_templates()

def save_templates():
    try:
        with open(TEMPLATE_FILE, "w") as f:
            json.dump(templates, f, indent=4)
        logger.info("Templates saved to file.")
    except Exception as e:
        logger.error("Error saving templates: " + str(e))

def manual_edit_templates(save_after=False):
    global templates
    print("\n--- Template Editor ---")
    print("Press Enter without typing anything to keep the current template.\n")
    for key in templates:
        print(f"Current {key.replace('_', ' ').title()}:")
        print(templates[key])
        new_val = input(f"Enter new value for {key.replace('_', ' ').title()} (or press Enter to keep unchanged):\n")
        if new_val.strip():
            templates[key] = new_val.strip()
            logger.info(f"Template '{key}' updated.")
    if save_after:
        save_templates()
        print("Templates updated and saved.\n")
    else:
        print("Templates updated in memory.\n")
    logger.info("Manual template editing complete.")

def edit_templates_wrapper():
    manual_edit_templates(save_after=True)

# ------------------------------
# Authentication and Watermark
# ------------------------------

def login():
    while True:
        user = input("Enter your username: ").strip()
        if user not in ALLOWED_USERS:
            print("Unauthorized user. Please try again.\n")
            logger.warning(f"Unauthorized login attempt with username: {user}")
            continue
        password = input("Enter your password: ").strip()
        if password not in ALLOWED_PASSWORDS:
            print("Incorrect password. Unauthorized access. Exiting in 5 seconds.")
            logger.error(f"Incorrect password attempt for user: {user}")
            time.sleep(5)
            exit(1)
        logger.info(f"User '{user}' logged in successfully.")
        return user

def print_watermark():
    watermark = """
*********************************************************************
*                                                                   *
*                This program is developed by Aaron Wagoner         *
*                                                                   *
*    Redistribution or modification without explicit consent is     *
*                   strictly prohibited.                          *
*                                                                   *
*           Unauthorized use, copying, or distribution               *
*            of this software is a violation of copyright laws.     *
*                                                                   *
*********************************************************************
"""
    print(watermark)
    logger.info("Watermark displayed.")

# ------------------------------
# Ensure Customer is Claimed
# ------------------------------

def ensure_customer_claimed():
    options = Options()
    options.debugger_address = "127.0.0.1:9222"
    try:
        driver = webdriver.Chrome(options=options)
        logger.info("Attached to Chrome for ensuring claim.")
    except Exception as e:
        logger.error("Error attaching for ensuring claim: " + str(e))
        return
    try:
        new_deal_elements = driver.find_elements(By.XPATH, "//div[@analyticsdetect='Sidebar|Open|NewDeal']")
        if new_deal_elements:
            logger.info("Customer is already claimed (New Deal detected).")
            print("Customer is already claimed.")
        else:
            logger.info("Customer not claimed; running claim process.")
            print("Customer not claimed. Running claim process...")
            driver.quit()
            claim_customer()
            time.sleep(5)
            return
    except Exception as e:
        logger.error("Error in ensure_customer_claimed: " + str(e))
    driver.quit()

# ------------------------------
# Lead Claim & Task Update Functions
# ------------------------------

def launch_chrome():
    chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    debugging_port = "9222"
    user_data_dir = r"C:\TempChromeProfile"
    command = [
        chrome_path,
        f"--remote-debugging-port={debugging_port}",
        f"--user-data-dir={user_data_dir}"
    ]
    logger.info("Launching Chrome in remote debugging mode...")
    print("Launching Chrome in remote debugging mode...")
    subprocess.Popen(command)
    time.sleep(5)
    print("Chrome launched. Please log into DriveCentric and navigate to the Claim Customer page.")
    logger.info("Chrome launched; navigate to Claim Customer page.")

def edit_task_after_claim(driver):
    try:
        logger.info("Waiting for 'Edit' option in Task modal...")
        edit_li = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//li[@analyticsdetect='Timeline|PerformAction|TaskToDo' and contains(., 'Edit')]"))
        )
        safe_click(driver, edit_li)
        logger.info("'Edit' clicked in Task modal.")
        print("✅ 'Edit' clicked.")
        time.sleep(2)
        
        logger.info("Waiting for Phone dropdown...")
        phone_dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'action-list__button') and .//span[contains(text(),'Phone')]]"))
        )
        safe_click(driver, phone_dropdown)
        logger.info("Phone dropdown clicked.")
        print("Clicked Phone dropdown.")
        time.sleep(1)
        
        logger.info("Waiting for 'Touchpoint' option...")
        touchpoint_option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'drc-action-list-item') and .//span[contains(text(),'Touchpoint')]]"))
        )
        safe_click(driver, touchpoint_option)
        logger.info("'Touchpoint' selected.")
        print("Selected 'Touchpoint'.")
        time.sleep(1)
        
        logger.info("Waiting for date input field...")
        date_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Select a date']"))
        )
        windows_date = get_windows_date()
        current_value = date_input.get_attribute("value").strip()
        logger.info(f"Date in input: '{current_value}', system date: '{windows_date}'")
        print(f"Windows system date: {windows_date}")
        if current_value == windows_date:
            logger.info("Date already set correctly.")
            print(f"Date already set to today ({windows_date}).")
        else:
            logger.info("Updating date field.")
            print(f"Current date is '{current_value}'. Updating to {windows_date}.")
            date_input.send_keys(Keys.CONTROL + "a")
            date_input.send_keys(Keys.BACKSPACE)
            date_input.send_keys(windows_date)
        time.sleep(1)
        
        logger.info("Waiting for Save button...")
        try:
            save_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.drc-button.kind-filled.type-primary.size-medium.state-default"))
            )
            safe_click(driver, save_button)
            logger.info("Save button clicked successfully.")
            print("✅ Save button clicked!")
        except Exception as ex:
            logger.error("Error clicking Save normally. Trying JS fallback: " + str(ex))
            print("❗ Error clicking Save normally. Trying JS fallback...")
            try:
                save_button = driver.find_element(By.CSS_SELECTOR, "button.drc-button.kind-filled.type-primary.size-medium.state-default")
                driver.execute_script("arguments[0].click();", save_button)
                logger.info("Save button clicked via JS fallback.")
                print("✅ Save button clicked via JS fallback!")
            except Exception as ex2:
                logger.error("JS fallback failed for Save button: " + str(ex2))
                print("❌ JS fallback failed to click Save.")
        time.sleep(2)
    except Exception as e:
        logger.error("Error editing task: " + str(e))
        print("❗ Error editing task:")
        logger.debug(traceback.format_exc())
        print(traceback.format_exc())

def click_claim_and_replace(driver):
    try:
        new_deal_elements = driver.find_elements(By.XPATH, "//div[@analyticsdetect='Sidebar|Open|NewDeal']")
        if new_deal_elements:
            logger.info("Customer already claimed (New Deal detected).")
            print("New Deal element found. Customer is already claimed. Proceeding to task edit...")
            edit_task_after_claim(driver)
            return
    except Exception as e:
        logger.error("Error checking for New Deal element: " + str(e))
    
    print("Looking for 'Claim Customer' button...")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "act-button")))
    buttons = driver.find_elements(By.CLASS_NAME, "act-button")
    claimed = False
    for btn in buttons:
        if "ClaimCustomer" in (btn.get_attribute("analyticsdetect") or ""):
            safe_click(driver, btn)
            logger.info("'Claim Customer' button clicked.")
            print("✅ 'Claim Customer' clicked.")
            claimed = True
            break
    if not claimed:
        logger.error("'Claim Customer' button not found. Aborting claim process.")
        print("❌ 'Claim Customer' button not found. Aborting.")
        return

    logger.info("Waiting for radio inputs...")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@type='radio']")))
    radios = driver.find_elements(By.XPATH, "//input[@type='radio']")
    if radios:
        radios[0].click()
        logger.info("Top salesperson radio selected.")
        print("👤 Selected top salesperson.")
    else:
        logger.error("No radio inputs found. Aborting claim process.")
        print("⚠️ No radio inputs found. Aborting.")
        return

    selector = "button.drc-button.kind-filled.type-primary.size-medium.state-default"
    print("Looking for final 'Claim' button...")
    try:
        final_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
        safe_click(driver, final_btn)
        logger.info("Final 'Claim' button clicked.")
        print("🎯 Final 'Claim' button clicked!")
    except Exception as e:
        logger.error("Error clicking final Claim button: " + str(e))
        print("❗ Error clicking final Claim button:", e)
        driver.execute_script(f'''
            let btn = document.querySelector("{selector}");
            if (btn) {{
                btn.click();
                console.log("🎯 Final 'Claim' clicked via JS fallback!");
            }} else {{
                console.log("❌ Final Claim button not found via JS fallback.");
            }}
        ''')
    time.sleep(2)
    logger.info("Proceeding to edit task.")
    print("Now editing the task...")
    edit_task_after_claim(driver)

def claim_customer():
    logger.info("Starting claim process.")
    print("\n--- Claim Process ---")
    options = Options()
    options.debugger_address = "127.0.0.1:9222"
    try:
        driver = webdriver.Chrome(options=options)
        logger.info("Attached to Chrome for claim process.")
    except Exception as e:
        logger.error("Error attaching to Chrome: " + str(e))
        print("❗ Error attaching to Chrome:", e)
        return
    try:
        logger.info(f"Current URL: {driver.current_url}")
        print("Current URL:", driver.current_url)
        click_claim_and_replace(driver)
    except Exception as e:
        logger.error("Error during claim process: " + str(e))
        print("❗ Error during claim process:", e)
        logger.debug(traceback.format_exc())
    # Do not quit driver.

# ------------------------------
# Texting Functions
# ------------------------------

def send_text_message(driver):
    """
    Sends a standard text message using the visible messaging area.
    Uses the standard text template from the global templates with {customer_name} and {sender_name}.
    """
    try:
        logger.info("Navigating to Text tab for standard text.")
        text_tab = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[@analyticsdetect='CustomerAction|Navigate|Text']"))
        )
        safe_click(driver, text_tab)
        time.sleep(1)
    except Exception as e:
        logger.error("Error navigating to Text tab: " + str(e))
        print("Couldn't navigate to the Text tab.", e)

    try:
        logger.info("Checking for opt-out status in standard text.")
        if driver.find_elements(By.XPATH, "//h4[contains(text(), 'Status: Opted out')]"):
            logger.warning("Customer opted out of texts; aborting standard text.")
            print("Customer opted out of texts. Aborting standard text.")
            return
    except Exception as e:
        logger.error("Error checking opt-out: " + str(e))
        print("Error checking opt-out status:", e)

    send_btn = None
    try:
        logger.info("Checking for Send button for standard text.")
        send_btn = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.XPATH, "//button[@analyticsdetect='CustomerActions|Send|Text']"))
        )
        logger.info("Send button detected.")
        print("Send button detected.")
    except Exception as e:
        logger.error("Send button not found in quick check: " + str(e))
        print("Send button not found in quick check.", e)

    if send_btn:
        try:
            try:
                cust_elem = driver.find_element(By.XPATH, "//div[contains(@class,'deal-customer')]//span[contains(@class,'cust-name')]")
                customer_name = cust_elem.text.strip().title()
                logger.info(f"Retrieved customer name: {customer_name}")
                print("Found customer name:", customer_name)
            except Exception as e:
                logger.error("Error retrieving customer name: " + str(e))
                print("Customer name not found; using generic greeting.", e)
                customer_name = ""
            logger.info("Locating standard text area (visible messaging area)...")
            textarea = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//textarea[contains(@class, 'emoji-input-action-text')]"))
            )
            if not textarea.get_attribute("value").strip():
                template = templates["standard_text"].format(customer_name=customer_name, sender_name=sender_name)
                textarea.send_keys(template)
                logger.info("Standard text template inserted.")
                print("Inserted generic text template into the text area.")
            logger.info("Clicking the Send button for standard text.")
            send_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@analyticsdetect='CustomerActions|Send|Text']"))
            )
            safe_click(driver, send_btn)
            logger.info("Standard text sent successfully.")
            print("✅ Standard text sent successfully.")
        except Exception as e:
            logger.error("Error sending standard text: " + str(e))
            print("❗ Error sending standard text.", e)
            logger.debug(traceback.format_exc())
    else:
        try:
            logger.info("No Send button; checking for opt-in/resend for standard text.")
            opt_in_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@analyticsdetect='CustomerActions|OptIn|Text']"))
            )
            safe_click(driver, opt_in_btn)
            logger.info("Opt-in/resend button clicked; standard text sent.")
            print("✅ Opt-in/resend button clicked. Standard text sent.")
        except Exception as e:
            logger.error("Error: Neither Send nor opt-in/resend available for standard text: " + str(e))
            print("❗ Error: Neither Send nor opt-in/resend button available.", e)
            logger.debug(traceback.format_exc())

def send_text_wrapper():
    logger.info("Starting standard text process.")
    print("\n--- Sending Standard Text ---")
    options = Options()
    options.debugger_address = "127.0.0.1:9222"
    try:
        driver = webdriver.Chrome(options=options)
        logger.info("Attached to Chrome for standard text.")
    except Exception as e:
        logger.error("Error attaching for standard text: " + str(e))
        print("❗ Error attaching for standard texting:", e)
        return
    try:
        current_url = driver.current_url
        logger.info(f"Current URL (text): {current_url}")
        print("Current URL (text):", current_url)
        ensure_customer_claimed()
        send_text_message(driver)
    except Exception as e:
        logger.error("Error in standard text process: " + str(e))
        print("❗ Error in standard text process.", e)
        logger.debug(traceback.format_exc())

# ------------------------------
# Custom Texting Functions (CTRL+ALT+T)
# ------------------------------

def send_custom_text_message(driver):
    """
    Sends a custom text message.
    Workflow:
      1. Navigate to the Text tab.
      2. Abort if opt-out is found.
      3. Check for opt-in/resend; if found, click it and abort further custom text.
      4. Otherwise, repeatedly prompt the user for a valid custom template choice (A, B, or C).
      5. Ask the user for confirmation to send.
      6. If confirmed, retrieve customer name from the deal-customer container,
         clear the text area, insert the chosen custom template (formatted from templates), and click Send.
    """
    try:
        logger.info("Navigating to Text tab for custom text.")
        text_tab = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[@analyticsdetect='CustomerAction|Navigate|Text']"))
        )
        safe_click(driver, text_tab)
        time.sleep(1)
    except Exception as e:
        logger.error("Error navigating to Text tab for custom text: " + str(e))
        print("Couldn't navigate to the Text tab for custom text.", e)
        return
    
    try:
        logger.info("Checking opt-out status for custom text.")
        if driver.find_elements(By.XPATH, "//h4[contains(text(), 'Status: Opted out')]"):
            logger.warning("Customer opted out of texts in custom text flow. Aborting custom text.")
            print("Customer opted out of texts. Aborting custom text.")
            return
    except Exception as e:
        logger.error("Error checking opt-out for custom text: " + str(e))
    
    try:
        logger.info("Checking for opt-in/resend button in custom text flow.")
        opt_in_btn = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.XPATH, "//button[@analyticsdetect='CustomerActions|OptIn|Text']"))
        )
        if opt_in_btn:
            logger.info("Opt-in/resend button detected in custom text flow. Clicking it and aborting custom text process.")
            print("Opt-in/resend button detected. Clicking it and aborting custom text process...")
            safe_click(driver, opt_in_btn)
            logger.info("✅ Pre-determined opt-in text sent in custom text flow.")
            return
    except Exception as e:
        logger.info("No opt-in/resend button found in custom text flow; proceeding with custom text input. Exception: %s", str(e))
    
    try:
        logger.info("Locating text area for custom text (visible messaging area)...")
        textarea = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//textarea[contains(@class, 'emoji-input-action-text')]"))
        )
        textarea.click()
        textarea.clear()
        
        valid_choice = False
        choice = ""
        while not valid_choice:
            print("Choose a custom text template:")
            print("A: General follow-up")
            print("B: Direct inquiry")
            print("C: Urgent reminder")
            choice = input("Enter choice (A/B/C): ").strip().upper()
            if choice in ["A", "B", "C"]:
                valid_choice = True
            else:
                print("Invalid choice. Please enter A, B, or C.")
        
        confirm = input("Do you want to send this custom text? (Y/N): ").strip().lower()
        if confirm != "y":
            print("Custom text sending cancelled by user.")
            logger.info("Custom text sending cancelled by user after confirmation prompt.")
            return
        
        try:
            cust_elem = driver.find_element(By.XPATH, "//div[contains(@class,'deal-customer')]//span[contains(@class,'cust-name')]")
            customer_name = cust_elem.text.strip().title()
            logger.info(f"Customer name retrieved for custom text: {customer_name}")
            print("Found customer name:", customer_name)
        except Exception as e:
            logger.error("Error retrieving customer name for custom text: " + str(e))
            print("Customer name not found; proceeding without personalization.")
            customer_name = ""
        
        if choice == "A":
            chosen_template = templates["custom_text_A"].format(customer_name=customer_name, sender_name=sender_name)
        elif choice == "B":
            chosen_template = templates["custom_text_B"].format(customer_name=customer_name, sender_name=sender_name)
        else:
            chosen_template = templates["custom_text_C"].format(customer_name=customer_name, sender_name=sender_name)
        
        textarea.send_keys(chosen_template)
        logger.info("Custom text template inserted.")
        print("Custom template inserted into the text area.")
        logger.info("Clicking the Send button for custom text.")
        send_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@analyticsdetect='CustomerActions|Send|Text']"))
        )
        safe_click(driver, send_btn)
        logger.info("Custom text sent successfully.")
        print("✅ Custom text sent successfully.")
    except Exception as e:
        logger.error("Error sending custom text: " + str(e))
        print("❗ Error sending custom text.", e)
        logger.debug(traceback.format_exc())

def send_custom_text_wrapper():
    logger.info("Initiating custom text process.")
    print("\n--- Sending Custom Text Message ---")
    options = Options()
    options.debugger_address = "127.0.0.1:9222"
    try:
        driver = webdriver.Chrome(options=options)
        logger.info("Attached to Chrome for custom text.")
    except Exception as e:
        logger.error("Error attaching for custom text: " + str(e))
        print("❗ Error attaching for custom texting:", e)
        return
    try:
        current_url = driver.current_url
        logger.info(f"Current URL (custom text): {current_url}")
        print("Current URL (custom text):", current_url)
        ensure_customer_claimed()
        send_custom_text_message(driver)
    except Exception as e:
        logger.error("Error in custom text process: " + str(e))
        print("❗ Error in custom text process.", e)
        logger.debug(traceback.format_exc())

# ------------------------------
# Email Functions (CTRL+ALT+E)
# ------------------------------

def send_email_message(driver):
    try:
        logger.info("Retrieving customer name for email.")
        cust_elem = driver.find_element(By.XPATH, "//div[contains(@class,'deal-customer')]//span[contains(@class,'cust-name')]")
        customer_name = cust_elem.text.strip().title()
        if not customer_name:
            logger.error("Customer name is empty; aborting email send.")
            print("Customer name is empty; aborting email send.")
            return
        logger.info(f"Retrieved customer name for email: {customer_name}")
    except Exception as e:
        logger.error("Could not retrieve customer name for email; aborting: " + str(e))
        print("Could not retrieve customer name for email; aborting.", e)
        return

    try:
        logger.info("Navigating to Email tab.")
        email_tab = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[@analyticsdetect='CustomerAction|Navigate|Email']"))
        )
        safe_click(driver, email_tab)
        time.sleep(1)
        logger.info("Email tab clicked.")
    except Exception as e:
        logger.error("Error navigating to Email tab: " + str(e))
        print("Error navigating to Email tab:", e)
        return

    try:
        logger.info("Waiting for subject field in email.")
        subject_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Subject']"))
        )
        subject_field.clear()
        subject_field.send_keys(templates["email_subject"].format(customer_name=customer_name, sender_name=sender_name))
        logger.info("Email subject entered.")
        print("Subject entered.")
    except Exception as e:
        logger.error("Error with subject field: " + str(e))
        print("Error with subject field:", e)
        return

    try:
        logger.info("Locating email body field.")
        try:
            iframe = driver.find_element(By.XPATH, "//iframe[contains(@id, '_ifr')]")
            driver.switch_to.frame(iframe)
            body_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//body[@contenteditable='true']"))
            )
            logger.info("Email body found inside iframe.")
            print("Found email body inside iframe.")
        except Exception as e:
            logger.info("No iframe found; locating email body directly. Exception: %s", str(e))
            driver.switch_to.default_content()
            body_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//body[@contenteditable='true']"))
            )
            logger.info("Email body found directly.")
            print("Found email body directly.")
        
        body_field.click()
        body_field.send_keys(Keys.CONTROL + "a")
        body_field.send_keys(Keys.BACKSPACE)
        
        email_message = templates["email_body"].format(customer_name=customer_name, sender_name=sender_name)
        body_field.send_keys(email_message)
        logger.info("Email body entered.")
        print("Email body entered.")
        driver.switch_to.default_content()
    except Exception as e:
        logger.error("Error with email body: " + str(e))
        print("Error with email body:", e)
        return

    try:
        logger.info("Waiting for Send Email button.")
        send_email_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@analyticsdetect='ComposeEmail|Send|Email']"))
        )
        safe_click(driver, send_email_button)
        logger.info("Email sent successfully.")
        print("✅ Email sent successfully!")
    except Exception as e:
        logger.error("Error clicking the Send Email button: " + str(e))
        print("Error clicking the Send Email button:", e)
        logger.debug(traceback.format_exc())

def send_email_wrapper():
    logger.info("Initiating email sending process.")
    print("\n--- Sending Email ---")
    options = Options()
    options.debugger_address = "127.0.0.1:9222"
    try:
        driver = webdriver.Chrome(options=options)
        logger.info("Attached to Chrome for emailing.")
    except Exception as e:
        logger.error("Error attaching to Chrome for emailing: " + str(e))
        print("❗ Error attaching to Chrome for emailing:", e)
        return
    try:
        current_url = driver.current_url
        logger.info(f"Current URL (email): {current_url}")
        print("Current URL (email):", current_url)
        ensure_customer_claimed()
        send_email_message(driver)
    except Exception as e:
        logger.error("Error during email sending process: " + str(e))
        print("❗ Error during email sending process:", e)
        logger.debug(traceback.format_exc())

# ------------------------------
# Manual Template Edit (CTRL+ALT+P)
# ------------------------------

def edit_templates_wrapper():
    manual_edit_templates(save_after=True)

# ------------------------------
# Main and Hotkey Bindings
# ------------------------------

def main():
    global sender_name
    sender_name = login()
    load_templates()
    print_watermark()
    logger.info(f"User '{sender_name}' logged in.")
    launch_chrome()
    print("\nSetup complete. Steps:")
    print("1. In the Chrome window, log into DriveCentric and navigate to the 'Claim Customer' page.")
    print("2. Press F8 to auto-claim and update the task.")
    print("3. Press F9 to send a standard text message.")
    print("4. Press CTRL+ALT+T to send a custom text message.")
    print("   (If an opt-in/resend button is detected, that text is sent immediately.)")
    print("5. Press CTRL+ALT+E to send a generic email message (if customer name is found).")
    print("6. Press CTRL+ALT+P to manually edit the text and email templates.")
    print("7. Press ESC to exit.\n")
    keyboard.add_hotkey("F8", claim_customer)
    keyboard.add_hotkey("F9", send_text_wrapper)
    keyboard.add_hotkey("ctrl+alt+t", send_custom_text_wrapper)
    keyboard.add_hotkey("ctrl+alt+e", send_email_wrapper)
    keyboard.add_hotkey("ctrl+alt+p", edit_templates_wrapper)
    logger.info("Hotkeys registered. Awaiting user input...")
    keyboard.wait("esc")
    logger.info("Exiting program.")
    print("Exiting script.")

if __name__ == "__main__":
    main()

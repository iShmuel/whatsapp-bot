from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
from docx import Document
import gspread
import time
import os
import random




# === הגדרות קבועות ===
CHROMEDRIVER_PATH = os.path.join(os.getcwd(), "chromedriver.exe")
WHATSAPP_URL = "https://web.whatsapp.com"
CREDENTIALS_PATH = "D:\\whatsapp-bot\\credentials.json"
SPREADSHEET_URL = "yoursheet"
# SPREADSHEET_URL = "yoursheet" #האנשים המיועדים 
ATTACHMENT_PATH = "D:\\whatsapp-bot\\Shmuel-Berger.docx"
AUTOIT_HANDLE_OPEN_WINDOW = "D:\whatsapp-bot\AutoIThandlesOpenWindow.exe"

# === הגדרת דפדפן Brave ===
chrome_options = Options()
chrome_options.binary_location = r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"
chrome_options.add_argument("--user-data-dir=D:\\whatsapp-bot\\brave-profile")
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--remote-debugging-port=9222")

service = Service(CHROMEDRIVER_PATH)
browser = webdriver.Chrome(service=service, options=chrome_options)

# === התחברות לגוגל שיט ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_PATH, scope)
client = gspread.authorize(creds)
sheet = client.open_by_url(SPREADSHEET_URL).sheet1

# === פונקציות עזר ===
# entering whatsap
def wait_for_whatsapp():
    print("🔄 Waiting for WhatsApp Web...")
    browser.get(WHATSAPP_URL)
    time.sleep(20)
    try:
        confirm_button = browser.find_element(By.XPATH, '//*[@id="app"]/div/span[2]/div/div/div/div/div/div/div[2]/div/button')
        confirm_button.click()
        print("✅ Clicked confirmation popup.")
        time.sleep(10)
    except NoSuchElementException:
        print("ℹ️ No confirmation popup found.")
# search contact & entering to chat & delete the search input
def search_contact(phone_number):
    try:
        search_box = browser.find_element(By.CSS_SELECTOR, "#side div[contenteditable='true']")
        # search_field = browser.find_element(By.CSS_SELECTOR, "div[role='textbox'][contenteditable='true']") #איזור החיפוש שלאחר ההקלקלה 
        search_box.click()
        time.sleep(1)
        search_box.send_keys(phone_number)
        search_box.send_keys(Keys.ENTER)
        print(f"🔍 Searching for {phone_number}")
        time.sleep(1) # המתן שייכנס לצ'אט
        search_box.send_keys(Keys.CONTROL, "a")  # בחר הכל
        search_box.send_keys(Keys.BACKSPACE)     # מחק
        print("🧹 Cleared search field.")
    except (NoSuchElementException, StaleElementReferenceException) as e:
        print(f"❌ Could not search for {phone_number}: {e}")
# File integrity check
def is_valid_docx(file_path):
    # האם הקובץ קיים
    if not os.path.exists(file_path):
        print("❌ File does not exist.")
        return False

    # האם סיומת הקובץ נכונה
    if not file_path.lower().endswith(".docx"):
        print("❌ File is not a .docx document.")
        return False

    # ניסיון לפתוח את הקובץ עם python-docx
    try:
        Document(file_path)
    except Exception as e:
        print("❌ Failed to open the DOCX file:", str(e))
        return False
    print("✅ DOCX file is valid and accessible.")
    return True

def send_attachment(name):
    try:
        if is_valid_docx(ATTACHMENT_PATH):
            try:
                # כפתור ה +
                plus_button = browser.find_element(By.CSS_SELECTOR, 'div._ak1r button')
                plus_button.click()
                time.sleep(2)
                # כפתור המסמך
                document_option = browser.find_element(By.XPATH, '//*[@id="app"]/div/span[5]/div/ul/div/div/div[1]/li/div')
                document_option.click()
                # הכנסת המסמך -מפה נשתמש ב-AutoIt 
                os.system(AUTOIT_HANDLE_OPEN_WINDOW)
                # ללא שימוש ב-AutoIT אלא ב-Selenium
                # file_input = browser.find_element(By.CSS_SELECTOR, "input[type='file']")
                # file_input.send_keys(ATTACHMENT_PATH)
                # print("📁 Word file attached successfully.")

                # ממתין ששדה ההודעה יהיה נגיש
                WebDriverWait(browser, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div[aria-label="הוספת כיתוב"]'))
                )

                # מוסיף את ההודעה לשדה של הקובץ המצורף
                message_box = browser.find_element(By.CSS_SELECTOR, 'div[aria-label="הוספת כיתוב"]')
                message = f"שלום {name}, מצורף קובץ קורות החיים שלי."
                message_box.send_keys(message) # כדאי לעשות שזה יהיה הקלקה בפעילות מקלדת 
                print(f"💬 Message typed for {name}")

                # שליחה
                WebDriverWait(browser, 15).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "[aria-label='שליחה']"))
                )
                send_button = browser.find_element(By.CSS_SELECTOR, "[aria-label='שליחה']")
                send_button.click()
                print(f"📤 File + message sent to {name}")
                time.sleep(10)

            except NoSuchElementException:
                print("❌ File input not found.")
        else:
            print("⚠️ Word file is invalid or corrupted.")

        time.sleep(5)
    except NoSuchElementException:
        print("❌ File input or send button not found.")
    except StaleElementReferenceException:
        print("🔁 Retrying file attachment...")

# Sending the message attached to files
 
# === הפעלת התהליך ===
try:
    wait_for_whatsapp()
    pause_every = random.randint(8, 10)  # הגרלה אחת בתחילת התוכנית
    messages_sent = 0

    for i, row in enumerate(sheet.get_all_values()[1:]):
        first_name = row[0]
        clean_phone = row[5]

        if clean_phone:
            search_contact(clean_phone)
            send_attachment(first_name)
            messages_sent += 1

            if messages_sent % pause_every == 0:
                sleep_minutes = random.randint(15, 20)
                print(f"⏸️ Pausing for {sleep_minutes} minutes after {messages_sent} messages...")
                time.sleep(sleep_minutes * 60)
                pause_every = random.randint(8, 10)  # הגרלה שוב
                messages_sent=0
                # if i > 0 and i % random.randint(8, 10) == 0:
                #     sleep_minutes = random.randint(10, 20)
                #     print(f"⏸️ מחכה {sleep_minutes} דקות...")
                #     time.sleep(sleep_minutes * 60)
        else:
            print(f"⚠️ מספר טלפון חסר עבור {first_name}, דילוג...")

except Exception as e:
    print("❌ Error during execution:", str(e))

finally:
    print("🏁 Done.")

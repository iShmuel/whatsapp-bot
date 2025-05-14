# WhatsApp Resume Bot

This project is a fully automated solution for sending a personalized message and an attached resume (Word document) to multiple contacts via WhatsApp Web. It leverages Google Sheets as the contact source, Selenium for browser automation, and AutoIt to handle the native file upload window that appears when attaching a document in WhatsApp.

---

## 🔍 Overview

The goal of this bot is to send a message and document to multiple people via WhatsApp with minimal manual intervention. This is useful for job applications, networking, or mass outreach where a personal touch is still maintained by using each contact's first name in the message.

---

## 📋 Features

* Fetches contact data live from a Google Sheet.
* Sends a personalized greeting message with the contact's first name.
* Attaches a Word document to each message.
* Uses Brave Browser for WhatsApp Web automation.
* Waits for user to scan the QR code once.
* Inserts pauses after every 8–10 messages (randomized) to mimic human behavior and avoid spam detection.
* Handles the native file upload dialog using AutoIt.

---

## 📁 File Structure

```
.
├── AutoIThandlesOpenWindow.au3       # AutoIt script to select file in Windows dialog
├── chromedriver.exe                  # ChromeDriver for Selenium
├── credentials.json                  # Google Service Account credentials for Sheets access
├── main.py                           # Main Python bot script
├── requirements.txt                  # Python dependencies list
├── Shmuel-Berger.docx                # Resume to send
└── README.md                         # This file
```

---

## 🧾 Data Source: Google Sheets

The contact list is stored in a Google Sheet:
[Google Sheet](Confidentiality)

### Sheet Columns:

* **First Name**
* **Middle Name**
* **Last Name**
* **Organization**
* **Mobile** (raw number)
* **Clean Phone** (9-digit number for Israeli WhatsApp format)
* **Home**

The script uses the "Clean Phone" column to search and contact the person in WhatsApp.

**Authorization:**
The Google Sheet is shared with this service account:

```
whatsapp-bot@whatsapp-bot-459108.iam.gserviceaccount.com
```

This allows programmatic access to the sheet without user login.

---

## ⚙️ Prerequisites

* Python 3.10+
* Windows 11
* Brave Browser installed
* Google Cloud project with a Service Account and `credentials.json`
* The following Python libraries:

  ```bash
  pip install -r requirements.txt
  ```
* AutoIt + SciTE4AutoIt3 (for scripting file selection)

---

## 🚀 Setup & Execution

1. **Clone the Repository**

   ```bash
   git clone https://github.com/your-username/whatsapp-bot
   cd whatsapp-bot
   ```

2. **Install Python Libraries**

   ```bash
   pip install -r requirements.txt
   ```

3. **Setup Google Service Account**

   * Create a project in [Google Cloud Console](https://console.cloud.google.com/).
   * Enable Google Sheets API.
   * Create a Service Account, download the `credentials.json` file.
   * Share your Google Sheet with the service account email.

4. **Edit `main.py`**

   * Confirm your Google Sheet ID is correct.
   * Ensure the path to `Shmuel-Berger.docx` is correct.

5. **Install AutoIt and compile the script**

   * Download [AutoIt](https://www.autoitscript.com/site/autoit/downloads/).
   * Open `AutoIThandlesOpenWindow.au3` with SciTE4AutoIt3.
   * Run or compile it to `.exe`.

6. **Run the Bot**

   ```bash
   python main.py
   ```

   * The script opens Brave browser and loads WhatsApp Web.
   * Scan the QR code once.
   * Automatically loops through contacts and sends messages.

---

## 🧠 Logic Details

* The `main.py` script fetches fresh rows from Google Sheets every run using:

  ```python
  for i, row in enumerate(sheet.get_all_values()[1:]):
  ```
* It skips any row with a missing phone number.
* After every 8–10 messages, it pauses randomly between 10–20 minutes:

  ```python
  pause_every = random.randint(8, 10)
  if i > 0 and i % pause_every == 0:
      time.sleep(random.randint(10, 20) * 60)
  ```
* It uses fixed XPath selectors to:

  * Click the paperclip icon
  * Select "Document"
  * Use AutoIt to complete file upload when the native Windows "Open" dialog appears

---

## 📞 Message Format

The message sent is:

```text
שלום [First Name], מצורף קובץ קורות החיים שלי.
```

The bot replaces `[First Name]` dynamically.

---

## 💡 Challenges & Solutions

| Challenge                     | Solution                                                                    |
| ----------------------------- | --------------------------------------------------------------------------- |
| Automating file upload dialog | Used AutoIt to detect and interact with native file window                  |
| Avoiding spam detection       | Random pauses and delay between each message                                |
| Dynamic contact list          | Fetch data from Google Sheets live at each run                              |
| Hebrew name formatting        | Used `First Name` directly from sheet for personalization                   |
| File dialog handling          | Created AutoIt script triggered right after clicking "Document" in WhatsApp |

---

## 🎯 Motivation / Background

During my time working at *Aroma Café*, I had the opportunity to connect with many professionals from the tech industry. Over time, I collected dozens (and eventually hundreds) of phone numbers of people I spoke with in person. I realized that instead of manually reaching out to each one, I could build a tool that automates this process — while still keeping it personal.

This project was born from the need to send my resume efficiently, along with a personalized message introducing myself, where we met, what I’m looking for, and what experience I bring. The idea is to scale meaningful outreach using code.

---

## 📃 License

MIT License. You are free to modify, use, and share this project.

---

## 🙋‍♂️ Author

Shmuel Berger
Developed with ❤️ for automation and efficiency.

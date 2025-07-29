Setup Instructions

1️⃣ Install Required Python Packages
pip install selenium twilio requests


2️⃣ Download and Install ChromeDriver
Selenium needs ChromeDriver to automate the Chrome browser.
Download from: https://chromedriver.chromium.org/downloads
Choose the version that matches your Chrome browser
Unzip and move the binary to a known location (e.g., /usr/local/bin/chromedriver on Mac/Linux or C:\chromedriver.exe on Windows)
Make sure the path is updated in your code:
path = '/usr/local/bin/chromedriver'  # update as needed


3️⃣ Set Up a Twilio Account
Sign up at https://www.twilio.com
Navigate to the Console Dashboard
Copy your:
Account SID
Auth Token
Create a Twilio WhatsApp sender number (usually a sandbox number)
Replace these in your code:
TWILIO_ACCOUNT_SID = '<your_twilio_account_sid>'
TWILIO_AUTH_TOKEN = '<your_twilio_auth_token>'
TWILIO_PHONE_NUMBER = 'whatsapp:+14155238886'  # Example Twilio sandbox number
TO_PHONE_NUMBER = 'whatsapp:+91xxxxxxxxxx'    # Your personal WhatsApp number
📨 To receive WhatsApp messages:
Send join <code> to the Twilio number (shown in their console) via WhatsApp
This authorizes Twilio to send messages to your number


4️⃣ Set Your DTU Login Credentials
Replace placeholders in your code:
search.send_keys("<your_rollnumber>")
search_pass.send_keys("<your_password>")


5️⃣ Identify the Course Rows You Want to Monitor
You must provide the row numbers (1-indexed) of the courses you're interested in.
row = [2, 4, 6]  # Example rows for desired courses
🧩 How to Get XPath for Your Course Rows
Open your DTU registration page in Chrome
Right-click the course name or register button → Click Inspect
The highlighted element appears in DevTools
Right-click that element in the Elements panel → Copy → Copy XPath
Optional: Use Chrome Extension like "XPath Helper" to simplify path extraction:
🔗 XPath Helper

Example XPath:
//tbody/tr[4]/td[6]/form/button
Find which tr[n] your desired course is on, then use that index in your row list.


🔁 How the Bot Works
Logs in to DTU student portal
Checks the number of remaining seats (td[5])
If available, finds the registration button (td[6]/form/button)
Scrolls to the element and clicks it
Sends a Twilio WhatsApp notification

Refreshes and continues checking in a loop

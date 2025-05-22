
# 📱 WhatsApp Message Automation

Automate personalized message sending through **WhatsApp Desktop** using Python. This tool reads an Excel file with contacts and URLs, then sends two Hebrew-supported text messages and an image to each contact using clipboard-based automation.

---

## 🚀 Features

- ✅ **Reads contact data from Excel** (`.xlsx`) and tracks processed rows.
- 📝 **Sends personalized text messages** (including Hebrew support with RTL).
- 🖼️ **Sends images** directly using clipboard automation.
- ⚠️ **Skips faulty URLs** and highlights successful deliveries.
- 🧠 **Remembers where it left off** using a configurable state file.
- 🧾 **Logs failed attempts** and reasons via Excel comments.

---

## 📂 Folder Structure

```plaintext
project-root/
│
├── messeges/
│   ├── first_message.txt     # First personalized message (supports Hebrew)
│   ├── second_message.txt    # Second follow-up message
│   └── image.jpg             # Image to be sent
│
├── images/
│   ├── not_found.png         # Screenshot for phone not found detection
│   ├── not_found2.png        # Additional screenshot
│   └── page_404_whatsapp.png # Screenshot for 404 detection
│
├── excel.xlsx                # Excel file with contact info
├── run_state.txt             # Keeps track of current row and file
└── whatsapp_automation.py    # Main script
````

---

## 🛠️ Dependencies

Install the following Python packages before running:

```bash
pip install openpyxl pyautogui pyperclip Pillow pywin32
```

---

## 📈 Excel Format

Ensure your Excel file (`excel.xlsx`) includes the following columns:

| Name       | Phone URL                  | Comments (optional) |
| ---------- | -------------------------- | ------------------- |
| John Doe   | `https://wa.me/1234567890` |                     |
| Jane Smith | `https://wa.me/9876543210` |                     |

> ✅ Successfully processed rows will be **highlighted** and/or **annotated** in the "Comments" column.

---

## 💬 Message Personalization

* First message includes the recipient’s **first name**.
* RTL (`\u200F`) is used to ensure Hebrew text alignment.
* Image and messages are pasted **via clipboard** into WhatsApp Desktop.

---

## 🧪 How It Works

1. Reads current position from `run_state.txt`.
2. Loads message templates and image from `messeges/`.
3. Iterates through rows in Excel:

   * Skips rows already processed or faulty.
   * Opens WhatsApp chat via URL.
   * Sends image and two messages.
   * Highlights or comments results in Excel.
4. Updates `run_state.txt` so it can resume later.

---

## 📌 Tips & Notes

* Run with **WhatsApp Desktop open** and properly configured.
* Make sure image paths and message files are **correct and UTF-8 encoded**.
* If the script encounters:

  * ❌ 404 page → Tab is closed and skipped.
  * ❌ Phone number not found → Message skipped.

---

## 🧑‍💻 Author

**Yuval Kogan**
🔗 [LinkedIn](https://www.linkedin.com/in/yuval-kogan)
💻 [GitHub](https://github.com/KoganTheDev)

---

## 📄 License

This project is licensed under the MIT License. See [`LICENSE`](LICENSE) for more details.


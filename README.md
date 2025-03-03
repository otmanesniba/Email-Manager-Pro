# Email Manager Pro 📧


**Email Manager Pro** is a Python-based tool for managing and organizing your Gmail sent emails. It allows you to search, extract, and delete emails based on subject lines, and saves the extracted email data to a file on your desktop. Built with **IMAP**, **Pyfiglet**, and **Colorama**, this tool provides a user-friendly terminal interface with colorful output.

---

## **Features**
- **Search Sent Emails**: Search emails in your Gmail "Sent Mail" folder by subject.
- **Extract Email Data**: Extract recipient, subject, and date information from matching emails.
- **Save to File**: Save extracted email data to a `.txt` file on your desktop.
- **Delete Emails**: Option to delete matching emails from your "Sent Mail" folder.
- **User-Friendly Interface**: Colorful terminal output with ASCII banners and progress tracking.
- **Secure**: Uses app passwords for authentication (recommended for Gmail accounts with 2FA).

---

## **Prerequisites**
Before running the script, ensure you have the following:
1. **Python 3.8+** installed.
2. Required Python libraries installed. You can install them using:

    ```bash
   pip install pyfiglet colorama

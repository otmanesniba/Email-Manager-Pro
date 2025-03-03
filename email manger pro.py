import imaplib
import email
import re
from email.header import decode_header
import getpass
import os
from datetime import datetime
import pyfiglet
from colorama import Fore, Style, init
import time

# Initialize colorama
init(autoreset=True)

# ASCII Banner
def show_banner():
    f = pyfiglet.Figlet(font='slant')
    print(Fore.CYAN + f.renderText('📧 Email Manager Pro'))
    print(Fore.GREEN + "=" * 70)
    print(Fore.YELLOW + "⚡ Created by Otmane Sniba 7 | Secure Email Management Tool ⚡ 🛡️")
    print(Fore.GREEN + "=" * 70 + Style.RESET_ALL + "\n")

def validate_email(email):
    """Basic email validation"""
    pattern = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return re.match(pattern, email) is not None

def get_valid_email(prompt):
    """Keep asking until a valid email is entered"""
    while True:
        email = input(Fore.CYAN + "📩 " + prompt + Style.RESET_ALL)
        if validate_email(email):
            return email
        print(Fore.RED + "❌ Invalid email format! Please use a valid email address")

def initialize_connection():
    """Handle IMAP connection setup"""
    print(Fore.YELLOW + "\n" + "═"*40 + " 🔌 INITIALIZING CONNECTION 🔌 " + "═"*40)
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(email_user, email_pass)
        print(Fore.GREEN + "✅ Login successful! 🎉" + Style.RESET_ALL)
        return mail
    except Exception as e:
        print(Fore.RED + f"❌ Connection failed: {str(e)} 💥" + Style.RESET_ALL)
        exit()

def save_emails_to_file(email_data):
    """Save found emails to desktop"""
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    filename = f"Found_Emails_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
    file_path = os.path.join(desktop_path, filename)
    
    try:
        with open(file_path, "w", encoding='utf-8') as f:
            f.write("=== Emails Found ===\n\n")
            f.write(email_data)
            f.write("\n\n=== End of List ===")
        print(Fore.GREEN + f"💾 Email list saved to: {file_path} 📂" + Style.RESET_ALL)
    except Exception as e:
        print(Fore.RED + f"❌ Error saving file: {str(e)} 💥" + Style.RESET_ALL)

def delete_emails_from_sent(email_ids, mail):
    """Handle email deletion with progress tracking"""
    print(Fore.YELLOW + "\n" + "═"*40 + " 🗑️ DELETING EMAILS 🗑️ " + "═"*40)
    success = 0
    total = len(email_ids)
    
    for idx, email_id in enumerate(email_ids, 1):
        try:
            mail.store(email_id, "+FLAGS", "\\Deleted")
            print(Fore.BLUE + f"🔥 Deleting email {idx}/{total} ({int((idx/total)*100)}% complete)..." + Style.RESET_ALL)
            success += 1
            time.sleep(1)
        except Exception as e:
            print(Fore.RED + f"❌ Failed to delete email {idx}: {str(e)} 💣")
    
    mail.expunge()
    print(Fore.GREEN + f"\n🎉 Successfully deleted {success}/{total} emails! ✅" + Style.RESET_ALL)

def search_sent_emails(mail):
    """Search and process sent emails"""
    print(Fore.YELLOW + "\n" + "═"*40 + " 🔍 SEARCHING EMAILS 🔍 " + "═"*40)
    subject_start = input(Fore.MAGENTA + "🔎 Enter the starting words of the subject: " + Style.RESET_ALL).strip()
    
    try:
        mail.select('"[Gmail]/Sent Mail"')
        status, messages = mail.search(None, f'SUBJECT "{subject_start}"')
        
        if status != "OK" or not messages[0]:
            print(Fore.RED + "❌ No emails found with that subject. 😞" + Style.RESET_ALL)
            return

        email_ids = messages[0].split()
        print(Fore.GREEN + f"🎯 Found {len(email_ids)} email(s) matching your search! 🎉" + Style.RESET_ALL)
        
        email_data = ""
        processed_emails = []

        for idx, email_id in enumerate(email_ids, 1):
            try:
                status, msg_data = mail.fetch(email_id, "(RFC822)")
                if isinstance(msg_data[0], tuple):
                    msg = email.message_from_bytes(msg_data[0][1])
                    subject = decode_header(msg["Subject"])[0][0]
                    if isinstance(subject, bytes):
                        subject = subject.decode()
                    
                    to_field = msg.get("To", "")
                    date_field = msg.get("Date", "")
                    
                    for recipient in to_field.split(","):
                        clean_recipient = recipient.strip()
                        if clean_recipient and clean_recipient != email_user:
                            email_entry = f"📅 Date: {date_field}\n📨 To: {clean_recipient}\n📝 Subject: {subject}\n\n"
                            email_data += email_entry
                            processed_emails.append(clean_recipient)
                            print(Fore.BLUE + f"📤 Processing {idx}/{len(email_ids)}: Found email to {clean_recipient} 👤" + Style.RESET_ALL)
                
            except Exception as e:
                print(Fore.RED + f"❌ Error processing email {idx}: {str(e)} ⚠️")

        if not processed_emails:
            print(Fore.RED + "❌ No valid recipients found in matching emails. 😕" + Style.RESET_ALL)
            return

        print(Fore.YELLOW + "\n" + "═"*40 + " 🛠️ ACTION MENU 🛠️ " + "═"*40)
        while True:
            choice = input(Fore.MAGENTA + "📝 Choose action:\n1️⃣  Save emails to file\n2️⃣  Delete these emails\n3️⃣  Return to menu\n👉 Enter choice (1-3): " + Style.RESET_ALL)
            
            if choice == "1":
                save_emails_to_file(email_data)
                break
            elif choice == "2":
                delete_emails_from_sent(email_ids, mail)
                break
            elif choice == "3":
                break
            else:
                print(Fore.RED + "❌ Invalid choice. Please try again. 🔄")

    except Exception as e:
        print(Fore.RED + f"❌ Search error: {str(e)} 💥" + Style.RESET_ALL)

def main_loop():
    """Main program loop"""
    show_banner()
    
    # Get credentials
    global email_user, email_pass
    email_user = get_valid_email("Enter your email address: ")
    email_pass = getpass.getpass(Fore.CYAN + "🔑 Enter your password/app password: " + Style.RESET_ALL)
    
    # Initialize connection
    mail = initialize_connection()
    
    while True:
        print(Fore.YELLOW + "\n" + "═"*40 + " 🚀 MAIN MENU 🚀 " + "═"*40)
        print(Fore.BLUE + "1. 🔍 Search Sent emails by subject")
        print(Fore.BLUE + "2. 🚪 Exit program")
        
        choice = input(Fore.MAGENTA + "\n🎯 Enter your choice (1-2): " + Style.RESET_ALL)
        
        if choice == "1":
            search_sent_emails(mail)
        elif choice == "2":
            print(Fore.CYAN + "\n" + "═"*40 + " 🙏 THANK YOU 🙏 " + "═"*40)
            print(Fore.GREEN + "👍 Thank you for using Email Manager Pro! 💌" + Style.RESET_ALL)
            break
        else:
            print(Fore.RED + "❌ Invalid choice. Please try again. 🔄")

    # Cleanup
    try:
        mail.logout()
    except:
        pass

if __name__ == "__main__":
    main_loop()
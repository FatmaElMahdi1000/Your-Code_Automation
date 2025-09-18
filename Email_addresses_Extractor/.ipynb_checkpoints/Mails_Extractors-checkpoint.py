import win32com.client as win32

# --- Domains to ignore (personal emails) ---
IGNORE_DOMAINS = ["gmail.com", "yahoo.com", "yahoo.fr", "outlook.com", "hotmail.com"]

def is_business_email(email):
    """Check if email does NOT belong to ignored domains."""
    if not email:
        return False
    email = email.lower()
    return not any(email.endswith("@" + domain) or domain in email for domain in IGNORE_DOMAINS)

def main():
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Folders: 6 = Inbox, 5 = Sent Items
    folders = {
        "Inbox": outlook.GetDefaultFolder(6),
        "Sent Items": outlook.GetDefaultFolder(5)
    }

    emails = set()  # avoid duplicates

    for folder_name, folder in folders.items():
        print(f"[+] Scanning {folder_name}...")

        messages = folder.Items
        for message in messages:
            try:
                if folder_name == "Inbox":  
                    # From Inbox → Sender
                    sender = message.SenderEmailAddress
                    if is_business_email(sender):
                        emails.add(sender)

                elif folder_name == "Sent Items":
                    # From Sent Items → Recipients
                    for recipient in message.Recipients:
                        try:
                            if recipient.AddressEntry.Type == "EX":
                                addr = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                            else:
                                addr = recipient.Address
                            if is_business_email(addr):
                                emails.add(addr)
                        except Exception:
                            continue
            except Exception:
                continue

    # Save results
    save_path = r"C:\Users\USER\Your-Code_Automation\Emails_Sender_clients\Emails_Addresses_list.txt"
    with open(save_path, "w", encoding="utf-8") as f:
        for email in sorted(emails):
            f.write(email + "\n")

    print(f"\n✅ Saved {len(emails)} unique business email addresses to {save_path}")

if __name__ == "__main__":
    main()


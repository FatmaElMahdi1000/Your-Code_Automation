import win32com.client as win32
from collections import defaultdict

class Email:
    outlook = win32.Dispatch('Outlook.Application')  # class variable
    namespace = outlook.GetNamespace("MAPI")
    Rate_card_link = "https://fatmaelmahdi1000.github.io/Your-Code-_-Pricelist/"
    
    def __init__(self, mails, subject, body, your_subject, domain):
        """
        mails: list of email addresses from the same domain
        subject: subject passed in
        body: body text (HTML supported)
        your_subject: the subject content you provided in Email_Subject.txt
        domain: domain string (e.g. 'onesky.com')
        """
        self.mails = mails
        self.mail_Item = self.outlook.CreateItem(0)
        self.your_subject = your_subject
        self.domain = domain
        
        # Extract client name from domain (e.g., "onesky.com" -> "onesky")
        client_name = domain.split('.')[0]
        updated_client_name = client_name.capitalize()
        
        # Personalize body
        if body and mails:
            Salutes = ["dear", "hello", "hi", "hey"]
            found = False
            for salute in Salutes:
                if salute in body.lower():
                    body = body.replace(
                        salute.capitalize(),
                        f"{salute.capitalize()} {updated_client_name}",
                        1
                    )
                    body = body.replace(
                        salute.lower(),
                        f"{salute.capitalize()} {updated_client_name}",
                        1
                    )
                    found = True
                    break
            if not found:
                # Use <br><br> for HTML blank lines
                body = f"Dear {updated_client_name},<br><br>" + body
        
        # Insert hyperlink if phrase exists
        if "Rate Card Webpage" in body:
            body = body.replace(
                "Rate Card Webpage",
                f'<a href="{self.Rate_card_link}">Rate Card Webpage</a>'
            )
                
        self.body = body

        # Subject (now uses self.your_subject)
        if subject is None:
            self.subject = f"Your Code x {updated_client_name} | {self.your_subject}"
        else:
            self.subject = subject

    def sendingEmail(self):
        drafts_folder = self.namespace.GetDefaultFolder(16)  # Drafts
        # If multiple emails, put all in To:
        if len(self.mails) > 1:
            self.mail_Item.To = "; ".join(self.mails)
        else:
            self.mail_Item.To = self.mails[0]
        self.mail_Item.Subject = self.subject
        self.mail_Item.HTMLBody = self.body   # HTMLBody supports <br>
        self.mail_Item.Move(drafts_folder)    # Save draft
        return f"Draft saved for domain {self.domain}: {', '.join(self.mails)}"


# Standalone function
def read_email_body(body_file):
    with open(body_file, 'r', encoding='utf-8') as f:
        raw_body = f.read()
        html_body = raw_body.replace('\n', '<br>')  # Convert plain text to HTML
    return html_body


# -------------------------
# Main
# -------------------------
file = 'Emails_Addresses_list.txt'
body_file = 'Email_Body.txt'
subject_file = "Email_Subject.txt"

# Read body and subject
body = read_email_body(body_file)
with open(subject_file, 'r', encoding='utf-8') as s:
    your_subject = s.read().strip()

subject = None
with open(file, 'r', encoding='utf-8') as txt:
    Emails_list = txt.read().splitlines()

# Group emails by domain
domain_groups = defaultdict(list)
for mail in Emails_list:
    if '@' in mail:
        domain = mail.split('@')[1]
        domain_groups[domain].append(mail)

# Create drafts per domain group
for domain, mails in domain_groups.items():
    Email_object = Email(mails, subject, body, your_subject, domain)
    print(Email_object.sendingEmail())

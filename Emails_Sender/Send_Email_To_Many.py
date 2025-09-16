import win32com.client as win32
from sys import argv

class Email:
    outlook = win32.Dispatch('Outlook.Application')  # class variable
    namespace = outlook.GetNamespace("MAPI")

    def __init__(self, mail, subject, body):
        self.mail = mail
        self.mail_Item = self.outlook.CreateItem(0)
        idx = mail.index('@')
        name = mail[:idx]

        # Personalize body
        if body and mail:
            Salutes = ["dear", "hello", "hi", "hey"]
            found = False
            for salute in Salutes:
                if salute in body.lower():
                    body = body.replace(
                        salute.capitalize(),
                        f"{salute.capitalize()} {name}",
                        1
                    )
                    body = body.replace(
                        salute.lower(),
                        f"{salute.capitalize()} {name}",
                        1
                    )
                    found = True
                    break
            if not found:
                body = f"Dear {name},\n\n" + body

        self.body = body

        # Subject
        if subject is None:
            if '@' in mail:
                self.subject = f"{name} | Confirming Your Payment Method!"
            else:
                self.subject = "Confirming Your Payment Method!"
        else:
            self.subject = subject

    def sendingEmail(self):
        drafts_folder = self.namespace.GetDefaultFolder(16)  # 16 = Drafts folder
        self.mail_Item.To = self.mail
        self.mail_Item.Subject = self.subject
        self.mail_Item.Body = self.body
        self.mail_Item.Move(drafts_folder)   # Move mail item to Drafts
        return f"Draft saved for {self.mail}"

file = 'Emails.txt'

body = """Hello,

Just a heads-up: To ensure we can process your payment, you'll need a Payoneer account. This is currently the only payment method we support at least at the moment.

Do you have a Payoneer account already? If so, could you please send us the email you used for it? If not, please create an account and then send us the email once you're done.

Thanks for your cooperation!

Best,
Your Code Team
"""

subject = None
with open(file, 'r', encoding='utf-8') as txt:
    Emails_list = txt.read().splitlines()

for address in Emails_list:
    Email_object = Email(address, subject, body)
    print(Email_object.sendingEmail())





        
    
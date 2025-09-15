import win32com.client as win32
from datetime import datetime

# --- Configuration Section ---

RATES = {
    "TRA": "$0.03/word",
    "MTPE": "$0.025/word",
    "PRF": "$15/hour",
    "REV": "$20/hour",
    "QA": "$15/hour",
    "TRANSCRIPTION": "$25/hour"
}

NDA_LINK = "https://docs.google.com/document/d/1te85odMLR4msBOmyS4NIQc378BvjdyWO/edit?usp=sharing&ouid=105338231675038195695&rtpof=true&sd=true"

REPLY_TEMPLATES = {
    "price_request": {
        "body": """
Dear {name},

Thank you for reaching out to us regarding your translation and localization needs. To provide you with an accurate quote, please provide the following details:
1. Source and target languages.
2. The word count of the document.
3. The CAT tool used.
4. The file format (e.g., .docx, .xlsx, .html).
5. Your required deadline.
6. Any specific context, guidelines or reference materials.

Once we have this information, we will prepare a detailed proposal for you. We look forward to the opportunity to work with you.
"""
    },
    "project_update": {
        "body": """
Dear {name},

Thank you for your inquiry about the status of your project.

I am happy to confirm that the translation and localization work for your project is on track and progressing according to the schedule. We will reach out for further updates soon! Please, rest assured! :)
"""
    },
    "vendor_application": {
        "body": """
Dear {name},

Thank you for your interest in joining our team of translators and linguists. We appreciate you taking the time to submit your application.

We receive a high volume of applications, and our team is currently reviewing all submissions. We will contact you if your skills and experience match our current needs.
"""
    },
    "detailed_negotiation": {
        "body": """
Dear {name},

Thank you for your prompt response and for your interest in a partnership. We've carefully reviewed the rates you provided and, in the spirit of a mutually beneficial and long-term collaboration, we'd like to propose the following rates for your consideration:

* Translation (TRA): {trans_rate}
* Machine Translation Post-Editing (MTPE): {mtpe_rate}
* Proofreading (PRF): {proof_rate}
* Revision (REV): {rev_rate}
* Quality Assurance (QA): {qa_rate}
* Transcription: {transcription_rate}

We hope these rates are acceptable for this initial period. Please share with us your language certificates, All information will be kept strictly confidential.

Please also sign this NDA to proceed: {nda_link}
"""
    },
    "Vendor_Negotiation_completion": {
        "body": """
Dear {name},

You’re always welcome dear! Please just fill in this NDA: {nda_link} and share it signed as soon as you can. I've created a profile for you
on LSP.expert (our platform). You may also receive a short, unpaid test in the coming weeks before working on actual projects.

Looking forward to moving forward with our collaboration soon!
"""
    }
}

KEYWORDS_TO_TRIGGER = {
    "price_request": ["quote", "cost", "price", "how much", "estimate"],
    "project_update": ["status update", "project status", "progress", "deadline"],
    "vendor_application": ["vendor application", "join team", "freelance linguist", "translator application", "CV", "rates"],
    "detailed_negotiation": ["Persian", "Farsi", "FR CA", "FR FR", "french", "ES LATAM", "ES MX", "ES MEXICO", "ES SPAIN", "ES AR", "ARGENTINA", "SPANISH",
                             "PT PT", "PT BR", "PORTUGUESE", "PORTUGUESE BRAZIL",
                             "rates", "pricing", "rate proposal", "rate sheet", "price list", "my rates are", "our updated cv", "CV", "Resume"],
    "Vendor_Negotiation_completion": ["very pleased", "agreeing to my proposed translation rate", "agreeing to my rate"]
}

KEYWORDS_TO_EXCLUDE = [
    "Arabic", "Egypt", "Egyptian", "AR", "Arabic Translator", "newsletter", "automatic reply", "Address not found",
    "no-reply", "do not reply", "EN <> [AR]", "EN <> AR", "English ↔ Arabic", "L.E.",
    "The following recipient(s) cannot be reached:", "Server error", "autoresponder", "Delivery Status Notification (Failure)"
]

# --- Script Logic ---

def process_inbox():
    try:
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        unread_items = inbox.Items.Restrict("[Unread] = True")
        unread_items_list = list(unread_items)
        print(f"Checking for unread emails... Found {len(unread_items_list)} unread items.")

        for message in unread_items_list:
            try:
                subject = message.Subject.lower() if message.Subject else ""
                body = message.Body.lower() if message.Body else ""
                sender_name = message.SenderName if message.SenderName else "Client"
                replied = False  

                # --- Exclusion filter ---
                for exclude in KEYWORDS_TO_EXCLUDE:
                    if exclude.lower() in subject or exclude.lower() in body:
                        print(f"Skipping email from {sender_name} (matched exclusion: '{exclude}')")
                        replied = True  # Considered "handled"
                        break
                if replied:
                    continue

                # --- Trigger matching ---
                for response_key, keywords in KEYWORDS_TO_TRIGGER.items():
                    if any(keyword.lower() in subject or keyword.lower() in body for keyword in keywords):
                        template = REPLY_TEMPLATES[response_key]
                        reply_to_email(message, template, sender_name)
                        replied = True
                        break

                if replied:
                    message.Unread = False
                    message.Save()
                    print(f"Draft reply created for email from {message.SenderName} with subject: '{message.Subject}'")

            except Exception as e:
                print(f"Error processing email: {e}")

    except Exception as e:
        print(f"Could not connect to Outlook. Please ensure Outlook is open and running. Error: {e}")

def reply_to_email(original_message, template, recipient_name):
    reply = original_message.Reply()
    reply.Subject = original_message.Subject if original_message.Subject else ""

    new_body = template["body"].format(
        name=recipient_name,
        trans_rate=RATES["TRA"],
        mtpe_rate=RATES["MTPE"],
        proof_rate=RATES["PRF"],
        rev_rate=RATES["REV"],
        qa_rate=RATES["QA"],
        transcription_rate=RATES["TRANSCRIPTION"],
        nda_link=NDA_LINK
    )

    reply.HTMLBody = new_body + "<br><br>" + (original_message.HTMLBody if original_message.HTMLBody else "")
    reply.Save()  # ✅ Save directly to Drafts folder (no window opens)

if __name__ == "__main__":
    print("Starting Outlook email automation script...")
    process_inbox()
    print("Script finished. Check your Drafts folder for replies.")

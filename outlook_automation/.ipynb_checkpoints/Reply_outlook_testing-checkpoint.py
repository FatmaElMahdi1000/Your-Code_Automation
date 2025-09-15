import win32com.client as win32
from datetime import datetime

# --- Configuration Section (unchanged) ---
RATES = {
    "TRA": "$0.03/word",
    "MTPE": "$0.025/word",
    "PRF": "$15/hour",
    "REV": "$20/hour",
    "QA": "$15/hour",
    "TRANSCRIPTION": "$1/Minute"
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

We hope these rates are acceptable for this initial period. Please share with us your language certificates, all information will be kept strictly confidential. Once you confirm the rates, I'll create a profile for you on LSP.expert (our platform). You may also receive a short, unpaid test in the coming weeks or before working on any actual projects.

Please also sign this NDA to proceed: {nda_link}!
"""
    },
    "Vendor_Negotiation_completion": {
        "body": """
Dear {name},

You’re always welcome dear! Please just fill in this NDA: {nda_link} and share it signed as soon as you can. I've created a profile for you on LSP.expert (our platform). You may also receive a short, unpaid test in the coming weeks before working on actual projects.

Looking forward to moving forward with our collaboration soon!
"""
    },
    "Translators offering Services": {
        "body": """
Dear {name},

Thank you for your interest in a partnership, I am impressed by your profile. In the spirit of a mutually beneficial and long-term collaboration, we'd like to propose the following rates for your consideration:

* Translation (TRA): {trans_rate}
* Machine Translation Post-Editing (MTPE): {mtpe_rate}
* Proofreading (PRF): {proof_rate}
* Revision (REV): {rev_rate}
* Quality Assurance (QA): {qa_rate}
* Transcription: {transcription_rate}

We hope these rates are acceptable for this initial period. Please share with us your language certificates, all information will be kept strictly confidential. Once you confirm the rates, I'll create a profile for you on LSP.expert (our platform). You may also receive a short, unpaid test in the coming weeks or before working on any actual projects.

Please also sign this NDA to proceed: {nda_link}!
"""
    }
}

KEYWORDS_TO_TRIGGER = {
    "price_request": ["quote", "cost", "price", "how much", "estimate"],
    "project_update": ["status update", "project status", "progress", "deadline"],
    "vendor_application": ["vendor application", "join team", "freelance linguist", "translator application"],
    "detailed_negotiation": ["persian", "farsi", "fr ca", "fr fr", "french", "es latam", "es mx", "es mexico", "es spain", "es ar", "argentina", "spanish",
                             "pt pt", "pt br", "portuguese", "portuguese brazil",
                             "rates", "pricing", "rate proposal", "rate sheet", "price list", "my rates are", "our updated cv", "cv", "resume"],
    "Vendor_Negotiation_completion": ["very pleased", "agreeing to my proposed translation rate", "agreeing to my rate"],
    "Translators offering Services": ["translator", "i would like to work with you", "you can count on my availability"]
}

KEYWORDS_TO_EXCLUDE = [
    "arabic", "egypt", "egyptian", "arabic translator", "newsletter", "automatic reply", "address not found",
    "no-reply", "do not reply", "en <> ar", "english ↔ arabic", "l.e.",
    "the following recipient(s) cannot be reached:", "server error", "autoresponder", "delivery status notification (failure)"
]

# >>> NEW: helper that robustly finds unread messages across stores (accounts)
def collect_unread_messages(namespace, max_scan_recent=500, debug=False):
    """
    Collect unread items from all stores' Inbox folders.
    - Uses Restrict("[UnRead] = True") first (fast).
    - Falls back to scanning recent items (last max_scan_recent) if Restrict yields zero.
    Returns a list of MailItem objects.
    """
    unread_messages = []

    stores = namespace.Stores
    store_count = stores.Count
    if debug:
        print(f"[debug] Found {store_count} store(s) in profile.")

    for si in range(1, store_count + 1):  # COM collections are 1-based
        try:
            store = stores.Item(si)
            store_display = store.DisplayName if hasattr(store, "DisplayName") else f"Store #{si}"
            inbox = store.GetDefaultFolder(6)  # 6 == olFolderInbox
        except Exception as e:
            if debug:
                print(f"[debug] Could not access store/index {si}: {e}")
            continue

        # Try Restrict first (fast)
        try:
            items = inbox.Items
            # recommended: sort before restrict for reliability
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception:
                pass

            restricted = items.Restrict("[UnRead] = True")
            restricted_count = restricted.Count if hasattr(restricted, "Count") else 0
        except Exception as e:
            restricted = None
            restricted_count = 0
            if debug:
                print(f"[debug] Restrict failed on store '{store_display}': {e}")

        if restricted_count and restricted is not None:
            if debug:
                print(f"[debug] Store '{store_display}' - unread via Restrict: {restricted_count}")
            # extract items from restricted collection
            for i in range(1, restricted_count + 1):
                try:
                    unread_messages.append(restricted.Item(i))
                except Exception:
                    continue
        else:
            # >>> Fallback: scan the most recent items (safe, avoids iterating entire huge mailbox)
            try:
                total_items = items.Count
            except Exception:
                total_items = 0

            if debug:
                print(f"[debug] Store '{store_display}' - Restrict found 0. Scanning last {max_scan_recent} items (total_items={total_items}).")

            # scan backwards from newest to older, up to max_scan_recent
            start_index = max(1, total_items - max_scan_recent + 1)
            for idx in range(total_items, start_index - 1, -1):
                try:
                    it = items.Item(idx)
                    # Only consider mail items that have UnRead property set
                    if getattr(it, "UnRead", False):
                        unread_messages.append(it)
                except Exception:
                    # some items may throw or be non-mail items; ignore and continue
                    continue

    return unread_messages

# --- main processing function (uses the robust unread collector) ---
def process_inbox(debug=False):
    try:
        outlook = win32.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        # >>> Collect unread messages across all stores (robust)
        unread_items_list = collect_unread_messages(namespace, max_scan_recent=800, debug=debug)
        print(f"Checking for unread emails... Found {len(unread_items_list)} unread item(s) across all accounts.")

        # Optional debug: list the subjects we found
        if debug and unread_items_list:
            print("[debug] Subjects of unread items found:")
            for i, m in enumerate(unread_items_list, start=1):
                subj = m.Subject if hasattr(m, "Subject") else "<no subject>"
                print(f"  {i}. {subj}")

        for message in unread_items_list:
            try:
                subject = message.Subject.lower() if message.Subject else ""
                body = message.Body.lower() if message.Body else ""
                sender_name = message.SenderName if message.SenderName else "Client"
                replied = False

                # Exclusion filter
                for exclude in KEYWORDS_TO_EXCLUDE:
                    if exclude.lower() in subject or exclude.lower() in body:
                        print(f"Skipping email from {sender_name} (matched exclusion: '{exclude}')")
                        replied = True
                        break
                if replied:
                    continue

                # Trigger matching (normalized to lowercase)
                for response_key, keywords in KEYWORDS_TO_TRIGGER.items():
                    for kw in keywords:
                        if kw.lower() in subject or kw.lower() in body:
                            template = REPLY_TEMPLATES.get(response_key)
                            if template:
                                reply_to_email(message, template, sender_name)
                                replied = True
                            break
                    if replied:
                        break

                if replied:
                    # Mark as read and save the MailItem (we already saved the reply as draft inside reply_to_email)
                    try:
                        message.UnRead = False
                        message.Save()
                    except Exception:
                        pass
                    print(f"Draft reply created for email from '{sender_name}' with subject: '{message.Subject}'")

            except Exception as e:
                print(f"Error processing message: {e}")

    except Exception as e:
        print(f"Could not connect to Outlook. Please ensure Outlook is open and running. Error: {e}")

# reply_to_email keeps the signature and saves to Drafts
def reply_to_email(original_message, template, recipient_name):
    try:
        reply = original_message.Reply()
        reply.Subject = original_message.Subject or ""

        # ✅ Convert NDA link to HTML hyperlink
        clickable_nda = f'<a href="{NDA_LINK}" target="_blank">NDA</a>'

        # Prepare body text with hyperlink
        new_body_text = template["body"].format(
            name=recipient_name,
            trans_rate=RATES["TRA"],
            mtpe_rate=RATES["MTPE"],
            proof_rate=RATES["PRF"],
            rev_rate=RATES["REV"],
            qa_rate=RATES["QA"],
            transcription_rate=RATES["TRANSCRIPTION"],
            nda_link=clickable_nda  # ✅ replaces plain text with clickable link
        )

        new_body_text = new_body_text.strip()
        signature_html = reply.HTMLBody.strip() if reply.HTMLBody else ""

        final_html = f"""
<html>
  <body style="font-family:Calibri, sans-serif; font-size:11pt;">
    <p>{new_body_text.replace(chr(10), "<br>")}</p>
    {signature_html}
    <br>
    {original_message.HTMLBody or ""}
  </body>
</html>
"""
        reply.HTMLBody = final_html
        reply.Save()
    except Exception as e:
        print(f"Error creating draft reply: {e}")


if __name__ == "__main__":
    print("Starting Outlook email automation script...")
    # Set debug=True for detailed output while troubleshooting
    process_inbox(debug=True)
    print("Script finished. Check your Drafts folder for replies.")


import base_page
from base_page import OutlookAccount


class OrderBooksForm:
    account_name = "netanel.amlinsky@jewsforjesus.org"
    outlook_account = OutlookAccount(account_name)
    if outlook_account.login():
        # set up the sender name to filter by
        sender_name = "Netanel Amlinsky"
        sender_email = 'wordpress@yeshuanekuda.co.il'
        count = 0
        # loop through the emails in the inbox folder
        for message in outlook_account.inbox_folder.Items:  # Books orders form
            # check if the message is from the specified sender name
            if message.Class == 43:
                if message.SenderEmailType == "EX":
                    if message.Sender.GetExchangeUser().PrimarySmtpAddress == sender_email and ("התקבלה בקשה להזמנת" in message.Subject.lower()):
                        # get the email content
                        parts = base_page.OutlookAccount.get_email_content(message)
                        count += 1
                else:
                    if message.SenderEmailAddress == sender_email and ("התקבלה בקשה להזמנת" in message.Subject.lower()):
                        # get the email content
                        parts = base_page.OutlookAccount.get_email_content(message)
                        count += 1

        print("Found {} emails from '{}'".format(count, sender_email))


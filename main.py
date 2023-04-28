import base_page
from base_page import OutlookAccount

account_name = "netanel.amlinsky@jewsforjesus.org"
outlook_account = OutlookAccount(account_name)
if outlook_account.login():
    # set up the sender name to filter by
    sender_name = "Netanel Amlinsky"
    sender_email = 'wordpress@yeshuanekuda.co.il'

    count = 0
    id_order = ""
    full_name = ""
    books = ""
    # loop through the emails in the inbox folder
    for message in outlook_account.inbox_folder.Items:
        # check if the message is from the specified sender name
        if message.Class == 43:
            if message.SenderEmailType == "EX":
                if message.Sender.GetExchangeUser().PrimarySmtpAddress == sender_email:
                    # get the email content
                    parts = base_page.OutlookAccount.get_email_content(message)
                    count += 1
                    print(parts)
            else:
                if message.SenderEmailAddress == sender_email:
                    # get the email content
                    parts = base_page.OutlookAccount.get_email_content(message)
                    count += 1
                    print(parts)

    print("Found {} emails from '{}'".format(count, sender_email))

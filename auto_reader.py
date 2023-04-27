import win32com.client
import re
import openpyxl

# Connect to Outlook and specify the email account name
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
account_name = "netanel.amlinsky@jewsforjesus.org"
account = None

# Log on to the specified account
for a in namespace.Accounts:
    if a.DisplayName == account_name:
        account = a
        break
# if account:
inbox_folder = account.DeliveryStore.GetDefaultFolder(6)
# else:
#     print("Could not find account '{}'".format(account_name))

# Print the number of items in the inbox folder
# print("Number of items in Inbox: {}".format(inbox_folder.Items.Count))

# set up the regular expression to match the shipping information
shipping_regex = r"מס' הזמנה: (\d+)\nשם מלא: (.+)\nכתובת: (.+)\nאימייל: (.+)\nמס' ליצירת קשר: (.+)\nסימן ליצור קשר טלפוני: (.+)\nהספרים שנבחרו: (.+)\n\nIP: (\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})"

# set up the workbook and worksheet for storing the data
wb = openpyxl.Workbook()
ws = wb.active
ws.append(["ID Order", "Full name", "Full Address", "Email", "Phone number", "Call me", "The product", "IP address"])

# set up the sender name to filter by
sender_name = "Netanel Amlinsky"
sender_email = 'netanel.amlinsky@jewsforjesus.org'

count = 0
# loop through the emails in the inbox folder
for message in inbox_folder.Items:
    # print(type(message))
    # print(dir(message))

    # check if the message is from the specified sender name
    if message.Class == 43:
        if message.SenderEmailType == "EX":
            if message.Sender.GetExchangeUser().PrimarySmtpAddress == "wordpress@yeshuanekuda.co.il":
                print("Found an email from 'wordpress@yeshuanekuda.co.il'")
                count += 1
            else:
                if message.SenderEmailAddress == "wordpress@yeshuanekuda.co.il":
                    print("Found an email from 'wordpress@yeshuanekuda.co.il'")
                    count += 1
print(count)
#     if message.SenderEmailAddress == "wordpress@yeshuanekuda.co.il":
#         # get the email content
#         email_body = message.body
#
#         # try to match the shipping information regex
#         shipping_info_match = re.search(shipping_regex, email_body)
#
#         # if the regex match was successful, extract the data and add it to the worksheet
#         if shipping_info_match:
#             id_order = shipping_info_match.group(1)
#             full_name = shipping_info_match.group(2)
#             full_address = shipping_info_match.group(3)
#             email = shipping_info_match.group(4)
#             phone_number = shipping_info_match.group(5)
#             call_me = shipping_info_match.group(6)
#             product = shipping_info_match.group(7)
#             ip_address = shipping_info_match.group(8)
#
#             ws.append([id_order, full_name, full_address, email, phone_number, call_me, product, ip_address])
# #
# # save and close the workbook
# wb.save('C:\\Users\\natan\\Desktop\\EmailAutomation\\shipping_info.xlsx')
# wb.close()
#
# # Load the workbook
# wb = openpyxl.load_workbook('C:\\Users\\natan\\Desktop\\EmailAutomation\\shipping_info.xlsx')
#
# # select a worksheet
# ws = wb.active
#
# # print the contents of the worksheet
# for row in ws.values:
#     print(row)
#
# # close the workbook
# wb.close()

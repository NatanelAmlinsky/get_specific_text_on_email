import win32com.client


class OutlookAccount:
    def __init__(self, account_name):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.account_name = account_name
        self.account = None

    def login(self):
        for a in self.namespace.Accounts:
            if a.DisplayName == self.account_name:
                self.account = a
                break

        if not self.account:
            print(f"Could not find account '{self.account_name}'")
            return False

        self.inbox_folder = self.account.DeliveryStore.GetDefaultFolder(6)
        return True

    def get_email_content(message):
        email_body = message.body
        parts = []

        for line in email_body.splitlines():
            parts.extend(line.split("\n"))


        return parts

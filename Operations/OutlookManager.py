import win32com.client
import logging
from io import BytesIO
import pandas as pd
from Misc import ColoredFormatter

class OutlookManager:
    """
    The OutlookManager class provides an interface for automating various tasks within Microsoft Outlook, such as sending emails, managing tasks, listing emails, and creating calendar events.
    """
    def __init__(self, b_enable_logging):

        # Create a logger
        self.logger = logging.getLogger(self.__class__.__name__)
        if b_enable_logging:
            self.logger.setLevel(logging.DEBUG)
            handler = logging.StreamHandler()
            formatter = ColoredFormatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)
        else:
            self.logger.setLevel(logging.ERROR)

        self.logger.info("Initializing OutlookManager class")

        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
        except Exception as e:
            print(f"Error initializing Outlook: {e}")

    def send_email(self, to, subject, body, cc=None, bcc=None, attachments=None):

        EmailSent = False
        try:
            mail = self.outlook.CreateItem(0)  # 0 represents a MailItem
            mail.To = to
            mail.Subject = subject
            mail.Body = body

            if cc:
                mail.CC = cc
            if bcc:
                mail.BCC = bcc
            if attachments:
                for attachment in attachments:
                    mail.Attachments.Add(attachment)

            mail.Send()
            print("Email sent successfully!")
        except Exception as e:
            print(f"Error sending email: {e}")

        EmailSent = True
        return EmailSent

    def send_email_with_html(self, to, subject, html_body, cc=None, bcc=None, attachments=None):

        HTLMEmailSent = False
        try:
            mail = self.outlook.CreateItem(0)  # 0 represents a MailItem
            mail.To = to
            mail.Subject = subject
            mail.HTMLBody = html_body

            if cc:
                mail.CC = cc
            if bcc:
                mail.BCC = bcc

            if attachments:
                for attachment in attachments:
                    mail.Attachments.Add(attachment)

            mail.Display(True)  # Show email for review before sending
            print("Email prepared successfully. Please review and send manually if needed.")
        except Exception as e:
            print(f"Error preparing email: {e}")

        HTLMEmailSent = True
        return HTLMEmailSent

    def create_task(self, subject, due_date, body=None):

        TasksCreated = False
        try:
            task = self.outlook.CreateItem(3)  # 3 represents a TaskItem
            task.Subject = subject
            task.DueDate = due_date
            if body:
                task.Body = body

            task.Save()
            print("Task created successfully!")
        except Exception as e:
            print(f"Error creating task: {e}")

        TasksCreated = True
        return TasksCreated

    def list_emails(self, folder_name="Inbox", count=10):

        EmailsListed = False
        try:
            folder = self.namespace.GetDefaultFolder(6)  # 6 represents the Inbox folder
            if folder_name.lower() != "inbox":
                for f in self.namespace.Folders:
                    if f.Name.lower() == folder_name.lower():
                        folder = f
                        break

            messages = folder.Items
            messages.Sort("[ReceivedTime]", True)  # Sort by ReceivedTime in descending order

            email_list = []
            for i, message in enumerate(messages, start=1):
                if i > count:
                    break
                email_list.append({
                    "Subject": message.Subject,
                    "Sender": message.SenderName,
                    "ReceivedTime": message.ReceivedTime,
                })
            return email_list

        except Exception as e:
            print(f"Error listing emails: {e}")

        EmailsListed = True
        return EmailsListed


    def create_calendar_event(self, subject, start_time, end_time, location=None, body=None):

        EventCalendarCreated = False
        try:
            appointment = self.outlook.CreateItem(1)  # 1 represents an AppointmentItem
            appointment.Subject = subject
            appointment.Start = start_time
            appointment.End = end_time

            if location:
                appointment.Location = location
            if body:
                appointment.Body = body

            appointment.Save()
            print("Calendar event created successfully!")
        except Exception as e:
            print(f"Error creating calendar event: {e}")

        EventCalendarCreated = True
        return EventCalendarCreated


    def read_latest_attachment_as_dataframe(self, parent_folder_name, subfolder_name, file_type="csv", header_row = None):
        """
        Reads the latest email attachment from a given subfolder directly into a DataFrame.
        :param parent_folder_name: The name of the main folder (e.g., "Inbox")
        :param subfolder_name: The name of the subfolder (e.g., "Oil Brokerage Curves")
        :param file_type: File type to look for ('csv' or 'xlsx')
        :return: pandas DataFrame or None
        """

        LatestAttachmentsReadToDF = False
        try:
            account_folder = self.namespace.Folders.Item(1)
            parent_folder = account_folder.Folders[parent_folder_name]
            target_folder = parent_folder.Folders[subfolder_name]

            messages = target_folder.Items
            messages.Sort("[ReceivedTime]", True)

            for message in messages:
                if message.Class == 43 and message.Attachments.Count > 0:
                    for attachment in message.Attachments:
                        if attachment.FileName.lower().endswith(file_type):
                            attachment_data = attachment.PropertyAccessor.GetProperty(
                                "http://schemas.microsoft.com/mapi/proptag/0x37010102")  # PR_ATTACH_DATA_BIN
                            file_bytes = BytesIO(attachment_data)

                            if file_type == "csv":
                                return pd.read_csv(file_bytes, header = header_row)
                            elif file_type == "xlsx":
                                return pd.read_excel(file_bytes, header = header_row)
            print("No matching attachment found.")
            return None
        except Exception as e:
            print(f"Error reading attachment into DataFrame: {e}")
            return None

        AttachmentsReadToDF = True
        return AttachmentsReadToDF

    def read_attachment_by_subject(self, parent_folder_name, subfolder_name, subject_keyword, file_type="csv",
                                   header_row=None):
        """
        Reads an email attachment based on a subject match into a DataFrame.

        :param parent_folder_name: The name of the main folder (e.g., "Inbox")
        :param subfolder_name: The name of the subfolder (e.g., "Oil Brokerage Curves")
        :param subject_keyword: Keyword or exact subject string to match
        :param file_type: File type to look for ('csv' or 'xlsx')
        :param header_row: Row index to use as header (optional)
        :return: pandas DataFrame or None
        """

        AttachmentBySubjectReadToDF = False
        try:
            account_folder = self.namespace.Folders.Item(1)
            parent_folder = account_folder.Folders[parent_folder_name]
            target_folder = parent_folder.Folders[subfolder_name]

            messages = target_folder.Items
            messages.Sort("[ReceivedTime]", True)

            for message in messages:
                if message.Class == 43 and subject_keyword.lower() in message.Subject.lower():
                    if message.Attachments.Count > 0:
                        for attachment in message.Attachments:
                            if attachment.FileName.lower().endswith(file_type):
                                attachment_data = attachment.PropertyAccessor.GetProperty(
                                    "http://schemas.microsoft.com/mapi/proptag/0x37010102")
                                file_bytes = BytesIO(attachment_data)

                                if file_type == "csv":
                                    return pd.read_csv(file_bytes, header=header_row)
                                elif file_type == "xlsx":
                                    return pd.read_excel(file_bytes, header=header_row)
            print("No matching email with specified subject or attachment found.")
            return None

        except Exception as e:
            print(f"Error reading attachment by subject: {e}")
            return None

        AttachmentBySubjectReadToDF = True
        return AttachmentBySubjectReadToDF

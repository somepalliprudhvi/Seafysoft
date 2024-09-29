# # # import imaplib  # For connecting to the email server
# # # import smtplib  # For sending emails via SMTP
# # # import email  # For parsing email messages
# # # import time  # For sleep functionality
# # # import re  # For regular expressions
# # # import os  # For operating system functionalities
# # # import pandas as pd  # For handling data in tabular format (Excel)
# # # from email.mime.text import MIMEText  # For creating MIME text emails
# # # from datetime import datetime, timedelta  # For handling dates and times

# # # # Configuration for email server and credentials
# # # imap_url = 'imap.gmail.com'  # URL for the IMAP server
# # # # username = 'sandhyacareerr@gmail.com'  # Email address for login
# # # # password = 'wixngueilqtdhpro'  # Password for login
# # # username = 'dailyrequriments@gmail.com'  # Email address for login
# # # password = 'yulyeentyuykfpav'  # Password for login
# # # filePath = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\old_process\SendEmailJobInput.xlsx'


# # # # List of keywords that will cause emails to be skipped (e.g., job titles or statuses to avoid)
# # # RejectedKeywords = ['w2', 'architect', 'fulltime', 'dice', 'gc', 'citizen', 'full-time', 'lead', 'gc-ead', 'usc', 'manager', 'f2f', 'onsite', 'on-site', 'hotlist']

# # # # Setting the date range for fetching emails: emails from the last 16 hours
# # # today = datetime.today()  # Get the current date and time
# # # since_date = (today - timedelta(hours=16)).strftime("%d-%b-%Y")  # Format date for IMAP query

# # # # Load Excel file into DataFrames for later use
# # # df_sheet1 = pd.read_excel(filePath, sheet_name='Sheet1')  # Load job-related data
# # # df_sheet2 = pd.read_excel(filePath, sheet_name='Sheet2')  # Load keywords related to job filtering

# # # # Initialize a DataFrame to store email addresses collected during processing
# # # myDataTable = pd.DataFrame(columns=['To Email'])  # Columns to hold the "To Email" information

# # # # Start an infinite loop to continuously check for new emails
# # # while True:
# # #     try:
# # #         # Connect to the email server using IMAP
# # #         mail = imaplib.IMAP4_SSL(imap_url)  # Establish a secure connection
# # #         mail.login(username, password)  # Login to the email account
# # #         mail.select("inbox")  # Select the inbox folder for processing
# # #         status, response = mail.search(None, '(UNSEEN)', f'SINCE "{since_date}"')  # Search for unseen emails since the specified date

# # #         # Process each email found in the search results
# # #         for num in response[0].split():  # Iterate over each email ID
# # #             time.sleep(1)  # Rate-limiting to avoid being blocked by the server
# # #             mail.store(num, '+FLAGS', '\\Seen')  # Mark the email as seen (read)
# # #             status, data = mail.fetch(num, "(RFC822)")  # Fetch the full email content in RFC822 format

# # #             # Check if fetching the email was successful
# # #             if status != 'OK':
# # #                 print(f"Failed to fetch message {num}. Status code: {status}")
# # #                 continue  # Skip to the next email if the fetch failed
            
# # #             # Parse the email message into a structured format
# # #             email_message = email.message_from_bytes(data[0][1]) if isinstance(data[0][1], bytes) else email.message_from_string(data[0][1])
# # #             email_subject = email_message['subject'].lower()  # Get the email subject and convert to lowercase
            
# # #             # Check if the email subject contains any rejected keywords
# # #             if re.search(r'\b(?:{})\b'.format('|'.join(map(re.escape, RejectedKeywords))), email_subject):
# # #                 continue  # Skip this email if rejected keywords are found

# # #             # Prepare a list to hold email addresses extracted from the message
# # #             email_list = [email_message['Reply-To'], email_message['From']]  # Initialize with Reply-To and From addresses
# # #             payload = email_message.get_payload()  # Get the email body (payload)

# # #             # Extract the body from the email payload
# # #             if isinstance(payload, list):  # If the payload is a list (multipart email)
# # #                 body = next(part.get_payload(decode=True).decode(part.get_content_charset()) for part in payload if part.get_content_type() == "text/plain")
# # #             else:  # If the payload is a single string
# # #                 body = payload if isinstance(payload, str) else payload.decode()

# # #             # Use regex to find email addresses in the body text
# # #             email_addresses = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', body)  # Find all email addresses in the body
# # #             email_list.extend(email_addresses)  # Add any found email addresses to the list

# # #             # Filter and get the first valid email address
# # #             toEmail = next((item for item in email_list if item and "phmailadmin" not in item and ".email" not in item), None)

# # #             # If a valid email address is found and it's not already in the DataFrame
# # #             if toEmail and toEmail not in myDataTable['To Email'].values:
# # #                 #myDataTable = myDataTable.append({'To Email': toEmail}, ignore_index=True)  # Add toDataFrame
# # #                 # Create a new DataFrame with the new email address
# # #                 new_entry = pd.DataFrame({'To Email': [toEmail]})

# # #                 # Concatenate the new entry with the existing DataFrame
# # #                 myDataTable = pd.concat([myDataTable, new_entry], ignore_index=True)

# # #                 # Find matching rows in df_sheet2 based on keywords
# # #                 matching_rows = []  # Initialize a list to hold matching rows
# # #                 for index, row in df_sheet2.iterrows():  # Iterate over each row in the keywords DataFrame
# # #                     # Check if the 'Keywords' cell is a string and contains any keywords that match the email subject
# # #                     if isinstance(row['Keywords'], str) and any(keyword.strip().lower() in email_subject for keyword in row['Keywords'].split(',')):
# # #                         # If matched, append relevant data from df_sheet1
# # #                         matching_rows += df_sheet1[df_sheet1['Technology'] == row['Technology']].values.tolist()

# # #                 # Process each matching row found
# # #                 for row in matching_rows:
# # # #                    employeeUsername, employeePassword, employeeName, employeeStatus, employeeEmailBody = row  # Unpack relevant fields
# # #                     # Clean the email body by removing unwanted signatures or text
                    
# # #                     if len(row) >= 7:
# # #                         # employeeUsername, employeePassword, employeeName, employeeStatus, employeeEmailBody = row[:5]
# # #                         empTech,employeeName,empTech,employeeStatus, employeeUsername, employeePassword, employeeEmailBody = row[:7]


# # #                         body = employeeEmailBody.replace("ThanksRemove/unsubscribe  |  Update your contact and subscribed mailing list(s)  |  Subscribe to mailing list(s) to receive requirements & resumes", "").strip()
                        
# # #                         # Function to remove hyperlinks from the email body
# # #                         def remove_hyperlinks(text):
# # #                             """Remove hyperlinks from the provided text."""
# # #                             return re.sub(r'<(?:https?://)?[^|>]+>', '', text).strip()  # Regex to find and remove hyperlinks

# # #                         body = remove_hyperlinks(body)  # Clean the body of hyperlinks
# # #                         # Regular expression to find phone numbers in the body
# # #                         phone_numbers = re.findall(r"(?:\(?\d{3}\)?-? *\d{3}-? *-?\d{4}|\b\d{4}-\d{6}\b|\b\d{4}-\d{3}-\d{3,4}\b|\b\d{3}[-\s]?\d{3}[-\s]?\d{4}\b)", body)

# # #                         # If phone numbers are found, send a self-reply
# # #                         if phone_numbers:  # Check if any phone numbers were extracted
# # #                             selfReply = MIMEText(body, 'plain')  # Create a reply email with the body
# # #                             selfReply["Subject"] = email_subject + " - call vendor"  # Set subject for self-reply
# # #                             selfReply["To"] = employeeUsername  # Set recipient as employee's email
# # #                             selfReply["Importance"] = "High"  # Mark email as important
                            
# # #                             # Initialize SMTP server for sending emails
# # #                             smtp_server = smtplib.SMTP('smtp.gmail.com', 587)  # Connect to the SMTP server
# # #                             smtp_server.starttls()  # Start TLS for secure communication
# # #                             smtp_server.login(username, password)  # Login to the email account
# # #                             smtp_server.sendmail(username, employeeUsername, selfReply.as_string())  # Send the self-reply
# # #                             smtp_server.close()  # Close the SMTP server connection

# # #                         # Prepare and send a reply to the original email
# # #                         reply = MIMEText(body, 'plain')  # Create the reply email
# # #                         reply["Subject"] = email_subject  # Set the subject of the reply
# # #                         reply["To"] = toEmail  # Set the recipient of the reply
# # #                         smtp_server = smtplib.SMTP('smtp.gmail.com', 587)  # Connect to SMTP server
# # #                         smtp_server.starttls()  # Start TLS
# # #                         smtp_server.login(employeeUsername, employeePassword)  # Login with employee's credentials
# # #                         smtp_server.sendmail(employeeUsername, toEmail, reply.as_string())  # Send the reply email
# # #                         mail.store(num, "+FLAGS", "\\Seen")  # Mark the email as seen
# # #                         mail.store(num, '+FLAGS', '\\Deleted')  # Mark the email for deletion
# # #                         mail.expunge()  # Permanently remove deleted emails from the server
# # #                         print(f"Sent reply on behalf of: {employeeName} to vendor email: {toEmail}")  # Log the action

# # #         # Save collected email addresses to an Excel file after processing all emails
# # #         if not myDataTable.empty:  # Check if there are any collected email addresses
# # #             myDataTable.to_excel('myExcelFile.xlsx', index=False)  # Save to an Excel file without the index

# # #     except imaplib.IMAP4.abort:  # Catch any IMAP-related errors
# # #         print("Error fetching email. Retrying in 10 seconds.")  # Log the error
# # #         time.sleep(10)  # Wait before retrying to avoid rapid consecutive errors

# # #     print("Will wait 500 sec and restart job")  # Log the waiting period
# # #     time.sleep(10)  # Wait before the next iteration of email fetching

# # # # Cleanup on exit
# # # mail.close()  # Close the mailbox connection
# # # mail.logout()  # Logout from the email server



# # import imaplib  # For connecting to the email server
# # import smtplib  # For sending emails via SMTP
# # import email  # For parsing email messages
# # import time  # For sleep functionality
# # import re  # For regular expressions
# # import pandas as pd  # For handling data in tabular format (Excel)
# # from email.mime.text import MIMEText  # For creating MIME text emails
# # from datetime import datetime, timedelta  # For handling dates and times

# # # Configuration for email server and credentials
# # imap_url = 'imap.gmail.com'  # URL for the IMAP server
# # username = 'dailyrequriments@gmail.com'  # Email address for login
# # password = 'yulyeentyuykfpav'  # Password for login
# # filePath = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\old_process\SendEmailJobInput.xlsx'

# # # List of keywords that will cause emails to be skipped
# # RejectedKeywords = ['w2', 'architect', 'fulltime', 'dice', 'gc', 'citizen', 'full-time', 'lead', 'gc-ead', 'usc', 'manager', 'f2f', 'onsite', 'on-site', 'hotlist']

# # # Setting the date range for fetching emails: emails from the last 16 hours
# # today = datetime.today()
# # since_date = (today - timedelta(hours=16)).strftime("%d-%b-%Y")

# # # Load Excel file into DataFrames
# # df_sheet1 = pd.read_excel(filePath, sheet_name='Sheet1')  # Load job-related data
# # df_sheet2 = pd.read_excel(filePath, sheet_name='Sheet2')  # Load keywords related to job filtering

# # # Start an infinite loop to continuously check for new emails
# # while True:
# #     myDataTable = pd.DataFrame(columns=['To Email'])  # Initialize a new DataFrame for each email

# #     try:
# #         # Connect to the email server using IMAP
# #         mail = imaplib.IMAP4_SSL(imap_url)
# #         mail.login(username, password)
# #         mail.select("inbox")
# #         status, response = mail.search(None, '(UNSEEN)', f'SINCE "{since_date}"')

# #         # Process each email found in the search results
# #         for num in response[0].split():
# #             time.sleep(1)
# #             mail.store(num, '+FLAGS', '\\Seen')
# #             status, data = mail.fetch(num, "(RFC822)")

# #             if status != 'OK':
# #                 print(f"Failed to fetch message {num}. Status code: {status}")
# #                 continue

# #             email_message = email.message_from_bytes(data[0][1]) if isinstance(data[0][1], bytes) else email.message_from_string(data[0][1])
# #             email_subject = email_message['subject'].lower()

# #             # Skip emails with rejected keywords
# #             if re.search(r'\b(?:{})\b'.format('|'.join(map(re.escape, RejectedKeywords))), email_subject):
# #                 continue

# #             email_list = [email_message['Reply-To'], email_message['From']]
# #             payload = email_message.get_payload()

# #             if isinstance(payload, list):
# #                 body = next(part.get_payload(decode=True).decode(part.get_content_charset()) for part in payload if part.get_content_type() == "text/plain")
# #             else:
# #                 body = payload if isinstance(payload, str) else payload.decode()

# #             email_addresses = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', body)
# #             email_list.extend(email_addresses)

# #             # Remove duplicates and filter valid emails
# #             email_list = list(set(email_list))  # Remove duplicates
# #             valid_emails = [item for item in email_list if item and "phmailadmin" not in item and ".email" not in item]

# #             # Debug output
# #             print("Valid Emails:", valid_emails)

# #             # Select the first valid email
# #             toEmail = valid_emails[0] if valid_emails else None

# #             # If a valid email address is found and it's not already in the DataFrame
# #             if toEmail and toEmail not in myDataTable['To Email'].values:
# #                 new_entry = pd.DataFrame({'To Email': [toEmail]})
# #                 myDataTable = pd.concat([myDataTable, new_entry], ignore_index=True)

# #                 # Find matching rows in df_sheet2 based on keywords
# #                 matching_rows = []
# #                 for index, row in df_sheet2.iterrows():
# #                     if isinstance(row['Keywords'], str) and any(keyword.strip().lower() in email_subject for keyword in row['Keywords'].split(',')):
# #                         matching_rows += df_sheet1[df_sheet1['Technology'] == row['Technology']].values.tolist()

# #                 # Process each matching row found
# #                 for row in matching_rows:
# #                     if len(row) >= 7:
# #                         empTech, employeeName, empTech, employeeStatus, employeeUsername, employeePassword, employeeEmailBody = row[:7]

# #                         body = employeeEmailBody.replace("ThanksRemove/unsubscribe  |  Update your contact and subscribed mailing list(s)  |  Subscribe to mailing list(s) to receive requirements & resumes", "").strip()

# #                         # Function to remove hyperlinks from the email body
# #                         def remove_hyperlinks(text):
# #                             return re.sub(r'<(?:https?://)?[^|>]+>', '', text).strip()

# #                         body = remove_hyperlinks(body)
# #                         phone_numbers = re.findall(r"(?:\(?\d{3}\)?-? *\d{3}-? *-?\d{4}|\b\d{4}-\d{6}\b|\b\d{4}-\d{3}-\d{3,4}\b|\b\d{3}[-\s]?\d{3}[-\s]?\d{4}\b)", body)

# #                         if phone_numbers:
# #                             selfReply = MIMEText(body, 'plain')
# #                             selfReply["Subject"] = email_subject + " - call vendor"
# #                             selfReply["To"] = employeeUsername
# #                             selfReply["Importance"] = "High"

# #                             smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
# #                             smtp_server.starttls()
# #                             smtp_server.login(username, password)
# #                             smtp_server.sendmail(username, employeeUsername, selfReply.as_string())
# #                             smtp_server.close()

# #                         # Prepare and send a reply to the original email using the first email in the DataFrame
# #                         if not myDataTable.empty:
# #                             toEmail = myDataTable['To Email'].iloc[0]  # Use the first email in myDataTable
# #                             reply = MIMEText(body, 'plain')
# #                             reply["Subject"] = email_subject
# #                             reply["To"] = toEmail
# #                             smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
# #                             smtp_server.starttls()
# #                             smtp_server.login(employeeUsername, employeePassword)
# #                             smtp_server.sendmail(employeeUsername, toEmail, reply.as_string())
# #                             mail.store(num, "+FLAGS", "\\Seen")
# #                             mail.store(num, '+FLAGS', '\\Deleted')
# #                             mail.expunge()
# #                             print(f"Sent reply on behalf of: {employeeName} to vendor email: {toEmail}")

# #     except imaplib.IMAP4.abort:
# #         print("Error fetching email. Retrying in 10 seconds.")
# #         time.sleep(10)

# #     print("Will wait 500 sec and restart job")
# #     time.sleep(500)  # Adjusted wait time to 500 seconds

# # # Cleanup on exit
# # mail.close()
# # mail.logout()



# import imaplib  # For connecting to the email server
# import smtplib  # For sending emails via SMTP
# import email  # For parsing email messages
# import time  # For sleep functionality
# import re  # For regular expressions
# import pandas as pd  # For handling data in tabular format (Excel)
# from email.mime.text import MIMEText  # For creating MIME text emails
# from datetime import datetime, timedelta  # For handling dates and times

# # Configuration for email server and credentials
# imap_url = 'imap.gmail.com'  # URL for the IMAP server
# username = 'dailyrequriments@gmail.com'  # Email address for login
# password = 'yulyeentyuykfpav'  # Password for login
# filePath = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\old_process\SendEmailJobInput.xlsx'

# # List of keywords that will cause emails to be skipped
# RejectedKeywords = ['w2', 'architect', 'fulltime', 'dice', 'gc', 'citizen', 'full-time', 'lead', 'gc-ead', 'usc', 'manager', 'f2f', 'onsite', 'on-site', 'hotlist']
# # List of restricted keywords for filtering
# restricted_keywords = ["phmailadmin", ".email", "restricted_domain.com", "example.com", "nvoids.com"]

# # Setting the date range for fetching emails: emails from the last 16 hours
# today = datetime.today()
# since_date = (today - timedelta(hours=16)).strftime("%d-%b-%Y")

# # Load Excel file into DataFrames
# df_sheet1 = pd.read_excel(filePath, sheet_name='Sheet1')  # Load job-related data
# df_sheet2 = pd.read_excel(filePath, sheet_name='Sheet2')  # Load keywords related to job filtering

# # Start an infinite loop to continuously check for new emails
# while True:
#     myDataTable = pd.DataFrame(columns=['To Email'])  # Initialize a new DataFrame for each email

#     try:
#         # Connect to the email server using IMAP
#         mail = imaplib.IMAP4_SSL(imap_url)
#         mail.login(username, password)
#         mail.select("inbox")
#         status, response = mail.search(None, '(UNSEEN)', f'SINCE "{since_date}"')

#         # Process each email found in the search results
#         for num in response[0].split():
#             time.sleep(1)
#             mail.store(num, '+FLAGS', '\\Seen')
#             status, data = mail.fetch(num, "(RFC822)")

#             if status != 'OK':
#                 print(f"Failed to fetch message {num}. Status code: {status}")
#                 continue

#             email_message = email.message_from_bytes(data[0][1]) if isinstance(data[0][1], bytes) else email.message_from_string(data[0][1])
#             email_subject = email_message['subject'].lower()

#             # Skip emails with rejected keywords
#             if re.search(r'\b(?:{})\b'.format('|'.join(map(re.escape, RejectedKeywords))), email_subject):
#                 continue

#             email_list = [email_message['Reply-To'], email_message['From']]
#             payload = email_message.get_payload()

#             if isinstance(payload, list):
#                 body = next(part.get_payload(decode=True).decode(part.get_content_charset()) for part in payload if part.get_content_type() == "text/plain")
#             else:
#                 body = payload if isinstance(payload, str) else payload.decode()

#             email_addresses = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', body)
#             email_list.extend(email_addresses)

#             # Remove duplicates and filter valid emails
#             email_list = list(set(email_list))  # Remove duplicates
#             valid_emails = [item for item in email_list if item and "phmailadmin" not in item and ".email" not in item]

#             # Filter out restricted emails
#             valid_emails = [email for email in valid_emails if not any(restricted in email for restricted in restricted_keywords)]

#             # Debug output
#             print("Valid Emails:", valid_emails)

#             # Select the first valid email
#             toEmail = valid_emails[0] if valid_emails else None

#             # If a valid email address is found and it's not already in the DataFrame
#             if toEmail and toEmail not in myDataTable['To Email'].values:
#                 new_entry = pd.DataFrame({'To Email': [toEmail]})
#                 myDataTable = pd.concat([myDataTable, new_entry], ignore_index=True)

#                 # Find matching rows in df_sheet2 based on keywords
#                 matching_rows = []
#                 for index, row in df_sheet2.iterrows():
#                     if isinstance(row['Keywords'], str) and any(keyword.strip().lower() in email_subject for keyword in row['Keywords'].split(',')):
#                         matching_rows += df_sheet1[df_sheet1['Technology'] == row['Technology']].values.tolist()

#                 # Process each matching row found
#                 for row in matching_rows:
#                     if len(row) >= 7:
#                         empTech, employeeName, empTech, employeeStatus, employeeUsername, employeePassword, employeeEmailBody = row[:7]

#                         body = employeeEmailBody.replace("ThanksRemove/unsubscribe  |  Update your contact and subscribed mailing list(s)  |  Subscribe to mailing list(s) to receive requirements & resumes", "").strip()

#                         # Function to remove hyperlinks from the email body
#                         def remove_hyperlinks(text):
#                             return re.sub(r'<(?:https?://)?[^|>]+>', '', text).strip()

#                         body = remove_hyperlinks(body)
#                         phone_numbers = re.findall(r"(?:\(?\d{3}\)?-? *\d{3}-? *-?\d{4}|\b\d{4}-\d{6}\b|\b\d{4}-\d{3}-\d{3,4}\b|\b\d{3}[-\s]?\d{3}[-\s]?\d{4}\b)", body)

#                         if phone_numbers:
#                             selfReply = MIMEText(body, 'plain')
#                             selfReply["Subject"] = email_subject + " - call vendor"
#                             selfReply["To"] = employeeUsername
#                             selfReply["Importance"] = "High"

#                             smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
#                             smtp_server.starttls()
#                             smtp_server.login(username, password)
#                             smtp_server.sendmail(username, employeeUsername, selfReply.as_string())
#                             smtp_server.close()

#                         # Prepare and send a reply to the original email using the first email in the DataFrame
#                         if not myDataTable.empty:
#                             toEmail = myDataTable['To Email'].iloc[0]  # Use the first email in myDataTable
#                             reply = MIMEText(body, 'plain')
#                             reply["Subject"] = email_subject
#                             reply["To"] = toEmail
#                             smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
#                             smtp_server.starttls()
#                             smtp_server.login(employeeUsername, employeePassword)
#                             smtp_server.sendmail(employeeUsername, toEmail, reply.as_string())
#                             mail.store(num, "+FLAGS", "\\Seen")
#                             mail.store(num, '+FLAGS', '\\Deleted')
#                             mail.expunge()
#                             print(f"Sent reply on behalf of: {employeeName} to vendor email: {toEmail}")

#     except imaplib.IMAP4.abort:
#         print("Error fetching email. Retrying in 10 seconds.")
#         time.sleep(10)

#     print("Will wait 500 sec and restart job")
#     time.sleep(500)  # Adjusted wait time to 500 seconds

# # Cleanup on exit
# mail.close()
# mail.logout()



import imaplib  # For connecting to the email server
import smtplib  # For sending emails via SMTP
import email  # For parsing email messages
import time  # For sleep functionality
import re  # For regular expressions
import pandas as pd  # For handling data in tabular format (Excel)
from email.mime.text import MIMEText  # For creating MIME text emails
from datetime import datetime, timedelta  # For handling dates and times

# Configuration for email server and credentials
imap_url = 'imap.gmail.com'  # URL for the IMAP server
username = 'dailyrequriments@gmail.com'  # Email address for login
password = 'yulyeentyuykfpav'  # Password for login
filePath = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\old\SendEmailJobInput.xlsx'

# List of keywords that will cause emails to be skipped
RejectedKeywords = ['w2', 'architect', 'fulltime', 'dice', 'gc', 'citizen', 'full-time', 'lead', 'gc-ead', 'usc', 'manager', 'f2f', 'onsite', 'on-site', 'hotlist']
# List of restricted keywords for filtering
restricted_keywords = ["phmailadmin", ".email", "restricted_domain.com", "example.com", "nvoids.com"]

# Setting the date range for fetching emails: emails from the last 16 hours
today = datetime.today()
since_date = (today - timedelta(hours=16)).strftime("%d-%b-%Y")

# Load Excel file into DataFrames
df_sheet1 = pd.read_excel(filePath, sheet_name='Sheet1')  # Load job-related data
df_sheet2 = pd.read_excel(filePath, sheet_name='Sheet2')  # Load keywords related to job filtering

# Start an infinite loop to continuously check for new emails
while True:
    myDataTable = pd.DataFrame(columns=['To Email'])  # Initialize a new DataFrame for each email

    try:
        # Connect to the email server using IMAP
        mail = imaplib.IMAP4_SSL(imap_url)
        mail.login(username, password)
        mail.select("inbox")
        status, response = mail.search(None, '(UNSEEN)', f'SINCE "{since_date}"')

        # Process each email found in the search results
        for num in response[0].split():
            time.sleep(1)
            mail.store(num, '+FLAGS', '\\Seen')
            status, data = mail.fetch(num, "(RFC822)")

            if status != 'OK':
                print(f"Failed to fetch message {num}. Status code: {status}")
                continue

            email_message = email.message_from_bytes(data[0][1]) if isinstance(data[0][1], bytes) else email.message_from_string(data[0][1])
            email_subject = email_message['subject'].lower()

            # Skip emails with rejected keywords
            if re.search(r'\b(?:{})\b'.format('|'.join(map(re.escape, RejectedKeywords))), email_subject):
                continue

            email_list = [email_message['Reply-To'], email_message['From']]
            payload = email_message.get_payload()

            if isinstance(payload, list):
                body = next(part.get_payload(decode=True).decode(part.get_content_charset()) for part in payload if part.get_content_type() == "text/plain")
            else:
                body = payload if isinstance(payload, str) else payload.decode()

            email_addresses = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', body)
            email_list.extend(email_addresses)

            # Remove duplicates and filter valid emails
            email_list = list(set(email_list))  # Remove duplicates
            valid_emails = [item for item in email_list if item and "phmailadmin" not in item and ".email" not in item]

            # Filter out restricted emails
            valid_emails = [email for email in valid_emails if not any(restricted in email for restricted in restricted_keywords)]

            # Debug output
            print("Valid Emails:", valid_emails)

            # Select the first valid email
            toEmail = valid_emails[0] if valid_emails else None

            # If a valid email address is found and it's not already in the DataFrame
            if toEmail and toEmail not in myDataTable['To Email'].values:
                new_entry = pd.DataFrame({'To Email': [toEmail]})
                myDataTable = pd.concat([myDataTable, new_entry], ignore_index=True)

                # Find matching rows in df_sheet2 based on keywords
                matching_rows = []
                for index, row in df_sheet2.iterrows():
                    if isinstance(row['Keywords'], str) and any(keyword.strip().lower() in email_subject for keyword in row['Keywords'].split(',')):
                        matching_rows += df_sheet1[df_sheet1['Technology'] == row['Technology']].values.tolist()

                # Process each matching row found
                for row in matching_rows:
                    if len(row) >= 7:
                        empTech, employeeName, empTech, employeeStatus, employeeUsername, employeePassword, employeeEmailBody = row[:7]

                        body = employeeEmailBody.replace("ThanksRemove/unsubscribe  |  Update your contact and subscribed mailing list(s)  |  Subscribe to mailing list(s) to receive requirements & resumes", "").strip()

                        # Function to remove hyperlinks from the email body
                        def remove_hyperlinks(text):
                            return re.sub(r'<(?:https?://)?[^|>]+>', '', text).strip()

                        body = remove_hyperlinks(body)
                        phone_numbers = re.findall(r"(?:\(?\d{3}\)?-? *\d{3}-? *-?\d{4}|\b\d{4}-\d{6}\b|\b\d{4}-\d{3}-\d{3,4}\b|\b\d{3}[-\s]?\d{3}[-\s]?\d{4}\b)", body)

                        if phone_numbers:
                            selfReply = MIMEText(body, 'plain')
                            selfReply["Subject"] = email_subject + " - call vendor"
                            selfReply["To"] = employeeUsername
                            selfReply["Importance"] = "High"

                            smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
                            smtp_server.starttls()
                            smtp_server.login(username, password)
                            smtp_server.sendmail(username, employeeUsername, selfReply.as_string())
                            smtp_server.close()

                        # Prepare and send a reply to the original email
                        if not myDataTable.empty:
                            toEmail = myDataTable['To Email'].iloc[0]  # Use the first email in myDataTable
                        else:
                            toEmail = email_message['Reply-To']  # Use the Reply-To address if DataFrame is empty

                        reply = MIMEText(body, 'plain')
                        reply["Subject"] = email_subject
                        reply["To"] = toEmail

                        smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
                        smtp_server.starttls()
                        smtp_server.login(employeeUsername, employeePassword)
                        smtp_server.sendmail(employeeUsername, toEmail, reply.as_string())
                        mail.store(num, "+FLAGS", "\\Seen")
                        mail.store(num, '+FLAGS', '\\Deleted')
                        mail.expunge()
                        print(f"Sent reply to: {toEmail}")

    except imaplib.IMAP4.abort:
        print("Error fetching email. Retrying in 10 seconds.")
        time.sleep(10)

    print("Will wait 500 sec and restart job")
    time.sleep(500)  # Adjusted wait time to 500 seconds

# Cleanup on exit
mail.close()
mail.logout()


# # # # # # # #v1
# # # # # import imaplib
# # # # # import smtplib
# # # # # import email
# # # # # import time
# # # # # import base64
# # # # # import re
# # # # # import os
# # # # # import textwrap
# # # # # from bs4 import BeautifulSoup
# # # # # import pandas as pd
# # # # # import datetime
# # # # # from email.mime.text import MIMEText
# # # # # from datetime import datetime, timedelta


# # # # # imaplib._MAXLINE = 20000000

# # # # # # Set the date range for emails to be fetched
# # # # # today = datetime.today()
# # # # # CurrentTime = datetime.now()

# # # # # fifteen_hours_ago = today - timedelta(hours=16)
# # # # # date_format = "%d-%b-%Y"
# # # # # since_date = fifteen_hours_ago.strftime(date_format)
# # # # # until_date = today.strftime(date_format)
# # # # # max_retries = 3
# # # # # retry_count = 0
# # # # # skip_count = 0
# # # # # finalEmailBody  ='';

# # # # # RejectedKeywords = ['w2','architect',
# # # # # 'fulltime', 'dice',
# # # # # 'architect','gc',
# # # # # 'citizen',
# # # # # 'full-time','lead',
# # # # # 'gc-ead','usc','manager','f2f','onsite','on-site', 'hotlist','hot list','hot-list']

# # # # # filePath = '/Users/prudhvirajsomepalli/Documents/Code/SendNotifications/SendEmailJobInput.xlsx'
# # # # # # filePath = 'SendEmailJobInput.xlsx'
# # # # # # Load the Excel file into a Pandas dataframe
# # # # # df_sheet2 = pd.read_excel(filePath, sheet_name='Sheet2')
# # # # # # /Users/prudhvirajsomepalli/Documents/Code/SendNotifications/SendEmailJobInput.xlsx

# # # # # # Load the Excel file into a Pandas dataframe for Sheet1
# # # # # df_sheet1 = pd.read_excel(filePath, sheet_name='Sheet1')

# # # # # if os.path.exists('myExcelFile.xlsx'):
# # # # #     existingData = pd.read_excel('myExcelFile.xlsx')
# # # # #     existingData = existingData.drop_duplicates(subset=['To Email'])
# # # # #     existingData.to_excel('myExcelFile.xlsx', index=False)



# # # # # # # Login information
# # # # # username = 'dailyrequriments@gmail.com'
# # # # # password = 'yulyeentyuykfpav'
# # # # # imap_url = 'imap.gmail.com'



# # # # # # Create an empty DataFrame to add emails to excel
# # # # # myDataTable = pd.DataFrame(columns=['To Email'])
# # # # # while True:

# # # # #     try:
# # # # #         # Connect to the server
# # # # #         mail = imaplib.IMAP4_SSL(imap_url)
# # # # #         mail.login(username, password)

# # # # #         time_limit = datetime.now() - timedelta(minutes=10)
# # # # #         since_date = time_limit.strftime('%d-%b-%Y')  # Simplified date format


# # # # #         # Search for emails since the calculated time
# # # # #         status, response = mail.search(None, f'SINCE "{since_date}"')
# # # # #         # Get the list of email IDs
# # # # #         email_ids = response[0].split()

# # # # #         for email_id in email_ids:
# # # # #             status, msg_data = mail.fetch(email_id, '(RFC822)')
# # # # #             # Process the email data
# # # # #             print(f"Email ID: {email_id}")
# # # # #             print(f"Message Data: {msg_data}")
# # # # #     except imaplib.IMAP4.abort as ex:
# # # # #         print("Error fetching email. Retrying in 10 seconds.")





# # # # import imaplib
# # # # import os
# # # # import pandas as pd
# # # # from datetime import datetime, timedelta
# # # # import time

# # # # # Set the date range for emails to be fetched
# # # # today = datetime.today()
# # # # time_limit = datetime.now() - timedelta(minutes=10)
# # # # since_date = time_limit.strftime('%d-%b-%Y')  # Simplified date format

# # # # # # Login information
# # # # username = 'dailyrequriments@gmail.com'
# # # # password = 'yulyeentyuykfpav'
# # # # imap_url = 'imap.gmail.com'

# # # # # Path for the Excel file
# # # # filePath = '/Users/prudhvirajsomepalli/Documents/Code/SendNotifications/SendEmailJobInput.xlsx'

# # # # # Load the Excel file into a Pandas dataframe
# # # # df_sheet1 = pd.read_excel(filePath, sheet_name='Sheet1')

# # # # # Create an empty DataFrame to add emails to excel
# # # # myDataTable = pd.DataFrame(columns=['To Email'])

# # # # while True:
# # # #     try:
# # # #         # Connect to the server
# # # #         mail = imaplib.IMAP4_SSL(imap_url)
# # # #         mail.login(username, password)

# # # #         # Select the mailbox
# # # #         status, _ = mail.select('inbox')
# # # #         if status != 'OK':
# # # #             print(f"Failed to select mailbox: {status}")
# # # #             mail.logout()
# # # #             continue

# # # #         # Search for emails since the calculated time
# # # #         status, response = mail.search(None, f'SINCE {since_date}')
# # # #         if status != 'OK':
# # # #             print(f"Search failed: {status}")
# # # #             mail.logout()
# # # #             continue

# # # #         # Get the list of email IDs
# # # #         email_ids = response[0].split()

# # # #         for email_id in email_ids:
# # # #             status, msg_data = mail.fetch(email_id, '(RFC822)')
# # # #             if status != 'OK':
# # # #                 print(f"Failed to fetch email ID {email_id}: {status}")
# # # #                 continue

# # # #             # Process the email data
# # # #             print(f"Email ID: {email_id}")
# # # #             print(f"Message Data: {msg_data}")

# # # #         # Logout and close the connection
# # # #         mail.logout()

# # # #     except imaplib.IMAP4.abort as ex:
# # # #         print("Error fetching email. Retrying in 10 seconds.")
# # # #         time.sleep(10)
# # # #     except Exception as ex:
# # # #         print(f"An unexpected error occurred: {ex}")
# # # #         time.sleep(10)





# # # import imaplib
# # # import os
# # # import pandas as pd
# # # from datetime import datetime, timedelta
# # # import email
# # # from email.header import decode_header
# # # import re
# # # import time
# # # from filelock import FileLock

# # # # Set the date range for emails to be fetched
# # # today = datetime.today()
# # # time_limit = datetime.now() - timedelta(minutes=10)
# # # since_date = time_limit.strftime('%d-%b-%Y')  # Simplified date format

# # # # Login information
# # # username = 'dailyrequriments@gmail.com'
# # # password = 'yulyeentyuykfpav'
# # # imap_url = 'imap.gmail.com'

# # # # Path for the Excel file
# # # filePath = '/Users/prudhvirajsomepalli/Documents/Code/vendorEmails.xlsx'
# # # lockPath = '/Users/prudhvirajsomepalli/Documents/Code/vendorEmails.lock'

# # # def fetch_emails():
# # # # Load the Excel file into a Pandas dataframe
# # #     df_sheet1 = pd.read_excel(filePath, sheet_name='Sheet1')

# # #     # Create an empty DataFrame to add emails to excel
# # #     myDataTable = pd.DataFrame(columns=['To Email'])
# # #      # Create a lock object
# # #     lock = FileLock(lockPath)


# # #     try:
# # #         with lock:
# # #             # Connect to the server
# # #             mail = imaplib.IMAP4_SSL(imap_url)
# # #             mail.login(username, password)

# # #             # Select the mailbox
# # #             status, _ = mail.select('inbox')
# # #             if status != 'OK':
# # #                 print(f"Failed to select mailbox: {status}")
# # #                 mail.logout()
# # #                 continue

# # #             # Search for emails since the calculated time
# # #             status, response = mail.search(None, f'SINCE {since_date}')
# # #             if status != 'OK':
# # #                 print(f"Search failed: {status}")
# # #                 mail.logout()
# # #                 continue

# # #             # Get the list of email IDs
# # #             email_ids = response[0].split()

# # #             for num in email_ids:
# # #                 # Limit the number of emails processed per minute
# # #                 time.sleep(1)

# # #                 status, data = mail.fetch(num, "(RFC822)")

# # #                 if status == 'OK':
# # #                     # Check if data[0][1] is a string or bytes object
# # #                     try:
# # #                         if isinstance(data[0][1], str):
# # #                             email_message = email.message_from_string(data[0][1])
# # #                         elif isinstance(data[0][1], bytes):
# # #                             email_message = email.message_from_bytes(data[0][1])
# # #                     except:
# # #                         continue
# # #                     # Initialize email list and email body
# # #                     email_list = []
# # #                     toEmail = ''
# # #                     payload = email_message.get_payload()

# # #                     # Define strings to replace in the email body
# # #                     string_to_replace = "Sign-Up for your account with PROHIRES POWERHOUSE Recruiting Portal to broadcast requirements & hotlists."
# # #                     string_to_replace2 = "ThanksRemove/unsubscribe  |  Update your contact and subscribed mailing list(s)  |  Subscribe to mailing list(s) to receive requirements & resumes"

# # #                     # Extract and clean email body
# # #                     if isinstance(payload, list):
# # #                         for part in payload:
# # #                             if part.get_content_type() == "text/plain":
# # #                                 body = part.get_payload(decode=True).decode(part.get_content_charset())
# # #                                 finalEmailBody = body.replace(string_to_replace, " ")
# # #                                 break
# # #                     else:
# # #                         finalEmailBody = payload
# # #                         finalEmailBody = finalEmailBody.replace(string_to_replace, " ")

# # #                     # Extract email addresses from the email body
# # #                     email_addresses = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', finalEmailBody)

# # #                     # Add email addresses to the list
# # #                     for email_address in email_addresses:
# # #                         email_list.append(email_address)
# # #                         break

# # #                     # Find the appropriate email address    
# # #                     for item in email_list:
# # #                         if item is None or not item:
# # #                             print(f'{item} Item is empty.')
# # #                         elif "phmailadmin" in item or ".email" in item:
# # #                             continue
# # #                         else:
# # #                             toEmail = item
# # #                             break

# # #                     if toEmail:
# # #                         #print(f"To Email: {toEmail}")

# # #                         if 'googlegroups.com' in toEmail:
# # #                         # Your code here
# # #                             print(f'The email {toEmail} contains googlegroups.com')
# # #                         else:
# # #                             #print(f'The email {toEmail} does not contain googlegroups.com')
                            
# # #                             # Check if the 'To Email' column exists, create it if not
# # #                             if 'To Email' not in df_sheet1.columns:
# # #                                 df_sheet1['To Email'] = pd.Series(dtype='str')

# # #                             # Check if the email already exists in the DataFrame
# # #                             if toEmail not in df_sheet1['To Email'].values:
# # #                                 # Append the new email to the DataFrame
# # #                                 new_row = pd.DataFrame({'To Email': [toEmail]})
# # #                                 df_sheet1 = pd.concat([df_sheet1, new_row], ignore_index=True)
                                
# # #                                 # Save the updated DataFrame back to the Excel file
# # #                                 df_sheet1.to_excel(filePath, sheet_name='Sheet1', index=False)
# # #                                 print(f'Email {toEmail} added successfully.')
# # #                             else:
# # #                                 print(f'Email {toEmail} already exists in the file.')

# # #                 else:
# # #                     print(f"Failed to fetch message {num}. Status code: {status}")

# # #             # Logout and close the connection
# # #             mail.logout()

# # #         except imaplib.IMAP4.abort as ex:
# # #             print("Error fetching email. Retrying in 10 seconds.")
# # #             time.sleep(10)
# # #         except Exception as ex:
# # #             print(f"An unexpected error occurred: {ex}")
# # #             time.sleep(10)
# # # while True:
# # #     fetch_emails()
# # #     time.sleep(600)  # Sleep for 10 minutes


# # import imaplib
# # import os
# # import pandas as pd
# # from datetime import datetime, timedelta
# # import email
# # from email.header import decode_header
# # import re
# # import time
# # from filelock import FileLock

# # # Path for the Excel file and lock file
# # filePath = '/Users/prudhvirajsomepalli/Documents/Code/vendorEmails.xlsx'
# # lockPath = '/Users/prudhvirajsomepalli/Documents/Code/vendorEmails.lock'

# # def fetch_emails():
# #     # Set the date range for emails to be fetched
# #     today = datetime.today()
# #     time_limit = datetime.now() - timedelta(minutes=10)
# #     since_date = time_limit.strftime('%d-%b-%Y')  # Simplified date format

# #     # Login information
# #     username = 'dailyrequriments@gmail.com'
# #     password = 'yulyeentyuykfpav'
# #     imap_url = 'imap.gmail.com'

# #     # Create a lock object
# #     lock = FileLock(lockPath)

# #     try:
# #         with lock:
# #             # Load the Excel file into a Pandas dataframe
# #             df_sheet1 = pd.read_excel(filePath, sheet_name='Sheet1')

# #             # Create an empty DataFrame to add emails to excel
# #             myDataTable = pd.DataFrame(columns=['To Email'])

# #             # Connect to the server
# #             mail = imaplib.IMAP4_SSL(imap_url)
# #             mail.login(username, password)

# #             # Select the mailbox
# #             status, _ = mail.select('inbox')
# #             if status != 'OK':
# #                 print(f"Failed to select mailbox: {status}")
# #                 mail.logout()
# #                 return

# #             # Search for emails since the calculated time
# #             status, response = mail.search(None, f'SINCE {since_date}')
# #             if status != 'OK':
# #                 print(f"Search failed: {status}")
# #                 mail.logout()
# #                 return

# #             # Get the list of email IDs
# #             email_ids = response[0].split()

# #             for num in email_ids:
# #                 # Limit the number of emails processed per minute
# #                 time.sleep(1)

# #                 status, data = mail.fetch(num, "(RFC822)")

# #                 if status == 'OK':
# #                     # Check if data[0][1] is a string or bytes object
# #                     try:
# #                         if isinstance(data[0][1], str):
# #                             email_message = email.message_from_string(data[0][1])
# #                         elif isinstance(data[0][1], bytes):
# #                             email_message = email.message_from_bytes(data[0][1])
# #                     except:
# #                         continue
# #                     # Initialize email list and email body
# #                     email_list = []
# #                     toEmail = ''
# #                     payload = email_message.get_payload()

# #                     # Define strings to replace in the email body
# #                     string_to_replace = "Sign-Up for your account with PROHIRES POWERHOUSE Recruiting Portal to broadcast requirements & hotlists."
# #                     string_to_replace2 = "ThanksRemove/unsubscribe  |  Update your contact and subscribed mailing list(s)  |  Subscribe to mailing list(s) to receive requirements & resumes"

# #                     # Extract and clean email body
# #                     if isinstance(payload, list):
# #                         for part in payload:
# #                             if part.get_content_type() == "text/plain":
# #                                 body = part.get_payload(decode=True).decode(part.get_content_charset())
# #                                 finalEmailBody = body.replace(string_to_replace, " ")
# #                                 break
# #                     else:
# #                         finalEmailBody = payload
# #                         finalEmailBody = finalEmailBody.replace(string_to_replace, " ")

# #                     # Extract email addresses from the email body
# #                     email_addresses = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', finalEmailBody)

# #                     # Add email addresses to the list
# #                     for email_address in email_addresses:
# #                         email_list.append(email_address)
# #                         break

# #                     # Find the appropriate email address    
# #                     for item in email_list:
# #                         if item is None or not item:
# #                             print(f'{item} Item is empty.')
# #                         elif "phmailadmin" in item or ".email" in item:
# #                             continue
# #                         else:
# #                             toEmail = item
# #                             break

# #                     if toEmail:
# #                         print(f"To Email: {toEmail}")

# #                         if 'googlegroups.com' in toEmail:
# #                             # Your code here
# #                             print(f'The email {toEmail} contains googlegroups.com')
# #                         else:
# #                             print(f'The email {toEmail} does not contain googlegroups.com')

# #                             # Check if the 'To Email' column exists, create it if not
# #                             if 'To Email' not in df_sheet1.columns:
# #                                 df_sheet1['To Email'] = pd.Series(dtype='str')

# #                             # Check if the email already exists in the DataFrame
# #                             if toEmail not in df_sheet1['To Email'].values:
# #                                 # Append the new email to the DataFrame
# #                                 new_row = pd.DataFrame({'To Email': [toEmail]})
# #                                 df_sheet1 = pd.concat([df_sheet1, new_row], ignore_index=True)

# #                                 # Save the updated DataFrame back to the Excel file
# #                                 df_sheet1.to_excel(filePath, sheet_name='Sheet1', index=False)
# #                                 print(f'Email {toEmail} added successfully.')
# #                             else:
# #                                 print(f'Email {toEmail} already exists in the file.')

# #                 else:
# #                     print(f"Failed to fetch message {num}. Status code: {status}")

# #             # Logout and close the connection
# #             mail.logout()

# #     except imaplib.IMAP4.abort as ex:
# #         print("Error fetching email. Retrying in 10 seconds.")
# #         time.sleep(10)
# #     except Exception as ex:
# #         print(f"An unexpected error occurred: {ex}")
# #         time.sleep(10)

# # # Run the fetch_emails function every 10 minutes
# # while True:
# #     fetch_emails()
# #     time.sleep(600)  # Sleep for 10 minutes


# # import imaplib
# # import pandas as pd
# # from datetime import datetime, timedelta
# # import re
# # import threading
# # import time
# # from filelock import FileLock
# # import email

# # # Path for the Excel file and lock file
# # filePath = '/Users/prudhvirajsomepalli/Documents/Code/vendorEmails.xlsx'
# # lockPath = '/Users/prudhvirajsomepalli/Documents/Code/vendorEmails.lock'

# # def fetch_emails():
# #     # Set the date range for emails to be fetched
# #     time_limit = datetime.now() - timedelta(minutes=10)
# #     since_date = time_limit.strftime('%d-%b-%Y')  # Simplified date format

# #     # Login information
# #     username = 'dailyrequriments@gmail.com'
# #     password = 'yulyeentyuykfpav'
# #     imap_url = 'imap.gmail.com'

# #     # Load the Excel file into a Pandas dataframe
# #     df_sheet1 = pd.read_excel(filePath, sheet_name='Sheet1')

# #     # Create a lock object
# #     lock = FileLock(lockPath)

# #     while True:
# #         try:
# #             # Connect to the server
# #             mail = imaplib.IMAP4_SSL(imap_url)
# #             mail.login(username, password)
# #             mail.select('inbox')

# #             # Search for emails since the calculated time
# #             status, response = mail.search(None, f'SINCE {since_date}')
# #             if status != 'OK':
# #                 print(f"Search failed: {status}")
# #                 continue

# #             # Get the list of email IDs
# #             email_ids = response[0].split()

# #             for num in email_ids:
# #                 # Limit the number of emails processed per minute
# #                 time.sleep(1)

# #                 status, data = mail.fetch(num, "(RFC822)")
# #                 if status == 'OK':
# #                     email_message = email.message_from_bytes(data[0][1])
# #                     payload = email_message.get_payload()

# #                     # Define strings to replace in the email body
# #                     string_to_replace = "Sign-Up for your account with PROHIRES POWERHOUSE Recruiting Portal to broadcast requirements & hotlists."

# #                     # Extract and clean email body
# #                     if isinstance(payload, list):
# #                         for part in payload:
# #                             if part.get_content_type() == "text/plain":
# #                                 body = part.get_payload(decode=True).decode(part.get_content_charset())
# #                                 finalEmailBody = body.replace(string_to_replace, " ")
# #                                 break
# #                     else:
# #                         finalEmailBody = payload
# #                         finalEmailBody = finalEmailBody.replace(string_to_replace, " ")

# #                     # Extract email addresses from the email body
# #                     email_addresses = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', finalEmailBody)

# #                     for email_address in email_addresses:
# #                         if email_address and "phmailadmin" not in email_address and ".email" not in email_address:
# #                             toEmail = email_address
# #                             break
# #                     else:
# #                         continue

# #                     #print(f"To Email: {toEmail}")

# #                     if 'googlegroups.com' in toEmail:
# #                         print(f'The email {toEmail} contains googlegroups.com')
# #                     else:
# #                         #print(f'The email {toEmail} does not contain googlegroups.com')

# #                         # Use the synchronous file lock to handle file operations
# #                         with lock:
# #                             # Check if the email already exists in the DataFrame
# #                             if toEmail not in df_sheet1['To Email'].values:
# #                                 # Append the new email to the DataFrame
# #                                 new_row = pd.DataFrame({'To Email': [toEmail]})
# #                                 df_sheet1 = pd.concat([df_sheet1, new_row], ignore_index=True)

# #                                 # Save the updated DataFrame back to the Excel file
# #                                 df_sheet1.to_excel(filePath, sheet_name='Sheet1', index=False)
# #                                 print(f'Email {toEmail} added successfully.')
# #                             else:
# #                                 print(f'Email {toEmail} already exists in the file.')

# #             # Logout and close the connection
# #             mail.logout()

# #         except imaplib.IMAP4.abort as e:
# #             print(f"IMAP4 abort error occurred: {e}")
# #         except imaplib.IMAP4.error as e:
# #             print(f"IMAP4 error occurred: {e}")
# #         except Exception as e:
# #             print(f"An unexpected error occurred: {e}")
        
# #         time.sleep(600)  # Sleep for 10 minutes

# # def main():
# #     fetch_thread = threading.Thread(target=fetch_emails)
# #     fetch_thread.start()
# #     fetch_thread.join()

# # if __name__ == "__main__":
# #     main()




# import imaplib
# import pandas as pd
# from datetime import datetime, timedelta
# import re
# import threading
# import time
# from filelock import FileLock
# import email

# # Path for the Excel file and lock file
# filePath = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\vendorEmails.xlsx'
# lockPath = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\vendorEmails.lock'

# def fetch_emails(last_run_time):
#     # Set the date range for emails to be fetched
#     since_date = last_run_time.strftime('%d-%b-%Y')

#     # Login information
#     username = 'dailyrequriments@gmail.com'
#     password = 'yulyeentyuykfpav'
#     imap_url = 'imap.gmail.com'

#     # Load the Excel file into a Pandas dataframe
#     df_sheet1 = pd.read_excel(filePath, sheet_name='Sheet1')

#     # Create a lock object
#     lock = FileLock(lockPath)

#     try:
#         # Connect to the server
#         mail = imaplib.IMAP4_SSL(imap_url)
#         mail.login(username, password)
#         mail.select('inbox')

#         # Search for emails since the calculated time
#         status, response = mail.search(None, f'SINCE {since_date}')
#         if status != 'OK':
#             print(f"Search failed: {status}")
#             return

#         # Get the list of email IDs
#         email_ids = response[0].split()

#         for num in email_ids:
#             # Limit the number of emails processed per minute
#             time.sleep(1)

#             status, data = mail.fetch(num, "(RFC822)")
#             if status == 'OK':
#                 email_message = email.message_from_bytes(data[0][1])
#                 payload = email_message.get_payload()

#                 # Define strings to replace in the email body
#                 string_to_replace = "Sign-Up for your account with PROHIRES POWERHOUSE Recruiting Portal to broadcast requirements & hotlists."

#                 # Extract and clean email body
#                 if isinstance(payload, list):
#                     for part in payload:
#                         if part.get_content_type() == "text/plain":
#                             body = part.get_payload(decode=True).decode(part.get_content_charset())
#                             finalEmailBody = body.replace(string_to_replace, " ")
#                             break
#                 else:
#                     finalEmailBody = payload
#                     finalEmailBody = finalEmailBody.replace(string_to_replace, " ")

#                 # Extract email addresses from the email body
                    
#                 email_addresses = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', finalEmailBody)

#                 for email_address in email_addresses:
#                     if email_address and "phmailadmin" not in email_address and ".email" not in email_address:
#                         toEmail = email_address
#                         break
#                 else:
#                     continue

#                 print(f"To Email: {toEmail}")

#                 if 'googlegroups.com' in toEmail:
#                     print(f'The email {toEmail} contains googlegroups.com')
#                 else:
#                    # print(f'The email {toEmail} does not contain googlegroups.com')

#                     # Use the synchronous file lock to handle file operations
#                     with lock:
#                         # Check if the email already exists in the DataFrame
#                         if toEmail not in df_sheet1['To Email'].values:
#                             # Append the new email to the DataFrame
#                             new_row = pd.DataFrame({'To Email': [toEmail]})
#                             df_sheet1 = pd.concat([df_sheet1, new_row], ignore_index=True)

#                             # Save the updated DataFrame back to the Excel file
#                             df_sheet1.to_excel(filePath, sheet_name='Sheet1', index=False)
#                             print(f'Email {toEmail} added successfully.')
#                         else:
#                             print(f'Email {toEmail} already exists in the file.')

#         # Logout and close the connection
#         mail.logout()

#     except imaplib.IMAP4.abort as e:
#         print(f"IMAP4 abort error occurred: {e}")
#     except imaplib.IMAP4.error as e:
#         print(f"IMAP4 error occurred: {e}")
#     except Exception as e:
#         print(f"An unexpected error occurred: {e}")

# def main():
#     last_run_time = datetime.now()
#     while True:
#         fetch_emails(last_run_time)
#         last_run_time = datetime.now()
#         print(f"waiting from: {last_run_time}")
#         time.sleep(600)  # Sleep for 10 minutes


# if __name__ == "__main__":
#     main()

import imaplib
import pandas as pd
from datetime import datetime
import re
import time
from filelock import FileLock
import email

# Path for the Excel file and lock file
filePath = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\vendorEmails.xlsx'
lockPath = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\vendorEmails.lock'

def fetch_emails(last_run_time):
    # Set the date range for emails to be fetched
    since_date = last_run_time.strftime('%d-%b-%Y')

    # # Login information
    username = 'dailyrequriments@gmail.com'
    password = 'yulyeentyuykfpav'
    imap_url = 'imap.gmail.com'

    # username = 'sandhyacareerr@gmail.com'
    # password = 'wixngueilqtdhpro'
    # imap_url = 'imap.gmail.com'
  


    # Load the Excel file into a Pandas dataframe
    df_sheet1 = pd.read_excel(filePath, sheet_name='Sheet1')

    # Create a lock object
    lock = FileLock(lockPath)

    try:
        with imaplib.IMAP4_SSL(imap_url) as mail:
            mail.login(username, password)
            mail.select('inbox')

            # Search for emails since the calculated time
            status, response = mail.search(None, f'SINCE {since_date}')
            if status != 'OK':
                print(f"Search failed: {status}")
                return

            # Get the list of email IDs
            email_ids = response[0].split()

            for num in email_ids:
                # Limit the number of emails processed per minute
                time.sleep(1)

                status, data = mail.fetch(num, "(RFC822)")
                if status == 'OK':
                    email_message = email.message_from_bytes(data[0][1])
                    payload = email_message.get_payload()

                    # Define strings to replace in the email body
                    string_to_replace = "Sign-Up for your account with PROHIRES POWERHOUSE Recruiting Portal to broadcast requirements & hotlists."

                    # Extract and clean email body
                    if isinstance(payload, list):
                        for part in payload:
                            if part.get_content_type() == "text/plain":
                                body = part.get_payload(decode=True).decode(part.get_content_charset(), errors='replace')
                                finalEmailBody = body.replace(string_to_replace, " ")
                                break
                        else:
                            finalEmailBody = ''
                    else:
                        finalEmailBody = payload
                        finalEmailBody = finalEmailBody.replace(string_to_replace, " ")

                    # Extract email addresses from the email body
                    email_addresses = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', finalEmailBody)

                    toEmail = None
                    # for email_address in email_addresses:
                    #     if email_address and "phmailadmin" not in email_address and ".email" not in email_address:
                    #         toEmail = email_address
                    #         break

                    restricted_keywords = ["phmailadmin", ".email", "restricted_domain.com", "example.com","nvoids.com"]
                    for email_address in email_addresses:
                        if email_address and not any(keyword in email_address for keyword in restricted_keywords):
                            toEmail = email_address
                            break
                    if toEmail is None:
                        continue

                    print(f"To Email: {toEmail}")

                    if 'googlegroups.com' in toEmail:
                        print(f'The email {toEmail} contains googlegroups.com')
                    else:
                        # Use the synchronous file lock to handle file operations
                        with lock:
                            # Check if the email already exists in the DataFrame
                            if toEmail not in df_sheet1['To Email'].values:
                                # Append the new email to the DataFrame
                                new_row = pd.DataFrame({'To Email': [toEmail]})
                                df_sheet1 = pd.concat([df_sheet1, new_row], ignore_index=True)

                                # Save the updated DataFrame back to the Excel file
                                df_sheet1.to_excel(filePath, sheet_name='Sheet1', index=False)
                                print(f'Email {toEmail} added successfully.')
                            else:
                                print(f'Email {toEmail} already exists in the file.')

    except imaplib.IMAP4.abort as e:
        print(f"IMAP4 abort error occurred: {e}")
    except imaplib.IMAP4.error as e:
        print(f"IMAP4 error occurred: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

def main():
    last_run_time = datetime.now()
    while True:
        fetch_emails(last_run_time)
        last_run_time = datetime.now()
        print(f"waiting from: {last_run_time}")
        time.sleep(600)  # Sleep for 10 minutes

if __name__ == "__main__":
    main()


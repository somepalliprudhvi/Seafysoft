# # # # # Email account details
# # # # # email_account_1 = 'seafysoftmail@gmail.com'
# # # # # password_1 = 'xauubupekdesdwpu'


# # # # import pandas as pd
# # # # import smtplib
# # # # from email.mime.text import MIMEText
# # # # from email.mime.multipart import MIMEMultipart
# # # # from filelock import FileLock

# # # # # Function to send emails
# # # # def send_email(sender_email, receiver_email, sender_password):
# # # #     try:
# # # #         # Setup the email message
# # # #         message = MIMEMultipart()
# # # #         message['From'] = sender_email
# # # #         message['To'] = receiver_email
# # # #         message['Subject'] = 'Seafy Soft HotList'

# # # #         # Email body
# # # #         body = 'This is an email body.'
# # # #         message.attach(MIMEText(body, 'plain'))

# # # #         # Sending email using SMTP
# # # #         with smtplib.SMTP('smtp.gmail.com', 587) as server:
# # # #             server.starttls()
# # # #             server.login(sender_email, sender_password)
# # # #             text = message.as_string()
# # # #             server.sendmail(sender_email, receiver_email, text)
# # # #             print(f"Email sent to {receiver_email}")

# # # #         return True  # Email sent successfully
# # # #     except Exception as e:
# # # #         print(f"Failed to send email to {receiver_email}: {str(e)}")
# # # #         return False

# # # # # Function to manage account selection
# # # # def get_account_details(email_sent_count, email_limit):
# # # #     email_accounts = [
# # # #         {'email': 'seafysoftmail@gmail.com', 'password': 'xauubupekdesdwpu'},
# # # #         # {'email': 'email22@gmail.com', 'password': 'password_for_email22'}
# # # #     ]
    
# # # #     # Determine which account to use
# # # #     if email_sent_count < email_limit:
# # # #         return email_accounts[0]['email'], email_accounts[0]['password']
# # # #     else:
# # # #         return email_accounts[1]['email'], email_accounts[1]['password']

# # # # # Path to the Excel file and the lock file
# # # # excel_file_path = '/Users/prudhvirajsomepalli/Documents/Code/vendorEmails.xlsx'
# # # # lock_file_path = f'{excel_file_path}.lock'  # File lock path

# # # # # Email sending logic with locking mechanism
# # # # lock = FileLock(lock_file_path)

# # # # # Run the process with file locking
# # # # with lock:
# # # #     try:
# # # #         # Load the Excel file and check for emails that haven't been sent yet
# # # #         df = pd.read_excel(excel_file_path, sheet_name='Sheet1')

# # # #         # Ensure 'Email Sent Status' column exists
# # # #         if 'Email Sent Status' not in df.columns:
# # # #             df['Email Sent Status'] = 'Not Sent'

# # # #         email_limit = 499
# # # #         email_sent_count = 0

# # # #         # Loop through the email list and send emails
# # # #         for index, row in df.iterrows():
# # # #             if row['Email Sent Status'] == 'Sent':
# # # #                 continue  # Skip if email already sent

# # # #             # Get the appropriate account details based on the sent count
# # # #             sender_email, sender_password = get_account_details(email_sent_count, email_limit)

# # # #             # Send email
# # # #             email_status = send_email(sender_email, row['To Email'], sender_password)

# # # #             if email_status:
# # # #                 # Update the DataFrame if email was successfully sent
# # # #                 df.at[index, 'Email Sent Status'] = 'Sent'
# # # #                 email_sent_count += 1
# # # #             else:
# # # #                 # If email sending fails, break the loop and stop
# # # #                 print("Stopping the email sending process due to failure.")
# # # #                 break

# # # #         # Save the updated Excel file with the email status
# # # #         df.to_excel(excel_file_path, index=False)
# # # #     except Exception as e:
# # # #         print(f"Failed to process the file: {str(e)}")


# # # import pandas as pd
# # # import smtplib
# # # from email.mime.text import MIMEText
# # # from email.mime.multipart import MIMEMultipart
# # # from filelock import FileLock

# # # # Function to generate HTML table from DataFrame
# # # def generate_html_table(df):
# # #     html_table = f"""
# # #         <table style="border-collapse: collapse; width: 100%; border: 2px solid black;">
# # #           <thead style="background-color: lightblue;">
# # #             <tr>
# # #               {''.join([f'<th style="border: 2px solid black; padding: 8px;">{col}</th>' for col in df.columns])}
# # #             </tr>
# # #           </thead>
# # #           <tbody>
# # #         """
# # #     for _, row in df.iterrows():
# # #         html_table += "<tr>" + ''.join([f'<td style="border: 2px solid black; padding: 8px;">{cell}</td>' for cell in row]) + "</tr>"
    
# # #     html_table += """
# # #           </tbody>
# # #         </table>
# # #     """
# # #     return html_table

# # # # Function to send emails
# # # def send_email(sender_email, receiver_email, sender_password, email_body):
# # #     try:
# # #         # Setup the email message
# # #         message = MIMEMultipart()
# # #         message['From'] = sender_email
# # #         message['To'] = receiver_email
# # #         message['Subject'] = 'Seafy Soft HotList'

# # #         # Attach HTML body
# # #         message.attach(MIMEText(email_body, 'html'))

# # #         # Sending email using SMTP
# # #         with smtplib.SMTP('smtp.gmail.com', 587) as server:
# # #             server.starttls()
# # #             server.login(sender_email, sender_password)
# # #             text = message.as_string()
# # #             server.sendmail(sender_email, receiver_email, text)
# # #             print(f"Email sent to {receiver_email}")

# # #         return True  # Email sent successfully
# # #     except Exception as e:
# # #         print(f"Failed to send email to {receiver_email}: {str(e)}")
# # #         return False

# # # # Function to manage account selection
# # # def get_account_details(email_sent_count, email_limit):
# # #     email_accounts = [
# # #         {'email': 'seafysoftmail@gmail.com', 'password': 'xauubupekdesdwpu'},
# # #         # {'email': 'email22@gmail.com', 'password': 'password_for_email22'}
# # #     ]
    
# # #     # Determine which account to use
# # #     if email_sent_count < email_limit:
# # #         return email_accounts[0]['email'], email_accounts[0]['password']
# # #     else:
# # #         return email_accounts[1]['email'], email_accounts[1]['password']

# # # # Paths to the Excel files and the lock file
# # # vendor_excel_file_path = '/Users/prudhvirajsomepalli/Documents/Code/vendorEmails.xlsx'
# # # hotlist_excel_file_path = '/Users/prudhvirajsomepalli/Documents/Code/DailyHotList.xlsx'
# # # lock_file_path = f'{vendor_excel_file_path}.lock'  # File lock path

# # # # Email sending logic with locking mechanism
# # # lock = FileLock(lock_file_path)

# # # # Run the process with file locking
# # # with lock:
# # #     try:
# # #         # Load the hotlist Excel file for the HTML table
# # #         hotlist_df = pd.read_excel(hotlist_excel_file_path, sheet_name='Sheet1')
# # #         email_body = generate_html_table(hotlist_df)

# # #         # Load the vendor email Excel file and check for emails that haven't been sent yet
# # #         vendor_df = pd.read_excel(vendor_excel_file_path, sheet_name='Sheet1')

# # #         # Ensure 'Email Sent Status' column exists
# # #         if 'Email Sent Status' not in vendor_df.columns:
# # #             vendor_df['Email Sent Status'] = 'Not Sent'

# # #         email_limit = 499
# # #         email_sent_count = 0

# # #         # Loop through the vendor email list and send emails
# # #         for index, row in vendor_df.iterrows():
# # #             if row['Email Sent Status'] == 'Sent':
# # #                 continue  # Skip if email already sent

# # #             # Get the appropriate account details based on the sent count
# # #             sender_email, sender_password = get_account_details(email_sent_count, email_limit)

# # #             # Send email
# # #             email_status = send_email(sender_email, row['To Email'], sender_password, email_body)

# # #             if email_status:
# # #                 # Update the DataFrame if email was successfully sent
# # #                 vendor_df.at[index, 'Email Sent Status'] = 'Sent'
# # #                 email_sent_count += 1
# # #             else:
# # #                 # If email sending fails, break the loop and stop
# # #                 print("Stopping the email sending process due to failure.")
# # #                 break

# # #         # Save the updated vendor email file with the email status
# # #         vendor_df.to_excel(vendor_excel_file_path, index=False)
# # #     except Exception as e:
# # #         print(f"Failed to process the file: {str(e)}")


# # # import pandas as pd
# # # import smtplib
# # # from email.mime.text import MIMEText
# # # from email.mime.multipart import MIMEMultipart
# # # from filelock import FileLock
# # # import time

# # # # Function to generate HTML table from DataFrame
# # # def generate_html_table(df):
# # #     html_table = f"""
# # #         <table style="border-collapse: collapse; width: 100%; border: 2px solid black;">
# # #           <thead style="background-color: lightblue;">
# # #             <tr>
# # #               {''.join([f'<th style="border: 2px solid black; padding: 8px;">{col}</th>' for col in df.columns])}
# # #             </tr>
# # #           </thead>
# # #           <tbody>
# # #         """
# # #     for _, row in df.iterrows():
# # #         html_table += "<tr>" + ''.join([f'<td style="border: 2px solid black; padding: 8px;">{cell}</td>' for cell in row]) + "</tr>"
    
# # #     html_table += """
# # #           </tbody>
# # #         </table>
# # #     """
# # #     return html_table

# # # # Function to send emails
# # # def send_email(sender_email, receiver_email, sender_password, email_body):
# # #     try:
# # #         # Setup the email message
# # #         message = MIMEMultipart()
# # #         message['From'] = sender_email
# # #         message['To'] = receiver_email
# # #         message['Subject'] = 'Seafy Soft HotList'

# # #         # HTML email body with table
# # #         body = f"""
# # #         <html>
# # #         <body>
# # #             <p>Hello,</p>
# # #             <p>I hope this message finds you well. At Safesoft Solutions, we have a team of highly skilled consultants ready to start working On-Site/Remote immediately.</p>
# # #             <p>If you have any questions or need further information, please don’t hesitate to reach out to me by email or phone.</p>
# # #             <p style="background-color: yellow; font-size: 30px; padding: 10px;">Reach us at: hr@seafysoft.com</p>
# # #             </br>
# # #             <p>{email_body}</p>
# # #             <p>Best regards,</p>
# # #             <p>Swetha,<br>
# # #                Seafy Soft Solutions <br>
# # #                Email: Hr@seafysoft.com <br>
# # #                Phone: +1 9105576339 <br>
# # #                <a href="https://seafysoft.com/">Seafysoft.com/</a> <br>
# # #                <a href="https://www.linkedin.com/company/seafy-soft-solutions/">LinkedIn: Seafy Soft Solutions</a>
# # #             </p>
# # #         </body>
# # #         </html>
# # #         """

# # #         # Attach HTML body
# # #         message.attach(MIMEText(body, 'html'))

# # #         # Sending email using SMTP
# # #         with smtplib.SMTP('smtp.gmail.com', 587) as server:
# # #             server.starttls()
# # #             server.login(sender_email, sender_password)
# # #             text = message.as_string()
# # #             time.sleep(1)
# # #             server.sendmail(sender_email, receiver_email, text)
# # #             # server.sendmail(sender_email, 'rajsomepa@gmail.com', text)
# # #             print(f"Email sent to {receiver_email}")

# # #         return True  # Email sent successfully
# # #     except Exception as e:
# # #         print(f"Failed to send email to {receiver_email}: {str(e)}")
# # #         return False

# # # # Function to manage account selection
# # # def get_account_details(email_sent_count, email_limit):
# # #     email_accounts = [
# # #         {'email': 'seafysoftmail@gmail.com', 'password': 'xauubupekdesdwpu'},
# # #         # {'email': 'email22@gmail.com', 'password': 'password_for_email22'}
# # #     ]
    
# # #     # Determine which account to use
# # #     if email_sent_count < email_limit:
# # #         return email_accounts[0]['email'], email_accounts[0]['password']
# # #     else:
# # #         return email_accounts[1]['email'], email_accounts[1]['password']

# # # # Paths to the Excel files and the lock file
# # # vendor_excel_file_path = '/Users/prudhvirajsomepalli/Documents/Code/vendorEmails.xlsx'
# # # hotlist_excel_file_path = '/Users/prudhvirajsomepalli/Documents/Code/DailyHotList.xlsx'
# # # lock_file_path = f'{vendor_excel_file_path}.lock'  # File lock path

# # # # Email sending logic with locking mechanism
# # # lock = FileLock(lock_file_path)

# # # # Run the process with file locking
# # # with lock:
# # #     try:
# # #         # Load the hotlist Excel file for the HTML table
# # #         hotlist_df = pd.read_excel(hotlist_excel_file_path, sheet_name='Sheet1')
# # #         email_body = generate_html_table(hotlist_df)

# # #         # Load the vendor email Excel file and check for emails that haven't been sent yet
# # #         vendor_df = pd.read_excel(vendor_excel_file_path, sheet_name='Sheet1')

# # #         # Ensure 'Email Sent Status' column exists
# # #         if 'Email Sent Status' not in vendor_df.columns:
# # #             vendor_df['Email Sent Status'] = 'Not Sent'

# # #         email_limit = 499
# # #         email_sent_count = 0

# # #         # Loop through the vendor email list and send emails
# # #         for index, row in vendor_df.iterrows():
# # #             if row['Email Sent Status'] == 'Sent':
# # #                 continue  # Skip if email already sent

# # #             # Get the appropriate account details based on the sent count
# # #             sender_email, sender_password = get_account_details(email_sent_count, email_limit)

# # #             # Send email
# # #             email_status = send_email(sender_email, row['To Email'], sender_password, email_body)

# # #             if email_status:
# # #                 # Update the DataFrame if email was successfully sent
# # #                 vendor_df.at[index, 'Email Sent Status'] = 'Sent'
# # #                 email_sent_count += 1
# # #             else:
# # #                 # If email sending fails, break the loop and stop
# # #                 print("Stopping the email sending process due to failure.")
# # #                 break

# # #         # Save the updated vendor email file with the email status
# # #         vendor_df.to_excel(vendor_excel_file_path, index=False)
# # #     except Exception as e:
# # #         print(f"Failed to process the file: {str(e)}")


# # import pandas as pd
# # import smtplib
# # from email.mime.text import MIMEText
# # from email.mime.multipart import MIMEMultipart
# # from filelock import FileLock
# # from datetime import datetime  # To get the current date and time
# # import time

# # # Function to generate HTML table from DataFrame
# # def generate_html_table(df):
# #     html_table = f"""
# #         <table style="border-collapse: collapse; width: 100%; border: 2px solid black;">
# #           <thead style="background-color: lightblue;">
# #             <tr>
# #               {''.join([f'<th style="border: 2px solid black; padding: 8px;">{col}</th>' for col in df.columns])}
# #             </tr>
# #           </thead>
# #           <tbody>
# #         """
# #     for _, row in df.iterrows():
# #         html_table += "<tr>" + ''.join([f'<td style="border: 2px solid black; padding: 8px;">{cell}</td>' for cell in row]) + "</tr>"
    
# #     html_table += """
# #           </tbody>
# #         </table>
# #     """
# #     return html_table

# # # Function to send emails
# # def send_email(sender_email, receiver_email, sender_password, email_body):
# #     try:
# #         # Setup the email message
# #         message = MIMEMultipart()
# #         message['From'] = sender_email
# #         message['To'] = receiver_email
# #         message['Subject'] = 'Seafy Soft HotList'

# #         # HTML email body with table
# #         body = f"""
# #         <html>
# #         <body>
# #             <p>Hello,</p>
# #             <p>I hope this message finds you well. At Safesoft Solutions, we have a team of highly skilled consultants ready to start working On-Site/Remote immediately.</p>
# #             <p>If you have any questions or need further information, please don’t hesitate to reach out to me by email or phone.</p>
# #             <p style="background-color: yellow; font-size: 30px; padding: 10px;">Reach us at: hr@seafysoft.com</p>
# #             <br/>
# #             <p>{email_body}</p>
# #             <p>Best regards,</p>
# #             <p>Swetha,<br>
# #                Seafy Soft Solutions <br>
# #                Email: Hr@seafysoft.com <br>
# #                Phone: +1 9105576339 <br>
# #                <a href="https://seafysoft.com/">Seafysoft.com/</a> <br>
# #                <a href="https://www.linkedin.com/company/seafy-soft-solutions/">LinkedIn: Seafy Soft Solutions</a>
# #             </p>
# #         </body>
# #         </html>
# #         """

# #         # Attach HTML body
# #         message.attach(MIMEText(body, 'html'))

# #         # Sending email using SMTP
# #         with smtplib.SMTP('smtp.gmail.com', 587) as server:
# #             server.starttls()
# #             server.login(sender_email, sender_password)
# #             text = message.as_string()
# #             time.sleep(1)
# #             server.sendmail(sender_email, receiver_email, text)
# #             print(f"Email sent to {receiver_email}")

# #         return True  # Email sent successfully
# #     except Exception as e:
# #         print(f"Failed to send email to {receiver_email}: {str(e)}")
# #         return False

# # # Function to manage account selection
# # def get_account_details(email_sent_count, email_limit):
# #     email_accounts = [
# #         {'email': 'seafysoftmail@gmail.com', 'password': 'xauubupekdesdwpu'},
# #         # {'email': 'email22@gmail.com', 'password': 'password_for_email22'}
# #     ]
    
# #     # Determine which account to use
# #     if email_sent_count < email_limit:
# #         return email_accounts[0]['email'], email_accounts[0]['password']
# #     else:
# #         return email_accounts[1]['email'], email_accounts[1]['password']

# # # Paths to the Excel files and the lock file
# # vendor_excel_file_path = '/Users/prudhvirajsomepalli/Documents/Code/vendorEmails.xlsx'
# # hotlist_excel_file_path = '/Users/prudhvirajsomepalli/Documents/Code/DailyHotList.xlsx'
# # lock_file_path = f'{vendor_excel_file_path}.lock'  # File lock path

# # # Email sending logic with locking mechanism
# # lock = FileLock(lock_file_path)

# # # Run the process with file locking
# # with lock:
# #     try:
# #         # Load the hotlist Excel file for the HTML table
# #         hotlist_df = pd.read_excel(hotlist_excel_file_path, sheet_name='Sheet1')
# #         email_body = generate_html_table(hotlist_df)

# #         # Load the vendor email Excel file and check for emails that haven't been sent yet
# #         vendor_df = pd.read_excel(vendor_excel_file_path, sheet_name='Sheet1')

# #         # Ensure 'Email Sent Status' and 'Email Sent Date' columns exist
# #         if 'Email Sent Status' not in vendor_df.columns:
# #             vendor_df['Email Sent Status'] = 'Not Sent'
# #         if 'Email Sent Date' not in vendor_df.columns:
# #             vendor_df['Email Sent Date'] = ''

# #         email_limit = 499
# #         email_sent_count = 0

# #         # Loop through the vendor email list and send emails
# #         for index, row in vendor_df.iterrows():
# #             if row['Email Sent Status'] == 'Sent':
# #                 continue  # Skip if email already sent

# #             # Get the appropriate account details based on the sent count
# #             sender_email, sender_password = get_account_details(email_sent_count, email_limit)

# #             # Send email
# #             email_status = send_email(sender_email, row['To Email'], sender_password, email_body)

# #             if email_status:
# #                 # Update the DataFrame if email was successfully sent
# #                 vendor_df.at[index, 'Email Sent Status'] = 'Sent'
# #                 vendor_df.at[index, 'Email Sent Date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
# #                 email_sent_count += 1
# #             else:
# #                 # If email sending fails, break the loop and stop
# #                 print("Stopping the email sending process due to failure.")
# #                 break

# #         # Save the updated vendor email file with the email status and date
# #         vendor_df.to_excel(vendor_excel_file_path, index=False)
# #     except Exception as e:
# #         print(f"Failed to process the file: {str(e)}")


# import pandas as pd
# import smtplib
# from email.mime.text import MIMEText
# from email.mime.multipart import MIMEMultipart
# from filelock import FileLock
# from datetime import datetime  # To get the current date and time
# import time

# # Function to generate HTML table from DataFrame
# def generate_html_table(df):
#     html_table = f"""
#         <table style="border-collapse: collapse; width: 100%; border: 2px solid black;">
#           <thead style="background-color: lightblue;">
#             <tr>
#               {''.join([f'<th style="border: 2px solid black; padding: 8px;">{col}</th>' for col in df.columns])}
#             </tr>
#           </thead>
#           <tbody>
#         """
#     for _, row in df.iterrows():
#         html_table += "<tr>" + ''.join([f'<td style="border: 2px solid black; padding: 8px;">{cell}</td>' for cell in row]) + "</tr>"
    
#     html_table += """
#           </tbody>
#         </table>
#     """
#     return html_table

# # Function to send emails
# def send_email(sender_email, receiver_email, sender_password, email_body):
#     try:
#         # Setup the email message
#         message = MIMEMultipart()
#         message['From'] = sender_email
#         message['To'] = receiver_email
#         message['Subject'] = 'Seafy Soft HotList'

#         # HTML email body with table
#         body = f"""
#         <html>
#         <body>
#             <p>Hello,</p>
#             <p>I hope this message finds you well. At Safesoft Solutions, we have a team of highly skilled consultants ready to start working On-Site/Remote immediately.</p>
#             <p>If you have any questions or need further information, please don’t hesitate to reach out to me by email or phone.</p>
#             <p style="background-color: yellow; font-size: 30px; padding: 10px;">Reach us at: hr@seafysoft.com</p>
#             <br/>
#             <p>{email_body}</p>
#             <p>Best regards,</p>
#             <p>Swetha,<br>
#                Seafy Soft Solutions <br>
#                Email: Hr@seafysoft.com <br>
#                Phone: +1 9105576339 <br>
#                <a href="https://seafysoft.com/">Seafysoft.com/</a> <br>
#                <a href="https://www.linkedin.com/company/seafy-soft-solutions/">LinkedIn: Seafy Soft Solutions</a>
#             </p>
#         </body>
#         </html>
#         """

#         # Attach HTML body
#         message.attach(MIMEText(body, 'html'))

#         # Sending email using SMTP
#         with smtplib.SMTP('smtp.gmail.com', 587) as server:
#             server.starttls()
#             server.login(sender_email, sender_password)
#             text = message.as_string()
#             time.sleep(1)
#             server.sendmail(sender_email, receiver_email, text)
#             print(f"Email sent to {receiver_email}")

#         return True  # Email sent successfully
#     except Exception as e:
#         print(f"Failed to send email to {receiver_email}: {str(e)}")
#         return False

# # Function to manage account selection
# def get_account_details(email_sent_count, email_limit):
#     email_accounts = [
#         {'email': 'seafysoftmail@gmail.com', 'password': 'xauubupekdesdwpu'},
#         # {'email': 'email22@gmail.com', 'password': 'password_for_email22'}
#     ]
    
#     # Determine which account to use
#     if email_sent_count < email_limit:
#         return email_accounts[0]['email'], email_accounts[0]['password']
#     else:
#         return email_accounts[1]['email'], email_accounts[1]['password']

# # Paths to the Excel files and the lock file
# vendor_excel_file_path = '/Users/prudhvirajsomepalli/Documents/Code/vendorEmails.xlsx'
# hotlist_excel_file_path = '/Users/prudhvirajsomepalli/Documents/Code/DailyHotList.xlsx'
# lock_file_path = f'{vendor_excel_file_path}.lock'  # File lock path

# # Email sending logic with locking mechanism
# lock = FileLock(lock_file_path)

# # Run the process with file locking
# with lock:
#     try:
#         # Load the hotlist Excel file for the HTML table
#         hotlist_df = pd.read_excel(hotlist_excel_file_path, sheet_name='Sheet1')
#         email_body = generate_html_table(hotlist_df)

#         # Load the vendor email Excel file and check for emails that haven't been sent yet
#         vendor_df = pd.read_excel(vendor_excel_file_path, sheet_name='Sheet1')

#         # Ensure 'Email Sent Status' and 'Email Sent Date' columns exist
#         if 'Email Sent Status' not in vendor_df.columns:
#             vendor_df['Email Sent Status'] = 'Not Sent'
#         if 'Email Sent Date' not in vendor_df.columns:
#             vendor_df['Email Sent Date'] = ''

#         email_limit = 499
#         email_sent_count = 0
#         today = datetime.now().strftime('%Y-%m-%d')  # Get today's date in YYYY-MM-DD format
#         emails_sent_today = True  # Flag to check if all emails were sent today

#         # Loop through the vendor email list and send emails
#         for index, row in vendor_df.iterrows():
#             email_sent_date = str(row['Email Sent Date'])[:10]  # Extract the date part

#             # Check if the email was sent today
#             if email_sent_date == today:
#                 continue  # Skip if the email was already sent today

#             emails_sent_today = False  # If at least one email is sent today, set the flag to False

#             # Get the appropriate account details based on the sent count
#             sender_email, sender_password = get_account_details(email_sent_count, email_limit)

#             # Send email
#             email_status = send_email(sender_email, row['To Email'], sender_password, email_body)

#             if email_status:
#                 # Update the DataFrame if email was successfully sent
#                 vendor_df.at[index, 'Email Sent Status'] = 'Sent'
#                 vendor_df.at[index, 'Email Sent Date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
#                 email_sent_count += 1
#             else:
#                 # If email sending fails, break the loop and stop
#                 print("Stopping the email sending process due to failure.")
#                 break

#         # Reset all 'Email Sent Status' to 'Not Sent' for the next day if all emails were sent today
#         if emails_sent_today:
#             vendor_df['Email Sent Status'] = 'Not Sent'

#         # Save the updated vendor email file with the email status and date
#         vendor_df.to_excel(vendor_excel_file_path, index=False)
#     except Exception as e:
#         print(f"Failed to process the file: {str(e)}")


import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from filelock import FileLock
from datetime import datetime
import time

# Function to generate HTML table from DataFrame
def generate_html_table(df):
    html_table = f"""
        <table style="border-collapse: collapse; width: 100%; border: 2px solid black;">
          <thead style="background-color: lightblue;">
            <tr>
              {''.join([f'<th style="border: 2px solid black; padding: 8px;">{col}</th>' for col in df.columns])}
            </tr>
          </thead>
          <tbody>
        """
    for _, row in df.iterrows():
        html_table += "<tr>" + ''.join([f'<td style="border: 2px solid black; padding: 8px;">{cell}</td>' for cell in row]) + "</tr>"
    
    html_table += """
          </tbody>
        </table>
    """
    return html_table

# Function to send emails
def send_email(sender_email, receiver_email, sender_password, email_body):
    try:
        if pd.isna(receiver_email) or not isinstance(receiver_email, str):
            raise ValueError("Invalid email address")
        
        # Setup the email message
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = receiver_email
        message['Subject'] = 'Available candidates from SeafySoft'

        # HTML email body with table
        body = f"""
        <html>
        <body>
            <h3>Hello,</H3>
            <p>At <strong>Safesoft Solutions</strong>, we take pride in our team of highly skilled consultants, who are ready to begin working on-site or remotely at your convenience. Whether you have immediate needs or future projects in mind, we are here to support you..</p>   <p>Should you have any questions or require further information, please feel free to contact me via email or phone. I look forward to the opportunity to assist you.</p>
             <h1><span style="background-color: yellow;  padding: 0 5px;">Reach us at: hr@seafysoft.com</span></h1>
            <br/>
            <p>{email_body}</p>
            <p>Best regards,</p>
            <h3>Swetha,</h3>
               <h3>Seafy Soft Solutions</h3>
               Email: Hr@seafysoft.com <br>
               Phone: +1 9105576339 <br>
               <a href="https://seafysoft.com/">Seafysoft.com/</a> <br>
               LinkedIn: <a href="https://www.linkedin.com/company/seafy-soft-solutions/">Seafy Soft Solutions</a>
            </p>
        </body>
        </html>
        """

        # Attach HTML body
        message.attach(MIMEText(body, 'html'))

        # Sending email using SMTP
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            text = message.as_string()
            time.sleep(2)
            server.sendmail(sender_email, receiver_email, text)
            print(f"Email sent to {receiver_email}")

        return True  # Email sent successfully
    except Exception as e:
        print(f"Failed to send email to {receiver_email}: {str(e)}")
        return False

# Function to manage account selection
def get_account_details(email_sent_count, email_limit):
    email_accounts = [
        {'email': 'dailyrequriments@gmail.com', 'password': 'yulyeentyuykfpav'}
    ]
    
    # Determine which account to use
    if email_sent_count < email_limit:
        return email_accounts[0]['email'], email_accounts[0]['password']
    else:
        return email_accounts[1]['email'], email_accounts[1]['password']

# Paths to the Excel files and the lock file

vendor_excel_file_path = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\GroupMails.xlsx'
hotlist_excel_file_path = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\DailyHotList.xlsx'

lock_file_path = f'{vendor_excel_file_path}.lock'  # File lock path

# Email sending logic with locking mechanism
lock = FileLock(lock_file_path)

# Run the process with file locking
with lock:
    try:
        # Load the hotlist Excel file for the HTML table
        hotlist_df = pd.read_excel(hotlist_excel_file_path, sheet_name='Sheet1')
        email_body = generate_html_table(hotlist_df)

        # Load the vendor email Excel file and check for emails that haven't been sent yet
        vendor_df = pd.read_excel(vendor_excel_file_path, sheet_name='Sheet1')

        # Ensure 'Email Sent Status' and 'Email Sent Date' columns exist
        if 'Email Sent Status' not in vendor_df.columns:
            vendor_df['Email Sent Status'] = 'Not Sent'
        if 'Email Sent Date' not in vendor_df.columns:
            vendor_df['Email Sent Date'] = ''

        email_limit = 499
        
        email_sent_count = 0
        today = datetime.now().strftime('%Y-%m-%d')  # Get today's date in YYYY-MM-DD format
        emails_sent_today = True  # Flag to check if all emails were sent today

        # Loop through the vendor email list and send emails
        for index, row in vendor_df.iterrows():
            email_sent_date = str(row['Email Sent Date'])[:10]  # Extract the date part

            # Check if the email was sent today
            # if email_sent_date == today:
            #     continue  # Skip if the email was already sent today

            receiver_email = row['To Email']

            # Check if email is valid
            if pd.isna(receiver_email) or not isinstance(receiver_email, str):
                print(f"Skipping invalid email address: {receiver_email}")
                continue  # Skip if the email address is invalid

            emails_sent_today = False  # If at least one email is sent today, set the flag to False

            # Get the appropriate account details based on the sent count
            sender_email, sender_password = get_account_details(email_sent_count, email_limit)

            # Send email
            email_status = send_email(sender_email, receiver_email, sender_password, email_body)

            if email_status:
                # Update the DataFrame if email was successfully sent
                vendor_df.at[index, 'Email Sent Status'] = 'Sent'
                vendor_df.at[index, 'Email Sent Date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                email_sent_count += 1
            else:
                # If email sending fails, break the loop and stop
                print("Stopping the email sending process due to failure.")
                break

        # Reset all 'Email Sent Status' to 'Not Sent' for the next day if all emails were sent today
        if emails_sent_today:
            vendor_df['Email Sent Status'] = 'Not Sent'

        # Save the updated vendor email file with the email status and date
        vendor_df.to_excel(vendor_excel_file_path, index=False)
    except Exception as e:
        print(f"Failed to process the file: {str(e)}")

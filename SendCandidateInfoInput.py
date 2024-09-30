
# import imaplib
# import email
# import smtplib
# from email.mime.text import MIMEText
# from email.mime.multipart import MIMEMultipart
# import pandas as pd
# import re
# import time
# from datetime import datetime
# import os

# # Login information
# SourceUsername = 'dailyrequriments@gmail.com'
# SourcePassword = 'yulyeentyuykfpav'
# imap_url = 'imap.gmail.com'
# smtp_server = 'smtp.gmail.com'
# smtp_port = 587

# # Load the Excel files
# daily_hotlist_path = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\DailyHotList.xlsx'
# today_requirements_path = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\TodayRequriments.xlsx'

# try:
#     send_candidate_df = pd.read_excel(daily_hotlist_path, sheet_name='Sheet1')
#     print("Loaded daily hotlist successfully.")
# except Exception as e:
#     print(f"Error loading daily hotlist: {e}")

# try:
#     keywords_df = pd.read_excel(daily_hotlist_path, sheet_name='Keywords')
#     print("Loaded keywords successfully.")
# except Exception as e:
#     print(f"Error loading keywords from Excel: {e}")

# # Gmail labels
# label_to_apply = 'Processed'
# restricted_keywords = ['phmailadmin', '.email']  # Add more as needed

# def check_new_emails():
#     mail = None
#     try:
#         mail = imaplib.IMAP4_SSL(imap_url)
#         mail.login(SourceUsername, SourcePassword)
#         mail.select('inbox')
#         print("Logged in to the email account and selected inbox.")
#     except Exception as e:
#         print(f"Error logging in or selecting inbox: {e}")
#         return

#     while True:
#         try:
#             status, response = mail.search(None, 'UNSEEN')
#             if status != 'OK':
#                 print(f"Failed to search for unseen emails: {response}")
#                 time.sleep(60)
#                 continue

#             email_ids = response[0].split()
#             print(f"Fetched {len(email_ids)} new unseen emails.")

#             if email_ids:
#                 for email_id in email_ids:
#                     try:
#                         process_email(mail, email_id)
#                     except Exception as e:
#                         print(f"Error processing email ID {email_id}: {e}")

#             else:
#                 print("No new emails found. Waiting for the next check.")

#             time.sleep(60)  # Wait before checking for new emails again

#         except Exception as e:
#             print(f"Error during email check: {e}")
#             time.sleep(60)

#     try:
#         mail.logout()
#         print("Logged out from the email account.")
#     except Exception as e:
#         print(f"Error logging out: {e}")

# def process_email(mail, email_id):
#     try:
#         status, msg_data = mail.fetch(email_id, '(RFC822)')
#         for response_part in msg_data:
#             if isinstance(response_part, tuple):
#                 msg = email.message_from_bytes(response_part[1])
#                 subject = msg['Subject']
#                 print(f"Processing email ID {email_id} with subject: {subject}")

#                 # Extract email body
#                 email_body = extract_body(msg)
#                 if email_body:
#                     # Send reply if subject matches keywords
#                     matching_candidates = check_keywords_in_subject(subject)
#                     if matching_candidates:
#                         send_reply(msg, subject, matching_candidates)
#                         update_today_requirements(email_body, subject, msg)

#                 apply_label(mail, email_id, label_to_apply)

#     except Exception as e:
#         print(f"Error fetching or processing email ID {email_id}: {e}")

# def apply_label(mail, email_id, label):
#     try:
#         mail.store(email_id, '+X-GM-LABELS', label)
#         print(f"Applied label '{label}' to email ID {email_id}.")
#     except Exception as e:
#         print(f"Error applying label to email ID {email_id}: {e}")

# def extract_body(msg):
#     try:
#         body = ""
#         if msg.is_multipart():
#             for part in msg.walk():
#                 content_type = part.get_content_type()
#                 content_disposition = str(part.get("Content-Disposition"))

#                 if content_type == "text/plain" and 'attachment' not in content_disposition:
#                     body += part.get_payload(decode=True).decode(part.get_content_charset(), errors='ignore')
#                 elif content_type == "text/html" and 'attachment' not in content_disposition:
#                     body += part.get_payload(decode=True).decode(part.get_content_charset(), errors='ignore')
#         else:
#             body = msg.get_payload(decode=True).decode(msg.get_content_charset(), errors='ignore')

#         print("Email body extracted successfully.")
#         return body
#     except Exception as e:
#         print(f"Error extracting email body: {e}")
#         return ""

# def extract_valid_emails(email_message):
#     email_list = [email_message['Reply-To'], email_message['From']]
#     body = extract_body(email_message)

#     email_addresses = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', body)
#     email_list.extend(email_addresses)

#     email_list = list(set(email_list))
#     valid_emails = [item for item in email_list if item and not any(restricted in item for restricted in restricted_keywords)]
    
#     return valid_emails

# def send_reply(msg, subject, matching_candidates):
#     server = None
#     try:
#         server = smtplib.SMTP(smtp_server, smtp_port)
#         server.starttls()
#         server.login(SourceUsername, SourcePassword)
#         print("Logged in to SMTP server.")
#     except smtplib.SMTPAuthenticationError as auth_error:
#         print(f"Authentication error occurred: {auth_error}")
#         return
#     except Exception as e:
#         print(f"Error connecting to SMTP server: {e}")
#         return

#     body = extract_body(msg)

#     try:
#         reply_to_email = msg['Reply-To'] if msg['Reply-To'] else msg['From']
#         print(f"Reply-to email: {reply_to_email}, Valid candidates: {matching_candidates}")

#         if not any(restricted in reply_to_email for restricted in restricted_keywords):
#             msg_out = create_reply_message(matching_candidates, reply_to_email, subject, body)
#             server.send_message(msg_out)
#             print(f'Sent reply to: {reply_to_email}, Subject: {msg_out["Subject"]}')
#     except Exception as e:
#         print(f"Error sending reply: {e}")
#     finally:
#         if server:
#             server.quit()
#             print("SMTP connection closed.")

# def create_reply_message(matching_candidates, reply_to_email, subject, email_body):
#     msg_out = MIMEMultipart()
#     msg_out['From'] = SourceUsername
#     msg_out['To'] = reply_to_email
#     msg_out['Subject'] = 'Re: ' + subject

#     body_html = create_html_table(matching_candidates)
#     msg_out.attach(MIMEText(body_html, 'html'))
#     print("Reply message created.")

#     return msg_out

# def update_today_requirements(email_body, subject, msg):
#     try:
#         phone_pattern = r'(?<!\d)(\+?\d{1,3}[-.\s]?)?(\(?\d{1,4}\)?[-.\s]?)?(\d{1,4})[-.\s]?(\d{1,4})[-.\s]?(\d{1,9})(?!\d)'
#         phone_matches = re.findall(phone_pattern, email_body)

#         today_date = datetime.now().strftime('%Y-%m-%d')
#         technology = subject  # Replace with logic to determine technology based on subject

#         if phone_matches:
#             vendor_emails = []
#             reply_to_email = msg['Reply-To'] if msg['Reply-To'] else msg['From']
#             valid_phones = []

#             for match in phone_matches:
#                 full_phone = ''.join(part for part in match if part).strip()
#                 if len(re.sub(r'\D', '', full_phone)) >= 10:
#                     valid_phones.append(full_phone)
#                     vendor_emails.append(reply_to_email)

#             if vendor_emails and valid_phones:
#                 vendor_emails_str = ', '.join(set(vendor_emails))  # Deduplicate emails if needed
#                 log_entry = {
#                     'VendorEmail': vendor_emails_str,
#                     'VendorPhone': ', '.join(set(valid_phones)),
#                     'Date': today_date,
#                     'Technology': technology
#                 }

#                 update_excel(today_requirements_path, log_entry)

#     except Exception as e:
#         print(f"Error updating TodayRequirements: {e}")

# def update_excel(file_path, log_entry):
#     try:
#         if os.path.exists(file_path):
#             df_today = pd.read_excel(file_path)
#         else:
#             df_today = pd.DataFrame(columns=['VendorEmail', 'VendorPhone', 'Date', 'Technology'])

#         # Create a DataFrame for the new entry
#         new_entry_df = pd.DataFrame([log_entry])

#         # Use pd.concat to append the new entry
#         df_today = pd.concat([df_today, new_entry_df], ignore_index=True)

#         # Write back to Excel
#         df_today.to_excel(file_path, index=False)
#         print(f"Updated {file_path} successfully with new entry.")
#     except Exception as e:
#         print(f"Error updating Excel file: {e}")

# def check_keywords_in_subject(subject):
#     matching_candidates = []
    
#     all_keywords = keywords_df['Keywords'].str.cat(sep=',').split(',')
#     all_keywords = [keyword.strip().lower() for keyword in all_keywords]
    
#     matching_keywords = [keyword for keyword in all_keywords if keyword in subject.strip().lower()]
#     print(f"Matching keywords in subject: {matching_keywords}")

#     if matching_keywords:
#         for _, row in send_candidate_df.iterrows():
#             candidate_job_title = row['Title'].strip().lower()
#             candidate_name = row['Name']
#             candidate_experience = row['Experience']
#             candidate_location = row['Location']
#             candidate_status = row['Status']
            
#             if any(keyword in candidate_job_title for keyword in matching_keywords):
#                 matching_candidates.append({
#                     'Name': candidate_name,
#                     'JobTitle': row['Title'],
#                     'Experience': candidate_experience,
#                     'Location': candidate_location,
#                     'Status': candidate_status
#                 })
#                 print(f"Matching candidate found: {candidate_name}")

#     return matching_candidates

# def create_html_table(candidates):
#     html = '''<html>
#         <body>
#             <h3>Hello,</h3>
#             <p>At <strong>Seafy Soft Solutions</strong>, we take pride in our team of highly skilled consultants, who are ready to begin working on-site or remotely at your convenience. Whether you have immediate needs or future projects in mind, we are here to support you.</p>
#             <p>Should you have any questions or require further information, please feel free to contact me via email or phone. I look forward to the opportunity to assist you.</p>
#             <h1><span style="background-color: yellow; padding: 0 5px;">Reach us at: hr@seafysoft.com</span></h1>
#             <br/>
#             <p>Here are the candidates that match your requirements:</p>
#             <table style="border-collapse: collapse; width: 100%;">
#                 <tr>
#                     <th style="border: 1px solid black; padding: 8px;">Name</th>
#                     <th style="border: 1px solid black; padding: 8px;">Job Title</th>
#                     <th style="border: 1px solid black; padding: 8px;">Experience</th>
#                     <th style="border: 1px solid black; padding: 8px;">Location</th>
#                     <th style="border: 1px solid black; padding: 8px;">Status</th>
#                 </tr>'''

#     for candidate in candidates:
#         html += f'''
#                 <tr>
#                     <td style="border: 1px solid black; padding: 8px;">{candidate["Name"]}</td>
#                     <td style="border: 1px solid black; padding: 8px;">{candidate["JobTitle"]}</td>
#                     <td style="border: 1px solid black; padding: 8px;">{candidate["Experience"]}</td>
#                     <td style="border: 1px solid black; padding: 8px;">{candidate["Location"]}</td>
#                     <td style="border: 1px solid black; padding: 8px;">{candidate["Status"]}</td>
#                 </tr>'''

#     html += '''
#             </table>
#             <p>Best regards,</p>
#             <h3>Swetha,</h3>
#             <h3>Seafy Soft Solutions</h3>
#             Email: hr@seafysoft.com <br>
#             Phone: +1 9105576339 <br>
#             <a href="https://seafysoft.com/">Seafysoft.com/</a> <br>
#             LinkedIn: <a href="https://www.linkedin.com/company/seafy-soft-solutions/">Seafy Soft Solutions</a>
#         </body>
#     </html>'''
    
#     return html

# if __name__ == "__main__":
#     check_new_emails()

#Above code works 

# import imaplib
# import email
# import smtplib
# from email.mime.text import MIMEText
# from email.mime.multipart import MIMEMultipart
# import pandas as pd
# import re
# import time
# from datetime import datetime
# import os

# # Login information
# SourceUsername = 'dailyrequriments@gmail.com'
# SourcePassword = 'yulyeentyuykfpav'
# imap_url = 'imap.gmail.com'
# smtp_server = 'smtp.gmail.com'
# smtp_port = 587

# # Load the Excel files
# daily_hotlist_path = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\DailyHotList.xlsx'
# today_requirements_path = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\TodayRequriments.xlsx'

# try:
#     send_candidate_df = pd.read_excel(daily_hotlist_path, sheet_name='Sheet1')
#     print("Loaded daily hotlist successfully.")
# except Exception as e:
#     print(f"Error loading daily hotlist: {e}")

# try:
#     keywords_df = pd.read_excel(daily_hotlist_path, sheet_name='Keywords')
#     print("Loaded keywords successfully.")
# except Exception as e:
#     print(f"Error loading keywords from Excel: {e}")

# # Gmail labels
# label_to_apply = 'Processed'
# restricted_keywords = ['phmailadmin', '.email']  # Add more as needed

# def check_new_emails():
#     connected = False
#     while True:
#         mail = None
#         try:
#             if not connected:
#                 mail = imaplib.IMAP4_SSL(imap_url)
#                 mail.login(SourceUsername, SourcePassword)
#                 mail.select('inbox')
#                 print("Logged in to the email account and selected inbox.")
#                 connected = True

#             while True:
#                 try:
#                     status, response = mail.search(None, 'UNSEEN')
#                     if status != 'OK':
#                         print(f"Failed to search for unseen emails: {response}")
#                         break  # Exit inner loop to reconnect

#                     email_ids = response[0].split()
#                     print(f"Fetched {len(email_ids)} new unseen emails.")

#                     if email_ids:
#                         for email_id in email_ids:
#                             try:
#                                 process_email(mail, email_id)
#                             except Exception as e:
#                                 print(f"Error processing email ID {email_id}: {e}")

#                     else:
#                         print("No new emails found. Waiting for the next check.")

#                     time.sleep(60)  # Wait before checking for new emails again

#                 except Exception as e:
#                     print(f"Error during email check: {e}")
#                     break  # Exit inner loop to reconnect

#         except Exception as e:
#             print(f"Error logging in or selecting inbox: {e}")
#             connected = False  # Set to reconnect next time
#             time.sleep(60)  # Wait before trying to reconnect

#         finally:
#             if mail:
#                 try:
#                     mail.logout()
#                     print("Logged out from the email account.")
#                 except Exception as e:
#                     print(f"Error logging out: {e}")

#             if not connected:
#                 print("Reconnecting in 60 seconds...")
#                 time.sleep(60)  # Wait before reconnecting

# def process_email(mail, email_id):
#     try:
#         status, msg_data = mail.fetch(email_id, '(RFC822)')
#         for response_part in msg_data:
#             if isinstance(response_part, tuple):
#                 msg = email.message_from_bytes(response_part[1])
#                 subject = msg['Subject']
#                 print(f"Processing email ID {email_id} with subject: {subject}")

#                 # Extract email body
#                 email_body = extract_body(msg)
#                 if email_body:
#                     # Send reply if subject matches keywords
#                     matching_candidates = check_keywords_in_subject(subject)
#                     if matching_candidates:
#                         send_reply(msg, subject, matching_candidates)
#                         update_today_requirements(email_body, subject, msg)

#                 apply_label(mail, email_id, label_to_apply)

#     except Exception as e:
#         print(f"Error fetching or processing email ID {email_id}: {e}")

# def apply_label(mail, email_id, label):
#     try:
#         mail.store(email_id, '+X-GM-LABELS', label)
#         print(f"Applied label '{label}' to email ID {email_id}.")
#     except Exception as e:
#         print(f"Error applying label to email ID {email_id}: {e}")

# def extract_body(msg):
#     try:
#         body = ""
#         if msg.is_multipart():
#             for part in msg.walk():
#                 content_type = part.get_content_type()
#                 content_disposition = str(part.get("Content-Disposition"))

#                 if content_type == "text/plain" and 'attachment' not in content_disposition:
#                     body += part.get_payload(decode=True).decode(part.get_content_charset(), errors='ignore')
#                 elif content_type == "text/html" and 'attachment' not in content_disposition:
#                     body += part.get_payload(decode=True).decode(part.get_content_charset(), errors='ignore')
#         else:
#             body = msg.get_payload(decode=True).decode(msg.get_content_charset(), errors='ignore')

#         print("Email body extracted successfully.")
#         return body
#     except Exception as e:
#         print(f"Error extracting email body: {e}")
#         return ""

# def extract_valid_emails(email_message):
#     email_list = [email_message['Reply-To'], email_message['From']]
#     body = extract_body(email_message)

#     email_addresses = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', body)
#     email_list.extend(email_addresses)

#     email_list = list(set(email_list))
#     valid_emails = [item for item in email_list if item and not any(restricted in item for restricted in restricted_keywords)]
    
#     return valid_emails

# def send_reply(msg, subject, matching_candidates):
#     server = None
#     try:
#         server = smtplib.SMTP(smtp_server, smtp_port)
#         server.starttls()
#         server.login(SourceUsername, SourcePassword)
#         print("Logged in to SMTP server.")
#     except smtplib.SMTPAuthenticationError as auth_error:
#         print(f"Authentication error occurred: {auth_error}")
#         return
#     except Exception as e:
#         print(f"Error connecting to SMTP server: {e}")
#         return

#     body = extract_body(msg)

#     try:
#         reply_to_email = msg['Reply-To'] if msg['Reply-To'] else msg['From']
#         print(f"Reply-to email: {reply_to_email}, Valid candidates: {matching_candidates}")

#         if not any(restricted in reply_to_email for restricted in restricted_keywords):
#             msg_out = create_reply_message(matching_candidates, reply_to_email, subject, body)
#             server.send_message(msg_out)
#             print(f'Sent reply to: {reply_to_email}, Subject: {msg_out["Subject"]}')
#     except Exception as e:
#         print(f"Error sending reply: {e}")
#     finally:
#         if server:
#             server.quit()
#             print("SMTP connection closed.")

# def create_reply_message(matching_candidates, reply_to_email, subject, email_body):
#     msg_out = MIMEMultipart()
#     msg_out['From'] = SourceUsername
#     msg_out['To'] = reply_to_email
#     msg_out['Subject'] = 'Re: ' + subject

#     body_html = create_html_table(matching_candidates)
#     msg_out.attach(MIMEText(body_html, 'html'))
#     print("Reply message created.")

#     return msg_out

# def update_today_requirements(email_body, subject, msg):
#     try:
#         phone_pattern = r'(?<!\d)(\+?\d{1,3}[-.\s]?)?(\(?\d{1,4}\)?[-.\s]?)?(\d{1,4})[-.\s]?(\d{1,4})[-.\s]?(\d{1,9})(?!\d)'
#         phone_matches = re.findall(phone_pattern, email_body)

#         today_date = datetime.now().strftime('%Y-%m-%d')
#         technology = subject  # Replace with logic to determine technology based on subject

#         if phone_matches:
#             vendor_emails = []
#             reply_to_email = msg['Reply-To'] if msg['Reply-To'] else msg['From']
#             valid_phones = []

#             for match in phone_matches:
#                 full_phone = ''.join(part for part in match if part).strip()
#                 if len(re.sub(r'\D', '', full_phone)) >= 10:
#                     valid_phones.append(full_phone)
#                     vendor_emails.append(reply_to_email)

#             if vendor_emails and valid_phones:
#                 vendor_emails_str = ', '.join(set(vendor_emails))  # Deduplicate emails if needed
#                 log_entry = {
#                     'VendorEmail': vendor_emails_str,
#                     'VendorPhone': ', '.join(set(valid_phones)),
#                     'Date': today_date,
#                     'Technology': technology
#                 }

#                 update_excel(today_requirements_path, log_entry)

#     except Exception as e:
#         print(f"Error updating TodayRequirements: {e}")

# def update_excel(file_path, log_entry):
#     try:
#         if os.path.exists(file_path):
#             df_today = pd.read_excel(file_path)
#         else:
#             df_today = pd.DataFrame(columns=['VendorEmail', 'VendorPhone', 'Date', 'Technology'])

#         # Create a DataFrame for the new entry
#         new_entry_df = pd.DataFrame([log_entry])

#         # Use pd.concat to append the new entry
#         df_today = pd.concat([df_today, new_entry_df], ignore_index=True)

#         # Write back to Excel
#         df_today.to_excel(file_path, index=False)
#         print(f"Updated {file_path} successfully with new entry.")
#     except Exception as e:
#         print(f"Error updating Excel file: {e}")

# def check_keywords_in_subject(subject):
#     matching_candidates = []
    
#     all_keywords = keywords_df['Keywords'].str.cat(sep=',').split(',')
#     all_keywords = [keyword.strip().lower() for keyword in all_keywords]
    
#     matching_keywords = [keyword for keyword in all_keywords if keyword in subject.strip().lower()]
#     print(f"Matching keywords in subject: {matching_keywords}")

#     if matching_keywords:
#         for _, row in send_candidate_df.iterrows():
#             candidate_job_title = row['Title'].strip().lower()
#             candidate_name = row['Name']
#             candidate_experience = row['Experience']
#             candidate_location = row['Location']
#             candidate_status = row['Status']
            
#             if any(keyword in candidate_job_title for keyword in matching_keywords):
#                 matching_candidates.append({
#                     'Name': candidate_name,
#                     'JobTitle': row['Title'],
#                     'Experience': candidate_experience,
#                     'Location': candidate_location,
#                     'Status': candidate_status
#                 })
#                 print(f"Matching candidate found: {candidate_name}")

#     return matching_candidates

# def create_html_table(candidates):
#     html = '''<html>
#         <body>
#             <h3>Hello,</h3>
#             <p>At <strong>Seafy Soft Solutions</strong>, we take pride in our team of highly skilled consultants, who are ready to begin working on-site or remotely at your convenience. Whether you have immediate needs or future projects in mind, we are here to support you.</p>
#             <p>Should you have any questions or require further information, please feel free to contact me via email or phone. I look forward to the opportunity to assist you.</p>
#             <h1><span style="background-color: yellow; padding: 0 5px;">Reach us at: hr@seafysoft.com</span></h1>
#             <br/>
#             <p>Here are the candidates that match your requirements:</p>
#             <table style="border-collapse: collapse; width: 100%;">
#                 <tr>
#                     <th style="border: 1px solid black; padding: 8px;">Name</th>
#                     <th style="border: 1px solid black; padding: 8px;">Job Title</th>
#                     <th style="border: 1px solid black; padding: 8px;">Experience</th>
#                     <th style="border: 1px solid black; padding: 8px;">Location</th>
#                     <th style="border: 1px solid black; padding: 8px;">Status</th>
#                 </tr>'''

#     for candidate in candidates:
#         html += f'''
#                 <tr>
#                     <td style="border: 1px solid black; padding: 8px;">{candidate["Name"]}</td>
#                     <td style="border: 1px solid black; padding: 8px;">{candidate["JobTitle"]}</td>
#                     <td style="border: 1px solid black; padding: 8px;">{candidate["Experience"]}</td>
#                     <td style="border: 1px solid black; padding: 8px;">{candidate["Location"]}</td>
#                     <td style="border: 1px solid black; padding: 8px;">{candidate["Status"]}</td>
#                 </tr>'''

#     html += '''
#             </table>
#             <p>Best regards,</p>
#             <h3>Swetha,</h3>
#             <h3>Seafy Soft Solutions</h3>
#             Email: hr@seafysoft.com <br>
#             Phone: +1 9105576339 <br>
#             <a href="https://seafysoft.com/">Seafysoft.com/</a> <br>
#             LinkedIn: <a href="https://www.linkedin.com/company/seafy-soft-solutions/">Seafy Soft Solutions</a>
#         </body>
#     </html>'''
    
#     return html

# if __name__ == "__main__":
#     check_new_emails()




# import imaplib
# import email
# import smtplib
# from email.mime.text import MIMEText
# from email.mime.multipart import MIMEMultipart
# import pandas as pd
# import re
# import time
# from datetime import datetime
# import os

# # Login information
# SourceUsername = 'dailyrequriments@gmail.com'
# SourcePassword = 'yulyeentyuykfpav'
# imap_url = 'imap.gmail.com'
# smtp_server = 'smtp.gmail.com'
# smtp_port = 587

# # Load the Excel files
# daily_hotlist_path = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\DailyHotList.xlsx'
# today_requirements_path = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\TodayRequriments.xlsx'

# try:
#     send_candidate_df = pd.read_excel(daily_hotlist_path, sheet_name='Sheet1')
#     print("Loaded daily hotlist successfully.")
# except Exception as e:
#     print(f"Error loading daily hotlist: {e}")

# try:
#     keywords_df = pd.read_excel(daily_hotlist_path, sheet_name='Keywords')
#     print("Loaded keywords successfully.")
# except Exception as e:
#     print(f"Error loading keywords from Excel: {e}")

# # Gmail labels
# label_to_apply = 'Processed'
# restricted_keywords = ['phmailadmin', '.email']  # Add more as needed

# def check_new_emails():
#     connected = False
#     mail = None
#     while True:
#         try:
#             if not connected:
#                 mail = imaplib.IMAP4_SSL(imap_url)
#                 mail.login(SourceUsername, SourcePassword)
#                 mail.select('inbox')
#                 print("Logged in to the email account and selected inbox.")
#                 connected = True

#             while True:
#                 try:
#                     status, response = mail.search(None, 'UNSEEN')
#                     if status != 'OK':
#                         print(f"Failed to search for unseen emails: {response}")
#                         break  # Exit inner loop to reconnect

#                     email_ids = response[0].split()
#                     print(f"Fetched {len(email_ids)} new unseen emails.")

#                     if email_ids:
#                         for email_id in email_ids:
#                             try:
#                                 process_email(mail, email_id)
#                             except Exception as e:
#                                 print(f"Error processing email ID {email_id}: {e}")

#                         # After processing all emails, log out and reconnect
#                         print("Processed all unseen emails, logging out...")
#                         mail.logout()
#                         connected = False
#                         time.sleep(10)  # Wait before reconnecting
#                         break  # Exit inner loop to reconnect

#                     else:
#                         print("No new emails found. Waiting for the next check.")
#                         time.sleep(60)  # Wait before checking for new emails again

#                 except Exception as e:
#                     print(f"Error during email check: {e}")
#                     break  # Exit inner loop to reconnect

#         except Exception as e:
#             print(f"Error logging in or selecting inbox: {e}")
#             connected = False  # Set to reconnect next time
#             time.sleep(60)  # Wait before trying to reconnect

#         finally:
#             if mail:
#                 try:
#                     mail.logout()
#                     print("Logged out from the email account.")
#                 except Exception as e:
#                     print(f"Error logging out: {e}")

#             if not connected:
#                 print("Reconnecting in 60 seconds...")
#                 time.sleep(60)  # Wait before reconnecting

# def process_email(mail, email_id):
#     try:
#         status, msg_data = mail.fetch(email_id, '(RFC822)')
#         for response_part in msg_data:
#             if isinstance(response_part, tuple):
#                 msg = email.message_from_bytes(response_part[1])
#                 subject = msg['Subject']
#                 print(f"Processing email ID {email_id} with subject: {subject}")

#                 # Extract email body
#                 email_body = extract_body(msg)
#                 if email_body:
#                     # Send reply if subject matches keywords
#                     matching_candidates = check_keywords_in_subject(subject)
#                     if matching_candidates:
#                         send_reply(msg, subject, matching_candidates)
#                         update_today_requirements(email_body, subject, msg)

#                 apply_label(mail, email_id, label_to_apply)

#     except Exception as e:
#         print(f"Error fetching or processing email ID {email_id}: {e}")

# def apply_label(mail, email_id, label):
#     try:
#         mail.store(email_id, '+X-GM-LABELS', label)
#         print(f"Applied label '{label}' to email ID {email_id}.")
#     except Exception as e:
#         print(f"Error applying label to email ID {email_id}: {e}")

# def extract_body(msg):
#     try:
#         body = ""
#         if msg.is_multipart():
#             for part in msg.walk():
#                 content_type = part.get_content_type()
#                 content_disposition = str(part.get("Content-Disposition"))

#                 if content_type == "text/plain" and 'attachment' not in content_disposition:
#                     body += part.get_payload(decode=True).decode(part.get_content_charset(), errors='ignore')
#                 elif content_type == "text/html" and 'attachment' not in content_disposition:
#                     body += part.get_payload(decode=True).decode(part.get_content_charset(), errors='ignore')
#         else:
#             body = msg.get_payload(decode=True).decode(msg.get_content_charset(), errors='ignore')

#         print("Email body extracted successfully.")
#         return body
#     except Exception as e:
#         print(f"Error extracting email body: {e}")
#         return ""

# def extract_valid_emails(email_message):
#     email_list = [email_message['Reply-To'], email_message['From']]
#     body = extract_body(email_message)

#     email_addresses = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', body)
#     email_list.extend(email_addresses)

#     email_list = list(set(email_list))
#     valid_emails = [item for item in email_list if item and not any(restricted in item for restricted in restricted_keywords)]
    
#     return valid_emails

# def send_reply(msg, subject, matching_candidates):
#     server = None
#     try:
#         server = smtplib.SMTP(smtp_server, smtp_port)
#         server.starttls()
#         server.login(SourceUsername, SourcePassword)
#         print("Logged in to SMTP server.")
#     except smtplib.SMTPAuthenticationError as auth_error:
#         print(f"Authentication error occurred: {auth_error}")
#         return
#     except Exception as e:
#         print(f"Error connecting to SMTP server: {e}")
#         return

#     body = extract_body(msg)

#     try:
#         reply_to_email = msg['Reply-To'] if msg['Reply-To'] else msg['From']
#         print(f"Reply-to email: {reply_to_email}, Valid candidates: {matching_candidates}")

#         if not any(restricted in reply_to_email for restricted in restricted_keywords):
#             msg_out = create_reply_message(matching_candidates, reply_to_email, subject, body)
#             server.send_message(msg_out)
#             print(f'Sent reply to: {reply_to_email}, Subject: {msg_out["Subject"]}')
#     except Exception as e:
#         print(f"Error sending reply: {e}")
#     finally:
#         if server:
#             server.quit()
#             print("SMTP connection closed.")

# def create_reply_message(matching_candidates, reply_to_email, subject, email_body):
#     msg_out = MIMEMultipart()
#     msg_out['From'] = SourceUsername
#     msg_out['To'] = reply_to_email
#     msg_out['Subject'] = 'Re: ' + subject

#     body_html = create_html_table(matching_candidates)
#     msg_out.attach(MIMEText(body_html, 'html'))
#     print("Reply message created.")

#     return msg_out

# def update_today_requirements(email_body, subject, msg):
#     try:
#         phone_pattern = r'(?<!\d)(\+?\d{1,3}[-.\s]?)?(\(?\d{1,4}\)?[-.\s]?)?(\d{1,4})[-.\s]?(\d{1,4})[-.\s]?(\d{1,9})(?!\d)'
#         phone_matches = re.findall(phone_pattern, email_body)

#         today_date = datetime.now().strftime('%Y-%m-%d')
#         technology = subject  # Replace with logic to determine technology based on subject

#         if phone_matches:
#             vendor_emails = []
#             reply_to_email = msg['Reply-To'] if msg['Reply-To'] else msg['From']
#             valid_phones = []

#             for match in phone_matches:
#                 full_phone = ''.join(part for part in match if part).strip()
#                 if len(re.sub(r'\D', '', full_phone)) >= 10:
#                     valid_phones.append(full_phone)
#                     vendor_emails.append(reply_to_email)

#             if vendor_emails and valid_phones:
#                 vendor_emails_str = ', '.join(set(vendor_emails))  # Deduplicate emails if needed
#                 log_entry = {
#                     'VendorEmail': vendor_emails_str,
#                     'VendorPhone': ', '.join(set(valid_phones)),
#                     'Date': today_date,
#                     'Technology': technology
#                 }

#                 update_excel(today_requirements_path, log_entry)

#     except Exception as e:
#         print(f"Error updating TodayRequirements: {e}")

# def update_excel(file_path, log_entry):
#     try:
#         if os.path.exists(file_path):
#             df_today = pd.read_excel(file_path)
#         else:
#             df_today = pd.DataFrame(columns=['VendorEmail', 'VendorPhone', 'Date', 'Technology'])

#         # Create a DataFrame for the new entry
#         new_entry_df = pd.DataFrame([log_entry])

#         # Use pd.concat to append the new entry
#         df_today = pd.concat([df_today, new_entry_df], ignore_index=True)

#         # Write back to Excel
#         df_today.to_excel(file_path, index=False)
#         print(f"Updated {file_path} successfully with new entry.")
#     except Exception as e:
#         print(f"Error updating Excel file: {e}")

# def check_keywords_in_subject(subject):
#     matching_candidates = []
    
#     all_keywords = keywords_df['Keywords'].str.cat(sep=',').split(',')
#     all_keywords = [keyword.strip().lower() for keyword in all_keywords]
    
#     matching_keywords = [keyword for keyword in all_keywords if keyword in subject.strip().lower()]
#     print(f"Matching keywords in subject: {matching_keywords}")

#     if matching_keywords:
#         for _, row in send_candidate_df.iterrows():
#             candidate_job_title = row['Title'].strip().lower()
#             candidate_name = row['Name']
#             candidate_experience = row['Experience']
#             candidate_location = row['Location']
#             candidate_status = row['Status']
            
#             if any(keyword in candidate_job_title for keyword in matching_keywords):
#                 matching_candidates.append({
#                     'Name': candidate_name,
#                     'JobTitle': row['Title'],
#                     'Experience': candidate_experience,
#                     'Location': candidate_location,
#                     'Status': candidate_status
#                 })
#                 print(f"Matching candidate found: {candidate_name}")

#     return matching_candidates

# def create_html_table(candidates):
#     html = '''<html>
#         <body>
#             <h3>Hello,</h3>
#             <p>At <strong>Seafy Soft Solutions</strong>, we take pride in our team of highly skilled consultants, who are ready to begin working on-site or remotely at your convenience. Whether you have immediate needs or future projects in mind, we are here to support you.</p>
#             <p>Should you have any questions or require further information, please feel free to contact me via email or phone. I look forward to the opportunity to assist you.</p>
#             <h1><span style="background-color: yellow; padding: 0 5px;">Reach us at: hr@seafysoft.com</span></h1>
#             <br/>
#             <p>Here are the candidates that match your requirements:</p>
#             <table style="border-collapse: collapse; width: 100%;">
#                 <tr>
#                     <th style="border: 1px solid black; padding: 8px;">Name</th>
#                     <th style="border: 1px solid black; padding: 8px;">Job Title</th>
#                     <th style="border: 1px solid black; padding: 8px;">Experience</th>
#                     <th style="border: 1px solid black; padding: 8px;">Location</th>
#                     <th style="border: 1px solid black; padding: 8px;">Status</th>
#                 </tr>'''

#     for candidate in candidates:
#         html += f'''
#                 <tr>
#                     <td style="border: 1px solid black; padding: 8px;">{candidate["Name"]}</td>
#                     <td style="border: 1px solid black; padding: 8px;">{candidate["JobTitle"]}</td>
#                     <td style="border: 1px solid black; padding: 8px;">{candidate["Experience"]}</td>
#                     <td style="border: 1px solid black; padding: 8px;">{candidate["Location"]}</td>
#                     <td style="border: 1px solid black; padding: 8px;">{candidate["Status"]}</td>
#                 </tr>'''

#     html += '''
#             </table>
#             <p>Best regards,</p>
#             <h3>Swetha,</h3>
#             <h3>Seafy Soft Solutions</h3>
#             Email: hr@seafysoft.com <br>
#             Phone: +1 9105576339 <br>
#             <a href="https://seafysoft.com/">Seafysoft.com/</a> <br>
#             LinkedIn: <a href="https://www.linkedin.com/company/seafy-soft-solutions/">Seafy Soft Solutions</a>
#         </body>
#     </html>'''
    
#     return html

# if __name__ == "__main__":
#     check_new_emails()



import imaplib
import email
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import re
import time
from datetime import datetime
import os
import html


# Login information
SourceUsername = 'dailyrequriments@gmail.com'
SourcePassword = 'yulyeentyuykfpav'
imap_url = 'imap.gmail.com'
smtp_server = 'smtp.gmail.com'
smtp_port = 587

# Load the Excel files
daily_hotlist_path = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\DailyHotList.xlsx'
today_requirements_path = r'C:\Users\somep\OneDrive\Documents\Automation\Seafysoft\TodayRequriments.xlsx'

# Load DataFrames
try:
    send_candidate_df = pd.read_excel(daily_hotlist_path, sheet_name='Sheet1')
    keywords_df = pd.read_excel(daily_hotlist_path, sheet_name='Keywords')
    print("Loaded data successfully.")
except Exception as e:
    print(f"Error loading data: {e}")

label_to_apply = 'Processed'
restricted_keywords = ['phmailadmin', '.email']

def check_new_emails():
    while True:
        mail = None
        try:
            # Login to the email
            mail = imaplib.IMAP4_SSL(imap_url)
            mail.login(SourceUsername, SourcePassword)
            mail.select('inbox')
            print("Logged in and selected inbox.")

            # Fetch unseen emails
            status, response = mail.search(None, 'UNSEEN')
            if status != 'OK':
                print(f"Failed to search for unseen emails: {response}")
                break

            email_ids = response[0].split()
            print(f"Fetched {len(email_ids)} new unseen emails.")

            if email_ids:
                for email_id in email_ids:
                    try:
                        process_email(mail, email_id)
                    except Exception as e:
                        print(f"Error processing email ID {email_id}: {e}")

                # After processing, log out and wait before logging back in
                print("Logging out after processing batch...")
                mail.logout()
                time.sleep(10)  # Wait before reconnecting
            else:
                print("No new emails found. Waiting for the next check...")
                time.sleep(60)  # Wait before checking for new emails again

        except Exception as e:
            print(f"Error: {e}")
        finally:
            if mail:
                try:
                    mail.logout()
                    print("Logged out from the email account.")
                except Exception as e:
                    print(f"Error logging out: {e}")

def process_email(mail, email_id):
    try:
        status, msg_data = mail.fetch(email_id, '(RFC822)')
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                subject = msg['Subject']
                print(f"Processing email ID {email_id} with subject: {subject}")

                email_body = extract_body(msg)
                if email_body:
                    matching_candidates = check_keywords_in_subject(subject)
                    if matching_candidates:
                        send_reply(msg, subject, matching_candidates)
                        update_today_requirements(email_body, subject, msg)

                apply_label(mail, email_id, label_to_apply)

    except Exception as e:
        print(f"Error fetching or processing email ID {email_id}: {e}")

def extract_body(msg):
    try:
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    return part.get_payload(decode=True).decode(part.get_content_charset(), errors='ignore')
        else:
            return msg.get_payload(decode=True).decode(msg.get_content_charset(), errors='ignore')
    except Exception as e:
        print(f"Error extracting email body: {e}")
        return ""

def apply_label(mail, email_id, label):
    try:
        mail.store(email_id, '+X-GM-LABELS', label)
        print(f"Applied label '{label}' to email ID {email_id}.")
    except Exception as e:
        print(f"Error applying label to email ID {email_id}: {e}")

def check_keywords_in_subject(subject):
    matching_candidates = []
    all_keywords = keywords_df['Keywords'].str.cat(sep=',').split(',')
    all_keywords = [keyword.strip().lower() for keyword in all_keywords]

    matching_keywords = [keyword for keyword in all_keywords if keyword in subject.strip().lower()]
    print(f"Matching keywords in subject: {matching_keywords}")

    if matching_keywords:
        for _, row in send_candidate_df.iterrows():
            candidate_job_title = row['JobTitle'].strip().lower()
            if any(keyword in candidate_job_title for keyword in matching_keywords):
                matching_candidates.append(row.to_dict())
                print(f"Matching candidate found: {row['Name']}")

    return matching_candidates

# def send_reply(msg, subject, matching_candidates):
#     try:
#         server = smtplib.SMTP(smtp_server, smtp_port)
#         server.starttls()
#         server.login(SourceUsername, SourcePassword)

#         reply_to_email = msg['Reply-To'] if msg['Reply-To'] else msg['From']
#         body_html = create_html_table(matching_candidates)
#         msg_out = MIMEMultipart()
#         msg_out['From'] = SourceUsername
#         msg_out['To'] = reply_to_email
#         msg_out['Subject'] = 'Re: ' + subject
#         msg_out.attach(MIMEText(body_html, 'html'))

#         server.send_message(msg_out)
#         print(f'Sent reply to: {reply_to_email}, Subject: {msg_out["Subject"]}')
#     except Exception as e:
#         print(f"Error sending reply: {e}")
#     finally:
#         server.quit()

def send_reply(msg, subject, matching_candidates):
    """Send a reply to the original email with matching candidates, redirecting replies to hr@seafysoft.com."""
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(SourceUsername, SourcePassword)

            reply_to_email = msg['Reply-To'] if msg['Reply-To'] else msg['From']
            body_html = create_html_table(matching_candidates)
            msg_out = MIMEMultipart()
            msg_out['From'] = SourceUsername
            msg_out['To'] = reply_to_email
            msg_out['Subject'] = 'Re: ' + subject
            
            # Set the Reply-To header to the HR email address
            msg_out.add_header('Reply-To', 'hr@seafysoft.com')
            
            msg_out.attach(MIMEText(body_html, 'html'))

            server.send_message(msg_out)
            print(f'Sent reply to: {reply_to_email}, Subject: {msg_out["Subject"]}')
            # logging.info(f'Sent reply to: {reply_to_email}, Subject: {msg_out["Subject"]}')
    except Exception as e:
        print(f"Error sending reply: {e}")


def create_html_table(candidates):
    if not isinstance(candidates, list):
        raise ValueError("candidates must be a list")
    
    html_content = '''<html>
        <body>
            <h3>Hello,</h3>
            <p>At <strong>Seafy Soft Solutions</strong>, we take pride in our team of highly skilled consultants...</p>
            <h1><span style="background-color: yellow; padding: 0 5px;">Reach us at: hr@seafysoft.com</span></h1>
            <br/>
            <p>Here are the candidates that match your requirements:</p>
            <table style="border-collapse: collapse; width: 100%;">
                <tr>
                    <th style="border: 1px solid black; padding: 8px; background-color: #4375d3; color: white;">Name</th>
                    <th style="border: 1px solid black; padding: 8px; background-color: #4375d3; color: white;">Job Title</th>
                    <th style="border: 1px solid black; padding: 8px; background-color: #4375d3; color: white;">Experience</th>
                    <th style="border: 1px solid black; padding: 8px; background-color: #4375d3; color: white;">Location</th>
                    <th style="border: 1px solid black; padding: 8px; background-color: #4375d3; color: white;">Status</th>
                </tr>'''

    for candidate in candidates:
        if not isinstance(candidate, dict):
            raise ValueError("Each candidate must be a dictionary")
        
        # Ensure all keys are present
        required_keys = ["Name", "JobTitle", "Experience", "Location", "Status"]
        for key in required_keys:
            if key not in candidate:
                raise ValueError(f"Candidate is missing required key: {key}")

        html_content += f'''
                <tr>
                    <td style="border: 1px solid black; padding: 8px;">{html.escape(candidate["Name"])}</td>
                    <td style="border: 1px solid black; padding: 8px;">{html.escape(candidate["JobTitle"])}</td>
                    <td style="border: 1px solid black; padding: 8px;">{html.escape(candidate["Experience"])}</td>
                    <td style="border: 1px solid black; padding: 8px;">{html.escape(candidate["Location"])}</td>
                    <td style="border: 1px solid black; padding: 8px;">{html.escape(candidate["Status"])}</td>
                </tr>'''

    html_content += '''
            </table>
            <p>Best regards,</p>
            <strong>Swetha</strong><br>
            <em>Seafy Soft Solutions</em><br>
            📧 <a href="mailto:hr@seafysoft.com">hr@seafysoft.com</a><br>
            📞 +1 910-557-6339<br>
            🌐 <a href="https://seafysoft.com/">Seafysoft.com</a><br>
            🔗 <a href="https://www.linkedin.com/company/seafy-soft-solutions/">LinkedIn: Seafy Soft Solutions</a><br>
            <br>
            <em>Disclaimer: This email and any attachments are confidential and intended solely for the use of the individual or entity to whom they are addressed. If you have received this email in error, please notify the sender and delete it from your system. Unauthorized use, disclosure, or distribution of this communication is prohibited.</em>
        </body>
    </html>'''
    
    return html_content


def update_today_requirements(email_body, subject, msg):
    try:
        phone_pattern = r'(?<!\d)(\+?\d{1,3}[-.\s]?)?(\(?\d{1,4}\)?[-.\s]?)?(\d{1,4})[-.\s]?(\d{1,4})[-.\s]?(\d{1,9})(?!\d)'
        phone_matches = re.findall(phone_pattern, email_body)

        today_date = datetime.now().strftime('%Y-%m-%d')
        technology = subject  # Replace with logic to determine technology based on subject

        if phone_matches:
            vendor_emails = []
            reply_to_email = msg['Reply-To'] if msg['Reply-To'] else msg['From']
            valid_phones = []

            for match in phone_matches:
                full_phone = ''.join(part for part in match if part).strip()
                if len(re.sub(r'\D', '', full_phone)) >= 10:
                    valid_phones.append(full_phone)
                    vendor_emails.append(reply_to_email)

            if vendor_emails and valid_phones:
                vendor_emails_str = ', '.join(set(vendor_emails))
                log_entry = {
                    'VendorEmail': vendor_emails_str,
                    'VendorPhone': ', '.join(set(valid_phones)),
                    'Date': today_date,
                    'Technology': technology
                }

                update_excel(today_requirements_path, log_entry)

    except Exception as e:
        print(f"Error updating TodayRequirements: {e}")

def update_excel(file_path, log_entry):
    try:
        if os.path.exists(file_path):
            df_today = pd.read_excel(file_path)
        else:
            df_today = pd.DataFrame(columns=['VendorEmail', 'VendorPhone', 'Date', 'Technology'])

        new_entry_df = pd.DataFrame([log_entry])
        df_today = pd.concat([df_today, new_entry_df], ignore_index=True)
        df_today.to_excel(file_path, index=False)
        print(f"Updated {file_path} successfully with new entry.")
    except Exception as e:
        print(f"Error updating Excel file: {e}")

if __name__ == "__main__":
    check_new_emails()

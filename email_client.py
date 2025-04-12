import os
import streamlit as st
from exchangelib import Credentials as ExchangeCredentials, Account, Configuration, DELEGATE
import imaplib
import email
from email.header import decode_header
from datetime import datetime, timedelta
import chardet
import base64
import re
import ssl
import logging

# Get logger for this module
logger = logging.getLogger(__name__)

# --- Helper Function for Decoding --- #
def decode_mime_header(header):
    """Decodes email headers to handle different charsets."""
    if not header:
        return ""
    decoded_parts = decode_header(header)
    parts = []
    for content, charset in decoded_parts:
        try:
            if isinstance(content, bytes):
                if charset:
                    parts.append(content.decode(charset, errors='replace'))
                else:
                    # Try default encodings if charset is not specified
                    try:
                        parts.append(content.decode('utf-8', errors='replace'))
                    except UnicodeDecodeError:
                        parts.append(content.decode('latin-1', errors='replace'))
            elif isinstance(content, str):
                parts.append(content)
        except (LookupError, UnicodeDecodeError) as e:
            logger.warning(f"Could not decode header part: {content} with charset {charset}. Error: {e}")
            parts.append(str(content) if isinstance(content, bytes) else content) # Append raw on error
    return "".join(parts)

class EmailClient:
    def __init__(self):
        self.gmail_connection = None
        self.outlook_account = None
        logger.info("EmailClient initialized.")

    def _extract_gmail_id(self, message_id_header):
        """Extracts the core 16-character hex ID from Gmail's Message-ID header."""
        if not message_id_header:
            return None
        # Example: <CAD=k_+a=b-cd...> or <blabla@mail.gmail.com>
        # We want the part within the angle brackets, often before the '@'
        match = re.search(r'<([^@>]+)@[^>]*>', message_id_header)
        if match:
            return match.group(1)
        else:
            # Fallback: If no '@', maybe the whole thing inside <> is the ID?
            match_no_at = re.search(r'<([^>]+)>', message_id_header)
            if match_no_at:
                # Simple check if it looks like a hex ID (crude but better than nothing)
                potential_id = match_no_at.group(1)
                if len(potential_id) > 10 and all(c in '0123456789abcdefABCDEF+=' for c in potential_id):
                     return potential_id
            logger.warning(f"Could not extract standard Gmail ID format from: {message_id_header}")
            # As a final fallback, return the cleaned header content
            return message_id_header.strip('<>')

    def connect_gmail(self, email, password):
        """Connect to Gmail using IMAP"""
        logger.info(f"Attempting to connect to Gmail IMAP for {email[:3]}...")
        try:
            # Use SSL context for security
            context = ssl.create_default_context()
            self.gmail_connection = imaplib.IMAP4_SSL('imap.gmail.com', ssl_context=context)
            r, data = self.gmail_connection.login(email, password)
            if r == 'OK':
                logger.info(f"Gmail IMAP login successful for {email[:3]}...")
                # Select Inbox after successful login
                r_select, data_select = self.gmail_connection.select('inbox')
                if r_select == 'OK':
                     logger.info("Gmail inbox selected successfully.")
                     return True
                else:
                     logger.error(f"Failed to select Gmail inbox: {data_select}")
                     self.disconnect_gmail() # Disconnect on failure
                     return False
            else:
                logger.error(f"Gmail IMAP login failed for {email[:3]}... Response: {data}")
                self.disconnect_gmail()
                return False
        except imaplib.IMAP4.error as e:
            logger.error(f"Gmail IMAP connection error for {email[:3]}...: {e}", exc_info=True)
            self.disconnect_gmail()
            # Re-raise a more specific error potentially?
            # raise ConnectionError(f"Gmail IMAP error: {e}") from e
            return False
        except Exception as e:
            logger.error(f"Unexpected error during Gmail connection for {email[:3]}...: {e}", exc_info=True)
            self.disconnect_gmail()
            # raise ConnectionError(f"Unexpected Gmail error: {e}") from e
            return False

    def connect_outlook(self, email, password):
        """Connect to Outlook using Exchange"""
        logger.info(f"Attempting to connect to Outlook (Exchange) for {email[:3]}...")
        try:
            creds = ExchangeCredentials(username=email, password=password)
            # Try autodiscover first
            config = Configuration(server='outlook.office365.com', credentials=creds)
            self.outlook_account = Account(primary_smtp_address=email, config=config,
                                           autodiscover=False, access_type=DELEGATE)
            # Test connection by accessing inbox (lightweight operation)
            inbox_info = self.outlook_account.inbox.total_count
            logger.info(f"Outlook connection successful for {email[:3]}... Inbox count: {inbox_info}")
            return True
        except Exception as e:
            logger.error(f"Outlook connection failed for {email[:3]}...: {e}", exc_info=True)
            self.outlook_account = None # Ensure account is None on failure
            # raise ConnectionError(f"Outlook connection error: {e}") from e
            return False

    def get_gmail_emails(self, offset=0, limit=20):
        """Fetch emails from Gmail with pagination"""
        if not self.gmail_connection:
            logger.error("Cannot fetch Gmail emails: Not connected.")
            return []

        emails = []
        try:
            logger.info(f"Fetching Gmail emails: offset={offset}, limit={limit}")
            self.gmail_connection.select('INBOX')
            _, messages = self.gmail_connection.search(None, 'ALL')
            
            # Get total number of messages
            total_messages = len(messages[0].split())
            
            # Calculate start and end indices for pagination
            start_idx = max(0, total_messages - offset - limit)
            end_idx = total_messages - offset
            
            # Get messages for current page
            message_nums = messages[0].split()[start_idx:end_idx]
            
            for num in message_nums:
                _, msg_data = self.gmail_connection.fetch(num, '(RFC822)')
                email_body = msg_data[0][1]
                email_message = email.message_from_bytes(email_body)
                
                # Get Gmail message ID
                message_id = email_message.get('Message-ID', '')
                if not message_id:
                    # Try to get UID as fallback
                    _, uid_data = self.gmail_connection.fetch(num, '(UID)')
                    if uid_data and uid_data[0]:
                        message_id = uid_data[0].decode().split()[2].rstrip(')')
                
                # Extract the clean Gmail ID
                gmail_id = self._extract_gmail_id(message_id)
                if not gmail_id:
                    gmail_id = num.decode()
                
                # Decode subject
                subject = ""
                raw_subject = email_message["Subject"]
                if raw_subject:
                    decoded_subject = decode_mime_header(raw_subject)
                    subject = decoded_subject
                
                # Decode sender
                sender = ""
                raw_sender = email_message.get("From")
                if raw_sender:
                    decoded_sender = decode_mime_header(raw_sender)
                    sender = decoded_sender
                
                date = email_message["Date"]
                
                # Get email content
                body = ""
                if email_message.is_multipart():
                    for part in email_message.walk():
                        if part.get_content_type() == "text/plain":
                            content = part.get_payload(decode=True)
                            if content:
                                body = self._decode_email_content(content)
                            break
                else:
                    content = email_message.get_payload(decode=True)
                    if content:
                        body = self._decode_email_content(content)
                
                emails.append({
                    'id': gmail_id,
                    'subject': subject,
                    'sender': sender,
                    'date': date,
                    'snippet': body[:200] if body else '',
                    'source': 'gmail'
                })
        
        except Exception as e:
            logger.error(f"Error fetching Gmail emails: {str(e)}", exc_info=True)
        
        return emails

    def get_outlook_emails(self, offset=0, limit=20):
        """Fetch emails from Outlook with pagination"""
        if not self.outlook_account:
            logger.error("Cannot fetch Outlook emails: Not connected.")
            return []

        emails = []
        try:
            # Get total number of messages
            total_messages = self.outlook_account.inbox.all().count()
            
            # Calculate start and end indices for pagination
            start_idx = offset
            end_idx = min(offset + limit, total_messages)
            
            # Get messages for current page
            for item in self.outlook_account.inbox.all().order_by('-datetime_received')[start_idx:end_idx]:
                # Get the full message to access all properties
                message = self.outlook_account.inbox.get(id=item.id)
                
                # For Outlook, we need the full item.id for deep linking
                outlook_id = str(item.id)
                
                emails.append({
                    'id': outlook_id,
                    'subject': message.subject,
                    'sender': message.sender.email_address if message.sender else '',
                    'date': message.datetime_received.strftime('%a, %d %b %Y %H:%M:%S %z'),
                    'snippet': message.text_body[:200] if message.text_body else '',
                    'source': 'outlook'
                })
        
        except Exception as e:
            logger.error(f"Error fetching Outlook emails: {str(e)}", exc_info=True)
        
        return emails

    def _decode_email_content(self, content):
        """Decode email content with proper encoding detection"""
        if not content:
            return ""
        
        try:
            # Try UTF-8 first
            return content.decode('utf-8')
        except UnicodeDecodeError:
            try:
                # Detect encoding
                encoding = chardet.detect(content)['encoding']
                if encoding:
                    return content.decode(encoding)
                else:
                    # If encoding detection fails, try common encodings
                    for enc in ['latin-1', 'iso-8859-1', 'windows-1252']:
                        try:
                            return content.decode(enc)
                        except UnicodeDecodeError:
                            continue
            except:
                pass
        return "[Content could not be decoded]"

    def disconnect_gmail(self):
        """Disconnect from Gmail"""
        if self.gmail_connection:
            try:
                logger.info("Logging out and closing Gmail IMAP connection.")
                self.gmail_connection.logout()
            except imaplib.IMAP4.error as e:
                logger.warning(f"Error during Gmail logout: {e}. Connection might already be closed.")
            except Exception as e:
                 logger.error(f"Unexpected error during Gmail disconnect: {e}", exc_info=True)
            finally:
                 self.gmail_connection = None
                 logger.info("Gmail connection set to None.")

    def disconnect_outlook(self):
        """Disconnect from Outlook"""
        if self.outlook_account:
            logger.info("Disconnecting from Outlook (clearing account object).")
            # exchangelib doesn't have an explicit disconnect, just stop using the object
            self.outlook_account = None

    def disconnect(self):
        """Disconnect from email services"""
        logger.info("Disconnecting from active email service.")
        self.disconnect_gmail()
        self.disconnect_outlook() 
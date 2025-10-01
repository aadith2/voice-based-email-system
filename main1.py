import imaplib
import email
from email.header import decode_header
import re
import chardet  # Library to detect encoding
import pyttsx3
import tkinter as tk
from tkinter import font as tkfont
import warnings
import time
import speech_recognition as sr
import threading

warnings.filterwarnings('ignore')

engine = pyttsx3.init()

# Function to speak text
def speak(text):
    print(f"Assistant: {text}")
    engine.say(text)
    engine.runAndWait()

# Hardwired email credentials
EMAIL = 'group14abhinav@gmail.com'
PASSWORD = 'jncd fuwu gjuz auxf'  # Use app password if 2FA is enabled
IMAP_SERVER = 'imap.gmail.com'  # Replace with your IMAP server
IMAP_PORT = 993  # Default IMAP over SSL port

def detect_encoding(byte_content):
    """
    Detect the encoding of the byte content using chardet.
    """
    result = chardet.detect(byte_content)
    return result.get('encoding', 'utf-8')  # Default to UTF-8 if detection fails

def extract_sender_name(from_header):
    """
    Extract the sender's name from the "From" header.
    """
    # Decode the "From" header
    decoded_header = decode_header(from_header)[0]
    header_text = decoded_header[0]
    if isinstance(header_text, bytes):
        header_text = header_text.decode(decoded_header[1] if decoded_header[1] else 'utf-8', errors='replace')
    
    # Extract the sender's name
    if '<' in header_text and '>' in header_text:
        # Format: "Sender Name <sender@example.com>"
        sender_name = header_text.split('<')[0].strip()
    else:
        # Format: "sender@example.com"
        sender_name = header_text.strip()
    
    return sender_name

def extract_email_details(EMAIL, PASSWORD):
    try:
        # Connect to the IMAP server
        print("Connecting to IMAP server...")
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)  # Explicitly specify port
        
        # Login to the email account
        print("Logging in...")
        speak("logging in")
        mail.login(EMAIL, PASSWORD)
        print("Login successful!")
        speak("")
        
        # Select the inbox
        mail.select('inbox')
        
        # Search for all emails (sorted by date in descending order)
        status, messages = mail.search(None, 'ALL')
        
        if status == 'OK' and messages[0]:
            # List to store email details
            email_details = []
            
            # Get the list of email IDs
            email_ids = messages[0].split()
            
            # Fetch the top 5 emails (most recent)
            top_5_email_ids = email_ids[-5:]  # Last 5 emails are the most recent
            
            # Iterate over the top 5 emails
            for email_id in top_5_email_ids:
                # Fetch the email
                status, msg_data = mail.fetch(email_id, '(RFC822)')
                
                if status == 'OK':
                    # Parse the email content
                    msg = email.message_from_bytes(msg_data[0][1])
                    
                    # Decode the subject
                    subject, encoding = decode_header(msg['Subject'])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding if encoding else 'utf-8', errors='replace')
                    
                    # Extract the sender's name
                    sender = extract_sender_name(msg['From'])
                    
                    # Initialize counters
                    special_char_count = 0
                    link_count = 0
                    
                    # Extract the email body
                    body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            content_type = part.get_content_type()
                            if content_type == 'text/plain' or content_type == 'text/html':
                                payload = part.get_payload(decode=True)
                                if payload:
                                    encoding = detect_encoding(payload)
                                    try:
                                        body += payload.decode(encoding, errors='replace')
                                    except LookupError:
                                        # Fallback to UTF-8 if encoding is not supported
                                        body += payload.decode('utf-8', errors='replace')
                    else:
                        payload = msg.get_payload(decode=True)
                        if payload:
                            encoding = detect_encoding(payload)
                            try:
                                body = payload.decode(encoding, errors='replace')
                            except LookupError:
                                # Fallback to UTF-8 if encoding is not supported
                                body = payload.decode('utf-8', errors='replace')
                    
                    # Count special characters
                    special_char_count = len(re.findall(r'[^\w\s]', body))
                    
                    # Count links
                    link_count = len(re.findall(r'https?://\S+', body))
                    
                    # Store details as a list
                    email_details.append([special_char_count, link_count, subject, sender])
        
        # Logout from the server
        mail.logout()
        
        # Return the list of email details
        return email_details
    
    except imaplib.IMAP4.error as e:
        print(f"IMAP error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

def on_click(event):
    # """Function to handle click event and capture voice input."""
    # # Create a new window for entering credentials
    # credential_window = tk.Toplevel()
    # credential_window.title("Enter Credentials")
    # credential_window.state('zoomed')  # Open the window in a zoomed state
    
    # # Set dark theme colors
    # bg_color = "#2E3440"  # Dark background
    # text_color = "#D8DEE9"  # Light text
    
    # credential_window.configure(bg=bg_color)
    
    # Create a label for instructions
    # instruction_label = tk.Label(credential_window, text="Please speak your Gmail ID and password.", font=("Helvetica", 14), bg=bg_color, fg=text_color)
    # instruction_label.pack(pady=20)
    
    # # Create a label to display the captured input
    # input_label = tk.Label(credential_window, text="", font=("Helvetica", 12), bg=bg_color, fg=text_color)
    # input_label.pack(pady=20)


    
    # Create a new window to display email details
    email_window = tk.Toplevel()
    email_window.title("Email Details")
    email_window.state('zoomed')  # Open the window in a zoomed state
    
    # Set dark theme colors
    bg_color = "#2E3440"  # Dark background
    text_color = "#D8DEE9"  # Light text
    
    email_window.configure(bg=bg_color)
    
    # Create a label to display email details
    email_label = tk.Label(email_window, text="", font=("Helvetica", 14), bg=bg_color, fg=text_color)
    email_label.pack(pady=20)
    
    # Fetch top 5 emails
    top_5_emails = extract_email_details(EMAIL, PASSWORD)
    if top_5_emails:
        print("Top 5 Emails Details:")
        for email in top_5_emails:
            print(email)
        
        # Initialize spam counter
        spamm = 0
        
        # Iterate over the top 5 emails
        for i in range(0, 5):
            # Update the email label with sender and subject
            email_label.config(text=f"Sender: {top_5_emails[i][3]}\nSubject: {top_5_emails[i][2]}")
            email_window.update()  # Refresh the window to display the updated label
            
            # Predict if the email is spam
            from ml import predict_op
            val = predict_op([[top_5_emails[i][0], top_5_emails[i][1]]])
            if val == 0:
                speak("sender is " + (top_5_emails[i][3]))
                speak("subject is " + (top_5_emails[i][2]))
            else:
                spamm += 1
            
            # Add a delay between iterations
            time.sleep(3)  # 3-second delay
        
        # Display the total number of spam emails found
        email_label.config(text=f"{spamm} spam mails found")
        speak(f"{spamm} spam mails found")
        
    else:
        email_label.config(text="No emails found.")
        speak("No emails found.")

def create_app():
    # Create the main application window
    root = tk.Tk()
    root.title("Voice Assisted Gmail System")
    
    # Open the window in a zoomed (maximized) state
    root.state('zoomed')
    
    # Set dark theme colors
    bg_color = "#2E3440"  # Dark background
    text_color = "#D8DEE9"  # Light text
    
    root.configure(bg=bg_color)
    
    # Create a frame to center the labels
    frame = tk.Frame(root, bg=bg_color)
    frame.place(relx=0.5, rely=0.5, anchor='center')
    
    # Create a custom font for the app name
    app_name_font = tkfont.Font(family="Helvetica", size=24, weight="bold")
    
    # Create a label for the app name
    app_name_label = tk.Label(frame, text="Voice Assisted Gmail System", font=app_name_font, bg=bg_color, fg=text_color)
    app_name_label.pack(pady=10)
    
    # Create a label for the welcome message
    welcome_label = tk.Label(frame, text="Welcome to Voice Assisted Gmail System", font=("Helvetica", 16), bg=bg_color, fg=text_color)
    welcome_label.pack(pady=10)
    
    # Speak the welcome message
    speak("Welcome to Voice Assisted Gmail System")
    time.sleep(2)
    speak("Click anywhere on the screen to start")
    
    # Bind the click event to the entire window
    root.bind("<Button-1>", on_click)
    
    # Run the application
    root.mainloop()

if __name__ == '__main__':
    create_app()
import smtplib
from email.message import EmailMessage
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
import os
import mimetypes
import webbrowser
import json

attachment_path = None
emails_file = "emails.json"

# --- Load or Initialize Emails ---
def load_emails():
    if os.path.exists(emails_file):
        with open(emails_file, 'r') as file:
            return json.load(file)
    return {}

def save_emails(emails):
    with open(emails_file, 'w') as file:
        json.dump(emails, file, indent=4)

emails = load_emails()

# --- Send Email Function ---
def send_email():
    sender = entry_sender.get()
    password = entry_password.get()
    receivers_raw = entry_receiver.get()
    subject = entry_subject.get()
    body = text_body.get("1.0", tk.END)

    if not all([sender, password, receivers_raw.strip(), subject, body.strip()]):
        messagebox.showerror("Error", "Please fill in all fields.")
        return

    receivers = [email.strip() for email in receivers_raw.split(',') if email.strip()]

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ", ".join(receivers)
    msg.set_content(body)

    if attachment_path:
        try:
            mime_type, _ = mimetypes.guess_type(attachment_path)
            mime_type, mime_subtype = mime_type.split('/')

            with open(attachment_path, 'rb') as file:
                msg.add_attachment(file.read(), maintype=mime_type, subtype=mime_subtype,
                                   filename=os.path.basename(attachment_path))
        except Exception as e:
            messagebox.showerror("Attachment Error", f"Failed to attach file:\n{e}")
            return

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender, password)
            smtp.send_message(msg)
        messagebox.showinfo("Success", "Email sent successfully!")
    except Exception as e:
        messagebox.showerror("Failed", f"Error: {e}")

# --- File Chooser ---
def choose_file():
    global attachment_path
    file_path = filedialog.askopenfilename()
    if file_path:
        attachment_path = file_path
        label_file.config(text=f"Attached: {os.path.basename(file_path)}")

# --- App Password Guide ---
def open_app_password_guide():
    result = messagebox.askyesno(
        "App Password Help",
        "To send Gmail from apps like this, you must use an App Password.\n\nWould you like to open the App Password setup page?"
    )
    if result:
        webbrowser.open("https://myaccount.google.com/apppasswords")

# --- Email Management ---
def add_email():
    group = simpledialog.askstring("Add Group", "Enter group name:")
    if not group:
        return
    name = simpledialog.askstring("Add Email", "Enter name:")
    email = simpledialog.askstring("Add Email", "Enter email address:")
    if group and name and email:
        if group not in emails:
            emails[group] = {}
        emails[group][name] = email
        save_emails(emails)
        messagebox.showinfo("Saved", f"'{name}' added to group '{group}'.")

def view_emails():
    if not emails:
        messagebox.showinfo("Emails", "No emails saved.")
        return
    email_list = ""
    for group, members in emails.items():
        email_list += f"\n{group}:\n"
        for name, email in members.items():
            email_list += f"  {name}: {email}\n"
    messagebox.showinfo("Saved Emails", email_list)

def insert_from_group():
    if not emails:
        messagebox.showinfo("No Groups", "You haven't added any emails.")
        return
    group = simpledialog.askstring("Choose Group", f"Enter group name:\nAvailable: {', '.join(emails.keys())}")
    if group and group in emails:
        email_addresses = emails[group].values()
        entry_receiver.insert(tk.END, ', '.join(email_addresses))
    else:
        messagebox.showerror("Not Found", "Group not found.")

# --- GUI Setup ---
root = tk.Tk()
root.title("Email Sender with Groups")
root.geometry("500x850")
root.resizable(False, False)

# Setting background color and font style
root.config(bg="#f2f2f2")

# Header Text
header_label = tk.Label(root, text="Email Sender with Groups", font=("Helvetica", 16, "bold"), bg="#f2f2f2")
header_label.pack(pady=20)

# Sender Email
tk.Label(root, text="Your Gmail Address:", bg="#f2f2f2", font=("Arial", 10)).pack()
entry_sender = tk.Entry(root, width=50, font=("Arial", 12))
entry_sender.pack(pady=5)

# App Password
tk.Label(root, text="Your App Password:", bg="#f2f2f2", font=("Arial", 10)).pack()
entry_password = tk.Entry(root, width=50, show="*", font=("Arial", 12))
entry_password.pack(pady=5)

# App Password Guide Button
tk.Button(root, text="How to get App Password?", command=open_app_password_guide, fg="blue", bg="#f2f2f2", font=("Arial", 9)).pack(pady=2)

# Receiver Email
tk.Label(root, text="To (Comma separated emails):", bg="#f2f2f2", font=("Arial", 10)).pack()
entry_receiver = tk.Entry(root, width=50, font=("Arial", 12))
entry_receiver.pack(pady=5)

# Emails section
tk.Button(root, text="➕ Add Email to Group", command=add_email, bg="#0078D7", fg="white", width=25, font=("Arial", 12)).pack(pady=5)
tk.Button(root, text="📋 View All Emails", command=view_emails, bg="#0078D7", fg="white", width=25, font=("Arial", 12)).pack(pady=5)
tk.Button(root, text="📤 Use Group", command=insert_from_group, bg="#0078D7", fg="white", width=25, font=("Arial", 12)).pack(pady=5)

# Subject
tk.Label(root, text="Subject:", bg="#f2f2f2", font=("Arial", 10)).pack()
entry_subject = tk.Entry(root, width=50, font=("Arial", 12))
entry_subject.pack(pady=5)

# Message
tk.Label(root, text="Message:", bg="#f2f2f2", font=("Arial", 10)).pack()
text_body = tk.Text(root, height=10, width=50, font=("Arial", 12))
text_body.pack(pady=5)

# Attachment
tk.Button(root, text="Choose Attachment", command=choose_file, bg="#0078D7", fg="white", font=("Arial", 12)).pack(pady=10)
label_file = tk.Label(root, text="No file selected", bg="#f2f2f2", font=("Arial", 10))
label_file.pack()

# Send Email Button
tk.Button(root, text="Send Email", command=send_email, bg="green", fg="white", width=25, font=("Arial", 14)).pack(pady=20)

root.mainloop()

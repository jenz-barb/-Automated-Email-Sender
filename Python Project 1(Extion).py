import smtplib
import getpass
import openpyxl
from tkinter import filedialog, Label, Entry, Button, messagebox, Tk

def read_recipients_from_text(file_path):
    with open(file_path, 'r') as file:
        recipients = file.read().splitlines()
    return recipients

def read_recipients_from_excel(file_path):
    recipients = []
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    for row in sheet.iter_rows(values_only=True):
        recipients.append(row[0])
    return recipients

def send_email(server, sender_email, password, subject, message, recipients):
    try:
        for recipient in recipients:
            server.sendmail(sender_email, recipient, f"Subject: {subject}\n\n{message}")
            print(f"Email sent successfully to {recipient}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to send email: {e}")

def send_emails():
    smtp_server = smtp_entry.get()
    sender_email = email_entry.get()
    password = password_entry.get()
    subject = subject_entry.get()
    message = message_entry.get('1.0', 'end')

    try:
        server = smtplib.SMTP(smtp_server, 587)
        server.starttls()
        server.login(sender_email, password)

        recipient_file = file_entry.get()
        if recipient_file.endswith('.txt'):
            recipients = read_recipients_from_text(recipient_file)
        elif recipient_file.endswith('.xlsx'):
            recipients = read_recipients_from_excel(recipient_file)
        else:
            messagebox.showerror("Error", "Invalid file format. Please provide a .txt or .xlsx file.")
            return

        send_email(server, sender_email, password, subject, message, recipients)
        server.quit()
        messagebox.showinfo("Success", "Emails sent successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def browse_file():
    filename = filedialog.askopenfilename()
    file_entry.delete(0, 'end')
    file_entry.insert(0, filename)

root = Tk()
root.title("Automated Email Sender")

Label(root, text="SMTP Server:").grid(row=0, column=0, padx=5, pady=5)
smtp_entry = Entry(root)
smtp_entry.grid(row=0, column=1, padx=5, pady=5)

Label(root, text="Your Email:").grid(row=1, column=0, padx=5, pady=5)
email_entry = Entry(root)
email_entry.grid(row=1, column=1, padx=5, pady=5)

Label(root, text="Password:").grid(row=2, column=0, padx=5, pady=5)
password_entry = Entry(root, show="*")
password_entry.grid(row=2, column=1, padx=5, pady=5)

Label(root, text="Subject:").grid(row=3, column=0, padx=5, pady=5)
subject_entry = Entry(root)
subject_entry.grid(row=3, column=1, padx=5, pady=5)

Label(root, text="Message:").grid(row=4, column=0, padx=5, pady=5)
message_entry = Entry(root, width=50)
message_entry.grid(row=4, column=1, padx=5, pady=5)

Label(root, text="Recipient List File:").grid(row=5, column=0, padx=5, pady=5)
file_entry = Entry(root)
file_entry.grid(row=5, column=1, padx=5, pady=5)
Button(root, text="Browse", command=browse_file).grid(row=5, column=2, padx=5, pady=5)

Button(root, text="Send Emails", command=send_emails).grid(row=6, column=0, columnspan=2, padx=5, pady=5)

root.mainloop()
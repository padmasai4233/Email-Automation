from flask import Flask, render_template, request
import openpyxl
import smtplib, ssl
from email.message import EmailMessage

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('upload.html')

@app.route('/send', methods=['POST'])
def send_emails():
    sender_email = request.form['sender_email']
    password = request.form['password']
    subject = request.form['subject']
    file = request.files['excel_file']

    # Load Excel file
    wb = openpyxl.load_workbook(file)
    sheet = wb.active

    status_list = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        name, receiver_email, message_body = row

        # ✅ Skip invalid or blank rows
        if not name or not receiver_email or not message_body:
            status_list.append((receiver_email if receiver_email else '❓ Missing Email', '❌ Skipped: Missing fields'))
            continue

        # Compose email
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg.set_content(f"Hi {name},\n\n{message_body}")

        # Send email using SMTP (Gmail)
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as server:
                server.login(sender_email, password)
                server.send_message(msg)
            status_list.append((receiver_email, '✅ Sent'))
        except Exception as e:
            status_list.append((receiver_email, f'❌ Failed: {str(e)}'))

    return render_template('status.html', results=status_list)

if __name__ == '__main__':
    app.run()

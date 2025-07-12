# app.py
import os
import pandas as pd
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from flask import Flask, render_template, request
import time
import random

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['LOG_FOLDER'] = 'sent_logs'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['LOG_FOLDER'], exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['excel_file']
        sender_email = request.form['sender_email']
        sender_password = request.form['sender_password']
        email_subject = request.form['email_subject']
        email_body = request.form['email_body']

        if file and file.filename.endswith('.xlsx'):
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)

            # ✅ Gmail Login Test
            try:
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as server:
                    server.login(sender_email, sender_password)
            except Exception as e:
                return f"<h3>❌ Gmail Login Failed:</h3><pre>{str(e)}</pre><a href='/'>Try again</a>"

            df = pd.read_excel(filepath, sheet_name='Sheet1')
            sample_row = df.iloc[0]
            sample_message = email_body.replace("{{Name}}", str(sample_row['Name']))
            sample_message = sample_message.replace("{{Email}}", str(sample_row['Email']))
            preview = df.head().to_html(index=False)

            return render_template(
                'preview.html',
                table=preview,
                sample_email=sample_message,
                filepath=filepath,
                sender_email=sender_email,
                sender_password=sender_password,
                email_subject=email_subject,
                email_body=email_body
            )

    return render_template('index.html')


@app.route('/send-emails', methods=['POST'])
def send_emails():
    filepath = request.form['filepath']
    sender_email = request.form['sender_email']
    sender_password = request.form['sender_password']
    email_subject = request.form['email_subject']
    email_body = request.form['email_body']

    df = pd.read_excel(filepath, sheet_name='Sheet1')
    context = ssl.create_default_context()
    successes = []

    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as server:
        try:
            server.login(sender_email, sender_password)
        except Exception as e:
            return f"Login failed: {str(e)}"

        for index, row in df.iterrows():
            time.sleep(random.uniform(4, 7))
            recipient_email = row['Email']
            recipient_name = row['Name']

            personalized_body = email_body.replace("{{Name}}", str(recipient_name))
            personalized_body = personalized_body.replace("{{Email}}", str(recipient_email))

            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient_email
            msg['Subject'] = email_subject
            msg.attach(MIMEText(personalized_body, 'plain'))

            try:
                server.sendmail(sender_email, recipient_email, msg.as_string())
                print(f"Sent to {recipient_name} at {recipient_email}")
                successes.append({'Name': recipient_name, 'Email': recipient_email})
            except Exception as e:
                print(f"Failed to send to {recipient_email}: {str(e)}")

    # Save log
    log_path = os.path.join(app.config['LOG_FOLDER'], 'sent_log.xlsx')
    pd.DataFrame(successes).to_excel(log_path, index=False)

    return f"Emails sent successfully! Log saved to: {log_path}"

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)

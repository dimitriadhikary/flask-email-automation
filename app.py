import os
import time
import ssl
import random
import datetime
import smtplib
import pandas as pd
from flask import Flask, render_template, request, send_file
from email.message import EmailMessage
from email.mime.application import MIMEApplication

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['LOG_FOLDER'] = 'sent_logs'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['LOG_FOLDER'], exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check for file
        if 'excel_file' not in request.files:
            return 'No file part'

        file = request.files['excel_file']

        if file.filename == '':
            return 'No selected file'

        # Validate file extension
        if file and file.filename.endswith('.xlsx'):
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)

            # Read excel
            df = pd.read_excel(filepath, sheet_name='Sheet1')
            preview = df.head().to_html(index=False)

            # Capture form data for email sending
            sender_email = request.form.get('sender_email', '')
            sender_password = request.form.get('sender_password', '')
            email_subject = request.form.get('email_subject', '')
            email_body = request.form.get('email_body', '')

            return render_template(
                'preview.html',
                table=preview,
                filepath=filepath,
                sender_email=sender_email,
                sender_password=sender_password,
                email_subject=email_subject,
                email_body=email_body
            )
        else:
            return 'Invalid file type. Only .xlsx allowed.'

    return render_template('index.html')


@app.route('/send-emails', methods=['POST'])
def send_emails():
    filepath = request.form['filepath']
    sender_email = request.form['sender_email']
    sender_password = request.form['sender_password']
    email_subject = request.form['email_subject']
    email_body = request.form['email_body']

    df = pd.read_excel(filepath, sheet_name='Sheet1')
    statuses = []

    context = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, sender_password)

            for _, row in df.iterrows():
                try:
                    recipient_email = row['Email']
                    name = row['Name']

                    personalized_body = email_body.replace("{{Name}}", name).replace("{{Email}}", recipient_email)

                    msg = EmailMessage()
                    msg['Subject'] = email_subject
                    msg['From'] = sender_email
                    msg['To'] = recipient_email
                    msg.set_content(personalized_body, subtype='plain', charset='utf-8')

                    server.send_message(msg)
                    statuses.append("Success")
                except Exception as e:
                    statuses.append(f"Failure: {str(e)}")

                time.sleep(random.uniform(4, 7))

    except Exception as e:
        return f"Failed to connect or login to Gmail: {str(e)}"

    df['Status'] = statuses
    timestamp = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
    log_filename = f"log_{timestamp}.xlsx"
    log_path = os.path.join(app.config['LOG_FOLDER'], log_filename)
    df.to_excel(log_path, index=False)

    # Send log file to sender
    try:
        msg = EmailMessage()
        msg["Subject"] = "Residency Email Log"
        msg["From"] = sender_email
        msg["To"] = sender_email
        msg.set_content("Attached is your log file showing which emails were successfully sent.", charset='utf-8')

        with open(log_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=log_filename)
            part['Content-Disposition'] = f'attachment; filename="{log_filename}"'
            msg.add_attachment(part.get_payload(decode=True), maintype='application', subtype='octet-stream', filename=part.get_filename())

        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, sender_password)
            server.send_message(msg)
    except Exception as e:
        print(f"Failed to send log email: {e}")

    failed_df = df[df['Status'].str.startswith('Failure')]
    failed_html = failed_df.to_html(index=False) if not failed_df.empty else None

    return render_template('log_download.html', log_filename=log_filename, failed_rows=failed_html)


@app.route('/download-log/<filename>')
def download_log(filename):
    return send_file(os.path.join(app.config['LOG_FOLDER'], filename), as_attachment=True)


if __name__ == '__main__':
    import os
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)

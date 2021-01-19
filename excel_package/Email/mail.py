import smtplib
from email.message import EmailMessage

EMAIL_ADDRESS = "dorongafsou13@gmail.com"
# EMAIL_PASSWORD = '<maill_pass>'

contacts = ['dorongafsou95@gmail.com', 'test@example.com']


def send_mail(stock_name, operation, mail_to_send=EMAIL_ADDRESS):
    print("send mail")
    msg = EmailMessage()
    msg['Subject'] = f'You need to {operation} {stock_name} stock'
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = mail_to_send

    msg.set_content('This is a plain text email')

    msg.add_alternative("""\
    <!DOCTYPE html>
    <html>
        <body>
            <h1 style="color:SlateGray;">Excel Stock!</h1>
        </body>
    </html>
    """, subtype='html')

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)



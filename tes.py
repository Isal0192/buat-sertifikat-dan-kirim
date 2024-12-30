import smtplib

SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 465

try:
    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
        server.ehlo()
        print("SMTP server connection successful!")
except Exception as e:
    print(f"Failed to connect to SMTP server: {e}")


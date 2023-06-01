import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from colorama import Fore
from time import sleep

your_email = "yourmail@mail.com" # Your Email
your_password = "password" # Your Email Password => App Password with Gmail SMTP

server = smtplib.SMTP('smtp.gmail.com', 587, timeout=3600)
server.starttls()
server.ehlo()
server.login(your_email, your_password)

email_list = pd.read_excel('Emails.xlsx')

emails = email_list['Email']
# Your Email Message
message = """Hi
This message is for testing

mihranw@gmail.com
"""
fs = open('success.txt', 'w') # For Log
fe = open('errors.txt', 'w') # For Log

for i in range(len(emails)):
    try:
        msg = MIMEMultipart()
        msg['Subject'] = 'Your Email Subject'
        msg['From'] = your_email
        msg['To'] = emails[i]
        msg.attach(MIMEText(message))
        server.sendmail(your_email, emails[i], msg.as_string())
        fs.write(f'{i + 1} : {emails[i]}\n') # Add To Log File
        print(Fore.GREEN + f"{i + 1}) Email Was Successfully Sent To : ", emails[i])
        sleep(2)
    except Exception as e:
        fe.write(f'{i + 1} : {emails[i]}\n')
        fe.write(f'{e}\n') # Add To Log File
        fe.write(f'---------------------------------\n')
        print(Fore.RED + f"{i + 1}) There Was An Error Sending Email To : ", emails[i])

print(Fore.WHITE)
server.close()



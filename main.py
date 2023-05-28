import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
# Set up email credentials
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_username = 'abs.tbzmed@gmail.com'
smtp_password = 'spikbkkpatwjzndl' # r2E;vsA5D4#Q2q

# Set up email content
email_subject = 'تاثیر شاخص توانایی کارآفرینانه در نانوفناوری'
email_body = """\
فعال محترم عرصه نانو
سلام
توجه به توانایی و مهارت های فردی در حوزه نوآوری و کارآفرینی امری بسیار مهم است. 
هدف پرسشنامه ی پیش رو بررسی توانمندی های افراد جهت استفاده از فرصت هایی که در زمینه کارآفرینی و فناوری نانو برای آن ها پیش می آید است .
این پرسشنامه داری 33 سوال می باشد، لذا با تخصيص زمان ارزشمندتان به طور دقيق آنرا تکميل نمایید. 
شايان ذکر است اين اطلاعات کاملاً محرمانه تلقي شده و صرفاً جهت دستيابي به اهداف پژوهش به صورت کلي مورد استفاده قرار خواهد گرفت. 
در ضمن با عنایت به اینکه این پرسشنامه جهت بررسی وضعیت موجود است ، و بر اساس آن در رابطه با سیاست گذاری و نقشه راه مربوطه تصمیم گیری خواهد شد ؛ خواهشمند است با پرهیز از ایده آل گرایی و به طور شفاف وضعیت موجود هریک از سوالات را در پاسخ اظهار دارید.
پيشاپيش از همکاري صميمانه شما سپاسگزاري مي‌شود.

لینک پرسشنامه
https://survey.porsline.ir/s/qyJq56Q
"""
# Read Excel file
excel_file = pd.read_excel('test.xlsx')
# Set up email message
msg = MIMEMultipart()
msg['Subject'] = email_subject
msg['From'] = smtp_username
msg['Cc'] = ', '.join(excel_file['Email'])
msg.attach(MIMEText(email_body))

# Send email
with smtplib.SMTP(smtp_server, smtp_port)as server:
    server.ehlo()
    server.starttls()
    server.login(smtp_username, smtp_password)
    server.sendmail(smtp_username, excel_file['Email'], msg.as_string(), msg['Cc'].split(','))
    server.quit()

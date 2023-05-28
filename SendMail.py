import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from colorama import Fore
from time import sleep

your_email = "abs.tbzmed@gmail.com"
your_password = "spikbkkpatwjzndl"

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.ehlo()
server.login(your_email, your_password)

email_list = pd.read_excel('test.xlsx')

emails = email_list['Email']

for i in range(len(emails)):
    try:
        message = """
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
        msg = MIMEMultipart()
        msg['Subject'] = 'تاثیر شاخص توانایی کارآفرینانه در نانوفناوری'
        msg['From'] = your_email
        msg['To'] = emails[i]
        msg.attach(MIMEText(message))
        server.sendmail(your_email, emails[i], msg.as_string())
        print(Fore.GREEN + "Email Was Successfully Sent To : ", emails[i])
        sleep(2)
    except Exception as e:
        print(Fore.RED +"There Was An Error Sending Email To : ", emails[i])

print(Fore.WHITE)
server.close()



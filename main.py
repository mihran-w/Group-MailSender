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
message = """فعال محترم عرصه نانو


سلام علیکم
     همانگونه که مستحضرید در حوزه نوآوری و کارآفرینی، توجه به توانایی ­ها و مهارت ­های فردی امری بسیار مهم می­ باشد.
فلذا پرسشنامه­ ی پیش رو به منظور بررسی توانمندی های افراد جهت استفاده از فرصت هایی که در زمینه کارآفرینی و فناوری نانو برای آن ها پیش می آید ایجاد گردیده است.
پرسشنامه مذکور داری 33 سوال می باشد و در انتهای نامه پیوست گردیده، لذا با تخصيص زمان ارزشمندتان به طور دقيق آنرا تکميل نمایید.
شايان ذکر است اين اطلاعات کاملاً محرمانه تلقي شده و صرفاً جهت دستيابي به اهداف پژوهش به صورت کلي مورد بررسی قرار خواهد گرفت.
در ضمن با توجه به اینکه این پرسشنامه جهت بررسی وضعیت موجود است و بر اساس آن در رابطه با سیاست گذاری و نقشه راه مربوطه تصمیم گیری خواهد شد؛ خواهشمند است با پرهیز از ایده آل گرایی و به طور شفاف وضعیت موجود هریک از سوالات را در پاسخ اظهار نمایید.


پيشاپيش از همکاري صميمانه شما کمال تشکر و قدردانی را داریم. 




لینک پرسشنامه:
https://survey.porsline.ir/s/qyJq56Q



زهرا اکبری مقدم 

دانشجو ارشد نانو فناوری پزشکی 

دانشکده علوم نوین پزشکی 

دانشگاه علوم پزشکی تبریز

z.akbarimoghadam77@gmail.com"""
for i in range(len(emails)):
    try:
        msg = MIMEMultipart()
        msg['Subject'] = 'بررسي توانایی و مهارت افراد در حوزه کارآفرینی(برای فعالین عرصه نانو)'
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



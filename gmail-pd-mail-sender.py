import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from cfp import EmailSender

sender=EmailSender('email.ini')

# 讀取 Excel 文件
df = pd.read_excel('Test.xlsx')

# 設定 SMTP 伺服器
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.login(sender.sender_email, sender.password)

for index, row in df.iterrows():
    receiver_email = row['Email']
    receiver_name = row['Name']
    
    # 建立郵件物件
    message = MIMEMultipart("alternative")
    message["Subject"] = "Email Test"
    message["From"] = sender.sender_email
    message["To"] = receiver_email
    
    # 郵件正文
    html = f"""\
    <html>
      <body>
        <p>Hi {receiver_name},<br>
           Test<br>
           This is a test email from Python.
        </p>
        <a href="https://www.google.com">Google</a>
      </body>
    </html>
    """
    
    # 加入郵件正文

    message.attach(MIMEText(html, "html"))
    
    # 發送郵件
    server.sendmail(sender.sender_email, receiver_email, message.as_string())

# 關閉 SMTP 伺服器連接
server.quit()

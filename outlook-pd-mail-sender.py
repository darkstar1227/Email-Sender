import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from cfp import EmailSender

sender=EmailSender('config.ini')

# 讀取 Excel 文件
df = pd.read_excel('Test.xlsx')
print('read excel success')
print(df)
print('start sending email')

with open(r'template.html', 'r', encoding='utf-8') as f:
  template = f.read()
print('template read success')
# 設定 SMTP 伺服器
try:
  server = smtplib.SMTP('smtp.outlook.com', 587)
  server.starttls()
  server.login(sender.sender_email, sender.password)
  print('login success')

  for index, row in df.iterrows():
      receiver_email = row['Email']
      receiver_name = row['Name']
      print(receiver_name+' : '+receiver_email)
      # 建立郵件物件
      message = MIMEMultipart("alternative")
      message["Subject"] = "Email Test"
      message["From"] = sender.sender_email
      message["To"] = receiver_email
      
      # 郵件正文
      html = template
      html = html.replace('{receiver_name}',receiver_name)
      
      # 加入郵件正文

      message.attach(MIMEText(html, "html"))
      try:
        # 發送郵件
        server.sendmail(sender.sender_email, receiver_email, message.as_string())
        print(f'{receiver_email} send success')
      except Exception as e:
        print(f'{receiver_email} send failed. Error message: {str(e)}')
        continue

  # 關閉 SMTP 伺服器連接
  server.quit()
except Exception as e:
        # print("郵件發送失敗。錯誤訊息：", str(e))
        print(f'Send mail failed. Error message: {str(e)}')
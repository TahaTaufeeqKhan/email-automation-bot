from email.message import EmailMessage
import pandas as pd
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def dataframe_to_html(df):
    return df.to_html()

excel_file_path = r"C:\Users\Tkhan\Downloads\Taha_Internship_Projects\EmailAutomation\mockExcel.xlsx"
df = pd.read_excel(excel_file_path, sheet_name='Invoice')

html_table = dataframe_to_html(df)

email_sender = 'taha.taufeeq7878@gmail.com'
email_password = 'qqbd dxmv qopo mtkj'
email_receiver = 'king.cobra7183@gmail.com'
subject = "Excel sheet in HTML Form"

em = EmailMessage()
em['From'] = email_sender
em['To'] = email_receiver
em['Subject'] = subject
em.set_content(html_table, subtype='html')  # Set HTML content directly

context = ssl.create_default_context()

with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
    smtp.login(email_sender, email_password)
    smtp.sendmail(email_sender, email_receiver, em.as_string())

import os, pandas as pd, mysql.connector, smtplib
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
timestamp = datetime.now() - timedelta()
date = timestamp.strftime('%m-%d')

dirExit = "Saidas/"
portinFileExit = dirExit + "Base_Dados_%s.xlsx" % (date)

query = '''
SELECT * FROM TESTE;
'''

conn = mysql.connector.connect(
  host="localhost:3306",
  user="usuario",
  password="12345",
  database="banco"
)

df = pd.read_sql(query, con=conn)

conn.close()

df.to_excel(portinFileExit, index=False)

email = 'teste@gmail.com'
password = 'Password123'
send_to_email = 'destinatario@gmail.com'
subject = 'Base de dados '+ date
message = '''

Bom dia,

Segue a base de dados atualizada. 

Email automático.
'''

msg = MIMEMultipart()
msg['From'] = email
msg['To'] = send_to_email
msg['Subject'] = subject

msg.attach(MIMEText(message, 'plain'))

# Setup the attachment
filename = os.path.basename(portinFileExit)
attachment = open(portinFileExit, "rb")
part = MIMEBase('application', 'octet-stream')
part.set_payload(attachment.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

# Attach the attachment to the MIMEMultipart object
msg.attach(part)

server = smtplib.SMTP('SMTP.office365.com',587)
server.starttls()
server.login(email, password)
text = msg.as_string()
server.sendmail(email, send_to_email, text)
server.quit()

attachment.close()


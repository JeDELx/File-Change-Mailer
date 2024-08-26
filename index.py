import os
import conf
import openpyxl
import smtplib
import time
import datetime
import logging
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Log ayarları
logging.basicConfig(filename='file_monitor.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Excel dosyasının yolu ve adı
home = os.path.expanduser("~")
excel_file = os.path.join(home, "Desktop", "Rapor_2023.xlsx")

# Son değiştirilme zamanı
last_modified_time = None

# Dosyanın kayıt edilip edilmediğini denetleme
class ExcelFileHandler(FileSystemEventHandler):
    def on_modified(self, event):
        global last_modified_time
        if event.src_path == excel_file:
            current_time = time.time()
            if last_modified_time is None or (current_time - last_modified_time) > 5:
                last_modified_time = current_time
                logging.info(f"{excel_file} dosyasi kaydedildi.")
                send_email()

def send_email():
    # E-posta ayarları
    sender_email = conf.sender_email
    receiver_email = conf.receiver_email
    password = conf.password
    smtp_server = conf.smtp_server
    smtp_port = conf.smtp_port

    # Şuanki tarihi oluşturma (Gün, Ay, Yıl)
    today = datetime.date.today().strftime("%d %B %Y")

    # E-posta içeriği
    subject = "Günlük Rapor"
    body = f"{today} Tarihinde yapmış olduğum işler ekte yer almaktadır. Saygılar."

    # E-posta oluşturma
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = subject

    # E-posta gövdesini ekleme
    message.attach(MIMEText(body, 'plain'))

    # Excel dosyasını ekleme
    with open(excel_file, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(excel_file)}")
        message.attach(part)

    # E-postayı gönderme
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, password)
            text = message.as_string()
            server.sendmail(sender_email, receiver_email, text)
            logging.info("Mail basariyla yollandi")
    except Exception as e:
        logging.error(f"Mail yollanamadi: {e}")

# Dosya varlığının kontrolü
if not os.path.exists(excel_file):
    logging.error(f"Error: File not found: {excel_file}")
    exit()

# Watchdog kullanarak dosyayı izleme
event_handler = ExcelFileHandler()
observer = Observer()
observer.schedule(event_handler, path=os.path.dirname(excel_file), recursive=False)
observer.start()

logging.info("Basariyla baslatildi ve dosya izleniyor.")

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()
    logging.info("Izleme durduruldu.")
observer.join()

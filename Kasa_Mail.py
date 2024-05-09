import os
import win32com.client
from pathlib import Path
import re
import datetime
import pytz
import pandas as pd

# E-posta kriterleri
current_time = datetime.datetime.now()
yesterday = current_time - datetime.timedelta(days=1)
kriter_tarih = yesterday.replace(hour=18, minute=0, second=0, microsecond=0)

# Dosyaları kaydedeceğiniz klasör yolu
kayit_klasoru = "C:\\Users\\LJ_Emre\\Desktop\\kasalarr"

# Outlook'u başlat
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# KASA klasörünü seç
root_folder = outlook.GetDefaultFolder(6)
kasa_folder = None
for folder in root_folder.Folders:
    if folder.Name == "KASA":
        kasa_folder = folder
        break

if not kasa_folder:
    print("KASA klasörü bulunamadı.")
    exit()

# E-postaları oku
messages = kasa_folder.Items

# Zaman damgasına göre e-postaları sırala
messages.Sort("[ReceivedTime]", True)

# Kriteri karşılayan e-postaları kontrol et ve belirtilen klasöre kaydet
counter = 0
seen_senders = dict()
for message in messages:
    try:
        received_time = message.ReceivedTime
    except Exception:  # AttributeError yerine Exception kullanın
        continue  # Bu ileti doğru özniteliklere sahip değil, devam et

    received_time = received_time.astimezone(pytz.utc).replace(tzinfo=None)  # Saat dilimi bilgisini UTC'ye çevir ve kaldır
    received_time = received_time.replace(microsecond=0)
    sender_email = message.SenderEmailAddress

    if received_time > kriter_tarih and (sender_email not in seen_senders or received_time > seen_senders[sender_email]):
        attachments = message.Attachments
        has_excel = False
        for attachment in attachments:
            if attachment.FileName.endswith(('.xls', '.xlsx')):
                has_excel = True
                break

        if has_excel:
            seen_senders[sender_email] = received_time
            for attachment in attachments:
                # Sadece Excel dosyalarını kaydet
                if attachment.FileName.endswith(('.xls', '.xlsx')):
                    # Geçerli dosya adını düzeltme
                    file_name = re.sub(r'[<>:"/\\|?*]', "", attachment.FileName)

                    # 06_23'den sonraki ifadeleri sil
                    file_name = re.sub(r'03_24.*(\.xls[x]?)', r'04_24\1', file_name)

                    attachment.SaveAsFile(os.path.join(kayit_klasoru, file_name))
                    print(f"{file_name} {kayit_klasoru} klasörüne kaydedildi.")

                    # Eğer dosya adı _23 ile bitmiyorsa, indirilen Excel dosyasını sil
                    if not file_name.endswith("_24.xls") and not file_name.endswith("_24.xlsx"):
                        os.remove(os.path.join(kayit_klasoru, file_name))
                        print(f"{file_name} silindi.")

            counter += 1
            if counter >= 53:
                break
# Eksik kasa dosyalarını belirleyin ve e-postaları gönderin
df = pd.read_excel("D:\\Kasalar\\Config\\Config.xlsx",  sheet_name= "map", index_col=None) #Mağaza Mail
df2 = pd.read_excel("D:\\Kasalar\\Config\\Config.xlsx",  sheet_name= "Parametre", index_col=None) #
df["File"]=df["File"]+df2.iloc[5, 1]
file_email_map = dict(zip(df['File'], df['Email']))

def send_email_outlook(to_address, subject, body, cc_address=None):
    outlook = win32com.client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.To =to_address
    if cc_address:
        message.CC = cc_address
    message.Subject = subject
    message.Body = body
    message.Send()

cc_address = "kasa@ljmagazacilik.com"
previous_day = datetime.datetime.now() - datetime.timedelta(days=1)
email_subject = f"Gönderilmeyen {previous_day.strftime('%d.%m.%Y')} Tarihli Kasa"
email_body = f"Merhaba, \n\n {previous_day.strftime('%d.%m.%Y')} tarihli kasanızı göndermenizi rica ederim.\n\nEren YİĞİT\nMuhasebe Personeli\n\nLJ Mağazacılık San.Ve Tic. A.Ş.   Company : 0850 420 0 501\nMacun Mah. 171 Cad. 2/35           www.ljmagazacilik.com\nYenimahalle / ANKARA"

print("İşlem tamamlandı.")


sent_emails = []
sent_email_count = 0

for file, email in file_email_map.items():
    file_path = os.path.join(kayit_klasoru, file + ".xls")
    file_path_xlsx = os.path.join(kayit_klasoru, file + ".xlsx")
    if not (os.path.exists(file_path) or os.path.exists(file_path_xlsx)):
        send_email_outlook(email, email_subject, email_body, cc_address=cc_address)
        sent_emails.append(email)
        sent_email_count += 1

print(f"Gönderilen e-postaların listesi:")
for email in sent_emails:
    print(email)

print(f"\nToplam gönderilen e-posta sayısı: {sent_email_count}")

print("İşlem tamamlandı.")
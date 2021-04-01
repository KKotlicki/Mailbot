import os
import smtplib
import ssl
import pandas as pd
import random
import time
from dotenv import load_dotenv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

if __name__ == "__main__":
    if not os.path.exists('.env'):
        USER = input("Podaj adres email wysyłający:\n")
        PASS = input("Podaj hasło do maila:\n")
        SERV = input("Podaj adres serwera (SMTP):\n")
        PORT = input("Podaj port serwera wychodzący (SMTP):\n")
        with open('.env', 'w+') as wr:
            wr.write(f"USER=\'{USER}\'\nPASS=\'{PASS}\'\nSERV=\'{SERV}\'\nPORT=\'{PORT}\'")

    if not os.path.exists('resources/config.py'):
        subject = input("Podaj tytuł maila:\n").encode('utf8')
        sender = input("Podaj nazwę wysyłającego:\n").encode('utf8')
        with open('resources/config.py', 'w+') as wr:
            wr.write(f"subject = {subject}\nsender = {sender}\n")

    load_dotenv()
    from resources.config import subject, sender

    EMAIL_ADDRESS = os.getenv('USER')
    EMAIL_PASSWORD = os.getenv('PASS')
    SERVER = os.getenv('SERV')
    SV_PORT = int(os.getenv('PORT'))
    ctx = ssl.create_default_context()
    text = ''
    xlsx = ''
    for file in os.listdir("resources"):
        if file.endswith(".txt"):
            with open(os.path.join("resources/", file), encoding='utf-8') as rd:
                text = rd.read()
    for file in os.listdir("resources"):
        if file.endswith(".xlsx"):
            xlsx = os.path.join("resources/", file)
    if xlsx == '':
        print('Nie znaleziono bazy danych .xlsx w folderze /resources')
        exit()
    if text == '':
        print('Nie znaleziono treści wiadomości .txt w folderze /resources')
        exit()

    df = pd.read_excel(xlsx)
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

    with smtplib.SMTP_SSL(SERVER, SV_PORT, context=ctx) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        success = 0
        err_count = 0
        for index, row in df.iterrows():
            if pd.isnull(row['Data Wysłania']):
                try:
                    message = MIMEMultipart("alternative")
                    message["Subject"] = Header(subject, 'utf-8')
                    message["From"] = sender
                    message["To"] = row['Mail']
                    newtext = f"{text}<p>"
                    for i in range(0, random.randint(0, 5)):
                        newtext = f"{newtext}&nbsp;"
                    newtext = f"{newtext}</p>"
                    message.attach(MIMEText(newtext, 'html'))
                    smtp.sendmail(EMAIL_ADDRESS, row['Mail'], message.as_string())
                    df.at[index, 'Data Wysłania'] = pd.to_datetime("today").date()
                    success += 1
                    print(f'{success}: {row["Firma"]}')
                except Exception as err:
                    print(f"Error - {row['Firma']}:\n{err}")
                    err_count += 1
                time.sleep(random.randint(8, 12))
            else:
                df.at[index, 'Data Wysłania'] = row['Data Wysłania'].date()
        print(f"Wysłano: {success}\nNie wysłano: {err_count}")
        df.to_excel(xlsx)
        print('Zapisano pomyślnie!')

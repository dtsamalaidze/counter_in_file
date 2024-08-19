import datetime
import email
import imaplib
import os
from email.header import decode_header

import pandas as pd
from pandas_ods_reader import read_ods

from template import of_xml, delete_files_in_folder

# Настройка подключения к почтовому ящику
mail_pass = "vfU4tHcNtuarf8ULLBxc"  # пароль от почтового ящика
username = "counter@eobuv.ru"  # адресс почтового ящика
imap_server = "imap.mail.ru"  # почтовый сервер
imap = imaplib.IMAP4_SSL(imap_server)  # Подключение к почтовому серверу
imap.login(username, mail_pass)  # авторизация на почтовом сервере

# Список магазинов
markets = [
    'EO-Vladikavkaz2',
    'EO-Vladikavkaz3'
]

# set_market = ''  # выбранный магазин для цикла
imap.select("INBOX")  # подключение к списку входящих сообщений
k = str(imap.select("INBOX")[1])  # получение общего количества входящих сообщений
count_msg = int(k[3:-2])  # Отформатированное количество входящих сообщений
sum_msg = imap.search(None, "ALL")  # Список сообщений в почтовом ящике

numb_msg = 1  # Порядковый номер сообщения

while numb_msg <= count_msg:  # Цикл перебора писем
    try:
        res, msg = imap.fetch(str(numb_msg), '(RFC822)')  # Хедер полученного письма
        msg = email.message_from_bytes(msg[0][1])  # Побайтовое прочтение письма
        decode_header(msg["Subject"])[0][0].decode()  # Тема письма

        # Дата письма
        letter_date = datetime.datetime(
            email.utils.parsedate_tz(msg["Date"])[0],
            email.utils.parsedate_tz(msg["Date"])[1],
            email.utils.parsedate_tz(msg["Date"])[2]
        ).strftime("%d.%m.%Y")

        date_now = datetime.datetime.now() - datetime.timedelta(days=1)  # отнимаем количество дней от текущей даты
        date_now = date_now.strftime("%d.%m.%Y")  # Форматируем дату

        if letter_date != date_now:
            numb_msg += 1
            continue
        else:
            for part in msg.walk():
                for market in markets:
                    if market in decode_header(msg["Subject"])[0][0].decode():
                        if "application" in part.get_content_type():
                            filename = part.get_filename()
                            filename = str(email.header.make_header(email.header.decode_header(filename)))
                            if ".ods" in filename:
                                fp = open(os.path.join(market, filename), 'wb')
                                fp.write(part.get_payload(decode=1))
                                fp.close()
                                path = f"{market}\\{filename}"
            numb_msg += 1
            continue


    except TypeError:
        continue

try:
    for market in markets:
        time = []
        count_in = []
        count_out = []
        date_format_msg = f'{letter_date[6:]}{letter_date[3:5]}{letter_date[:2]}'
        start = 0
        files = os.listdir(market)
        df = read_ods(f'{market}\\{files[0]}')
        ex = df.to_excel(f'{market}\\{market}-{date_format_msg}-1000.xlsx')
        events_data = pd.read_excel(
            f'{market}\\{market}-{date_format_msg}-1000.xlsx', skiprows=3
        )
        while start < 47:
            value_time = events_data.iloc[start, 1]
            value_in = events_data.iloc[start, 2]
            value_out = events_data.iloc[start, 3]
            time.append(f'{value_time}:00')

            count_in.append(int(value_in))
            count_out.append(int(value_out))
            start += 1

        of_xml(market, letter_date, time, count_in, count_out)
        delete_files_in_folder(f'{market}\\')

except FileNotFoundError:
    print(f"Файл не существует: {numb_msg}, дата = {count_msg}")

import win32com.client as win32
import requests

import tkinter as tk
from tkinter import *
from tkinter import scrolledtext
from tkinter.font import Font

from bs4 import BeautifulSoup
from datetime import datetime
import datetime as date


# TK WINDOW START ----------------------------------------------
window = tk.Tk()

window.title("TAZZ Reviews Checker")
window.geometry('600x400')

rows = 0
while rows < 10:
    window.rowconfigure(rows, weight=1)
    window.columnconfigure(rows, weight=1)
    rows += 1

lbl = Label(window, text="Minimum rating", justify="right",
            font=Font(family='Helvetica', size=18))
lbl.grid(column=0, row=0, pady=10, padx=10, columnspan=1, sticky="W")

default_min_rating = IntVar()
default_min_rating.set(3)

spin = Spinbox(window, from_=1, to=5, width=3, font=Font(family='Helvetica', size=14),
               textvariable=default_min_rating)
spin.grid(column=0, row=0, pady=10, padx=10)

default_mails = StringVar()
default_mails.set('Cristian.Spafiu@hardrockcafe.ro; Marius.Baban@hardrockcafe.ro; Andrei.Meluca@Hardrockcafe.ro; '
                  'Eduard.Garbea@Hardrockcafe.ro; Fron.Theophile@Hardrockcafe.ro; Marian.Andrei@Hardrockcafe.ro; '
                  'Paul.Iacob@hardrockcafe.ro; Corneliu.Carstea@Hardrockcafe.ro')


mails = Entry(window, width=250, font=Font(
    family='Helvetica', size=13), textvariable=default_mails)
mails.grid(column=0, row=3, pady=5, padx=5)

txt = scrolledtext.ScrolledText(window, width=40, height=10)
txt.grid(column=0, row=4, columnspan=2, pady=10, padx=10)
txt.tag_config('error', foreground='red')
txt.tag_config('sent', foreground='green')
txt.yview(END)


# TK WINDOW END ----------------------------------------------


MAX_ERRORS = 5
RE_RUN_TIME_MS = 40000
error_notification_sent = False
first_order_id = None
error_count = 0

error_receiver = ''

session = requests.Session()

payload = {'email': '',
           'password': '',
           'remember': 1
           }

page = session.post("https://eucemananc.ro/supplier/login", data=payload)


def get_next_day():
    tomorrow = (date.datetime.now() + date.timedelta(days=1))
    day = tomorrow.day
    month = tomorrow.month
    return f'{day}.{month}'


def read_file(file):
    with open(file, 'r') as f:
        return f.read()


def write_log(text: str) -> str:
    with open("log.txt", "a") as f:
        if text == 'log':
            f.write(
                f'Executed at: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")} \n')
            return f'Executed at: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")} \n'
        elif text == 'error':
            f.write(
                f'Error at: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")} \n')
            return f'Error at: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")} \n'
        elif text == 'sent':
            f.write(
                f'Mail sent at: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")} \n')
            return f'Mail sent at: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")} \n'


def create_soup() -> dict:
    global page
    page = session.get('https://eucemananc.ro/supplier/reviews')
    soup = BeautifulSoup(page.content, "html.parser")
    data_rows = soup.find_all("tr", limit=2)  # limit the results of table rows to 2 since the one that we need is the second one
    body_dict = {}

    for index, data in enumerate(data_rows):

        if index == 1:  # index 1

            body_dict['name'] = data.find("a").text.strip()
            body_dict['phone'] = data.find("b").text.strip()
            body_dict['order_no'] = data.find_all("td")[1].text.strip()
            date_and_time = data.find_all("td")[2].text.strip()
            body_dict['date'] = date_and_time.split(" ")[0]
            body_dict['time'] = date_and_time.split(" ")[1]
            body_dict['rating'] = int(data.find_all("td")[3].text.strip()[:-3])
            body_dict['comment'] = data.find_all("td")[4].text.strip()

            return body_dict


def send_mail(to: str, body_dict: dict, error=False, error_mail_subject=None):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.To = to
    if not error:
        name = body_dict['name']
        phone = body_dict['phone']
        order_no = body_dict['order_no']
        date = body_dict['date']
        time = body_dict['time']
        rating = body_dict['rating']
        comment = body_dict['comment']

        mail.Subject = f'Review nou de {rating} ★'
        mail.Body = ''
        mail.HTMLBody = f'''
            <h3 align="center">In data de {date}, la ora {time} ati primit un review de {rating} ★ de la {name}, <a href="tel:{phone}">{phone}</a></h3>
            <h4 align="center">Numar comanda: <a href="https://eucemananc.ro/supplier/orders/details/{order_no[1:]}">{order_no}</a></h4>
                        '''

        if comment != "":
            mail.HTMLBody += f'''
            <div style="max-width: 80%;" align="center">
                <h3 align="center">Mesaj: <br> {comment}</h3>
            </div>
            '''

        mail.HTMLBody += f'<h3 align="center">https://eucemananc.ro/supplier/reviews</h3>'

        mail.HTMLBody += f'''
    <h5>Buna,</h5> 
    <h5>Am nevoie de un voucher in valoare de [] de lei, valabil timp de o luna, incepand cu data de {get_next_day()} pentru:</h5> 
    <h5>{name}</h5>
    <h5>{phone}</h5>
    <h5>{order_no}</h5>
    <h5>Mersi,</h5>
    '''

        mail.HTMLBody += f'''<a href = "mailto:gabriela.bogdan@tazz.ro?cc=Operations@Hardrockcafe.ro&subject=Voucher%20Hard%20Rock"> 
    Click pentru mail catre Gabriela </a> '''

    else:
        error_to_write = body_dict['error']
        mail.Subject = error_mail_subject
        attachment = mail.Attachments.Add(
            'C:/Users/Delivery/Pictures/error.png')
        attachment.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "errorId")
        mail.Body = ''
        mail.HTMLBody = f'''
                    <div align="center">
                        <h3>{error_to_write}</h3>
                        <img src="cid:errorId" align="center" alt="error-img" />
                    </div>
                        '''
    mail.Send()


def send_review_mail():
    global first_order_id, error_notification_sent, error_count
    receiver_list = mails.get()
    min_rating = int(spin.get(), 16)
    mail_body = create_soup()
    try:
        if (mail_body['rating'] <= min_rating and mail_body['order_no'] != first_order_id) \
                or (mail_body['name'] == 'Cristian Spafiu' and mail_body['order_no'] != first_order_id):
            send_mail(receiver_list, mail_body)
            txt.insert(END, write_log('sent'), 'sent')
            txt.yview(END)

    except TypeError:
        error_count += 1
        txt.insert(END, write_log('error'), 'error')
        txt.yview(END)
        error_body_dict = {
            'error': f'''Oh no! We have an error getting the reviews. Trying to resolve it...'''}

        if not error_notification_sent:
            send_mail(error_receiver,
                      body_dict=error_body_dict,
                      error=True,
                      error_mail_subject='Error getting reviews')
            error_notification_sent = True

        if error_count < MAX_ERRORS:
            window.after(RE_RUN_TIME_MS, send_review_mail)
        elif error_count == MAX_ERRORS:
            error_not_resolved = {
                'error': f'''APPLICATION STOPPED! We encountered an error that couldn't be resolved! Call Spaf...'''}
            send_mail(receiver_list,
                      body_dict=error_not_resolved,
                      error=True,
                      error_mail_subject="APPLICATION STOPPED! We encountered an error that couldn't be resolved!")

    else:
        if error_count != 0:
            error_count = 0
            error_body_dict = {'error': f'Error resolved. Rock On!'}
            send_mail(error_receiver,
                      body_dict=error_body_dict,
                      error=True,
                      error_mail_subject="Error resolved, resuming normal activity")
        error_notification_sent = False
        first_order_id = mail_body['order_no']
        txt.insert(END, write_log('log'))
        txt.yview(END)
        window.after(RE_RUN_TIME_MS, send_review_mail)


btn = Button(window, text="Run program", font=Font(
    family='Helvetica', size=12), command=send_review_mail)
btn.grid(column=0, row=2, columnspan=2, pady=10, padx=10)

window.mainloop()

import smtplib
import imghdr
from email.message import EmailMessage
from PIL import ImageFont, ImageDraw, Image
import pandas as pd
import cv2
import numpy as np
import time
import xlrd
import os

def clean():
       x="aids2"
       print("Cleaning {} folder ........".format(str(x)))
       for passs in os.listdir("aids1/"):
        os.remove("aids1/{}".format(passs))
       print("completed.......")

def read_info():
    df = pd.read_excel("list2.xlsx")
    print(df)
    names = df['Name'].to_list()
    emails = df['Email'].to_list()
    rolls = df['Roll'].to_list()
    print(names)
    print(emails)
    print(rolls)
    return emails, names, rolls


def create_pass(names, rolls):
    for name, roll in zip(names, rolls):
        template = cv2.imread('template/pass.png')
        template_conv = cv2.cvtColor(template, cv2.COLOR_BGR2RGB)
        arr_img = Image.fromarray(template_conv)
        var_draw = ImageDraw.Draw(arr_img)
        earth = ImageFont.truetype("template/earth.ttf", 25)
        #var_draw.text((80, 500), name, font=earth, fill='white')
        earth = ImageFont.truetype("template/earth.ttf", 30)
        var_draw.text((130, 552), roll, font=earth, fill='white')
        final_res = cv2.cvtColor(np.array(arr_img), cv2.COLOR_RGB2BGR)
        cv2.imwrite("aids2/{}.png".format(name), final_res)
        print("{}'s Pass generated".format(name))


def send_mail(emails, names):
    Sender_Email = "sasivatsal7122@gmail.com"
    Password = "xxxxxxxxxx"

    for email, name in zip(emails, names):
        Reciever_Email = email

        smtp = smtplib.SMTP('smtp.gmail.com', 587)
        smtp.ehlo()
        smtp.starttls()

        smtp.login(Sender_Email, Password)

        newMessage = EmailMessage()
        newMessage['Subject'] = "We Gladly Invite You to Euphoria 2k22!"
        newMessage['From'] = Sender_Email
        newMessage['To'] = Reciever_Email
        newMessage.set_content(f"""
Hey Folk,
We are absolutely thrilled to invite you {name} to EUPHORIA 2k22: a senior - junior interaction party on 19th May, 2022 (Thursday) at Seminar Hall-1, Main Block. Having you at this party will be wonderful! 
This is the 1st of the legacy of EUPHORIA. The party is going to be loads of fun and exuberance.
There's an entry pass attached to this mail by which you'll be let in.
Looking forward to having you there!☺️

Venue : Visveswaraya Hall (seminar hall-1)
Time : 9AM to 4:00PM
Lunch : 1:00PM
                               """)

        with open(f'aids2/{name}.png', 'rb') as f:
            image_data = f.read()
            image_type = imghdr.what(f.name)
            image_name = f.name

        newMessage.add_attachment(
            image_data, maintype='image', subtype=image_type, filename=image_name)
        smtp.sendmail(from_addr=Sender_Email,
                      to_addrs=Reciever_Email, msg=newMessage.as_string())
        print(f'{name} Email sent')

    smtp.quit()


if __name__ == '__main__':  
    clean()
    emails, names, rolls = read_info()
    create_pass(names, rolls)
    send_mail(emails, names)

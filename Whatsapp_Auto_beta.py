
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
import time
import datetime
import os
import argparse

root = tk.Tk()
message_var=tk.StringVar(value='Place Your Message Here')

parser = argparse.ArgumentParser(description='PyWhatsapp Guide')
parser.add_argument('--chrome_driver_path', action='store', type=str, default=r'./driver/chromedriver.exe', help='chromedriver executable path')
parser.add_argument('--message', action='store', type=str, default='', help='Enter the msg you want to send')
parser.add_argument('--remove_cache', action='store', type=str, default='False', help='Remove Cache | Scan QR again or Not')
args = parser.parse_args()

if args.remove_cache == 'True':
    os.system('rm -rf User_Data/*')
browser = None
message = []
Link = "https://web.whatsapp.com/"
wait = None
unsaved_Contacts = None
choice = None
docChoice = None
doc_filename = None
imgname= None
not_sent_contacts=[]
not_sent_contacts_try=[]
unsent_message=[]
#unsent_message_try=[]
date_time = datetime.datetime.now().strftime("%H.%M.%S-%d.%m.%y")


canvas1 = tk.Canvas(root, width=300, height=300, bg='lightsteelblue')
canvas1.pack()

def getExcel():
    global unsaved_Contacts, user_name, message,raw_message,message_inp_box
    import_file_path = filedialog.askopenfilename()
    df1 = pd.read_excel(import_file_path, usecols='A')
    df2 = pd.read_excel(import_file_path, usecols='B')
    df3 = pd.read_excel(import_file_path, usecols='C')
    raw_user_name=list(df1['Names'])
    unsaved_Contacts = list(df2['Contact'])
    raw_message = list(df3['Message'])
    user_name=[]
    for s in raw_user_name:
        user_name.append(s.replace(' ', '%20').replace('&','%26').replace('?', '%3F'))
    #print(user_name)
    #print(unsaved_Contacts)
    message_inp_box= input_message.get()

    if(message_inp_box== 'Place Your Message Here'):
        message=[]
        for (r, x) in zip(raw_message,user_name):
            message.append(r.replace(' ', '%20').replace('\n', '%0D%0A').replace('NAME', x).replace('&','%26').replace('?', '%3F'))
            #print(message)
        return unsaved_Contacts, message, user_name



    else:
        msg1 = message_inp_box.replace(' ', '%20').replace('\n', '%0D%0A').replace('&','%26').replace('?', '%3F')
        message = []
        for x in user_name:
            message.append(msg1.replace('NAME', x))
        #print(message)
        return unsaved_Contacts, message, user_name

def try_again():
    global not_sent_contacts_try,unsent_message_try

    for (b, u) in zip(not_sent_contacts, unsent_message):
        link = "https://web.whatsapp.com/send?phone={}&text={}".format(b, u)
        print(link)
        browser.get(link)
        try:
            time.sleep(8)
            wait = WebDriverWait(browser, 40).until(
                EC.presence_of_element_located((By.ID, "pane-side")))
            print("Page is ready!")
            send = browser.find_element_by_xpath("//span[@data-testid='send']")
            send.click()
            time.sleep(1)

        except:
            print('Not sent')
            not_sent_contacts_try.append(b)
            #unsent_message_try.append(u)


        else:
            continue


    def report_folder():
        def report_write():
            save_path = r'./Whatsapp_Auto_Report/'
            name_of_file = date_time
            Report_file_Create = os.path.join(save_path, name_of_file + ".txt")
            file = open(Report_file_Create, "w")
            file.write("Total Number Of People:-  " + str(len(unsaved_Contacts)))
            file.write('\n')
            file.write("Not Sent:- " + str(len(not_sent_contacts_try)))
            file.write('\n')
            file.write("Message:-  " + message[0].replace('%20', ' ').replace('%0D%0A', '\n'))
            file.write('\n')
            file.write('\n')
            file.write("List of Not Sent:- \n" + str(not_sent_contacts_try))
            file.write('\n')
            file.write('\n')
            file.write("List of total contacts:- \n" + str(unsaved_Contacts))
            file.write('\n')
            file.write("Name of People:- \n"+str(user_name))

        parent_dir = r'./'
        directory = "Whatsapp_Auto_Report"
        path = os.path.join(parent_dir, directory)

        try:
            os.mkdir(path)
            report_write()
        except OSError:
            print('folder Exists')
            report_write()
    report_folder()






def auto():
    global imgname, message, mainwait
    def whatsapp_login(chrome_path):
        global  browser, Link
        chrome_options = Options()
        chrome_options.add_argument('--user-data-dir=./User_Data')
        browser = webdriver.Chrome(executable_path=chrome_path, options=chrome_options)
        browser.get(Link)
        wait = WebDriverWait(browser, 90).until(
                        EC.presence_of_element_located((By.ID, "pane-side")))
        browser.maximize_window()
        print("QR scanned")

    def send_unsaved_contact_message():
        send = browser.find_element_by_xpath("//span[@data-testid='send']")
        send.click()
        return




    def sender():
        global unsaved_Contacts,message
        #time.sleep(1)

        for(i,g) in zip(unsaved_Contacts,message):
                link = "https://web.whatsapp.com/send?phone={}&text={}".format(i,g)
                print(link)
                browser.get(link)
                try:
                    time.sleep(5)
                    wait_ = WebDriverWait(browser, 20).until(
                        EC.presence_of_element_located((By.ID, "pane-side")))
                    print("Page is ready!")
                    send_unsaved_contact_message()
                    time.sleep(1)

                except:
                    print('Not sent')
                    not_sent_contacts.append(i)
                    unsent_message.append(g)
                    #print(not_sent_contacts)

                else:
                    continue






    if __name__ == "__main__":

        print("Web Page Open")


        # Let us login and Scan
        print("SCAN YOUR QR CODE FOR WHATSAPP WEB")
        whatsapp_login(args.chrome_driver_path)

    sender()
    print(not_sent_contacts)

    # First time message sending Task Complete
    print("Task Completed")


browseButton_Excel = tk.Button(root,text='   Import Contacts  ', command=getExcel, bg='steelblue3', fg='white',
                               font=('helvetica', 12, 'bold'))
browseButton_Excel.pack(pady = 5)


submit_button = tk.Button(root,text=' Send > ', command=auto, bg='green', fg='white',
                               font=('helvetica', 12, 'bold'))
submit_button.pack(pady = 5)

try_again_button = tk.Button(root,text='Try Again', command=try_again, bg='seagreen', fg='white',
                               font=('helvetica', 12, 'bold'))
try_again_button.pack(pady = 5)


input_message=tk.Entry(root,textvariable= message_var,width=40)
input_message.pack()

Exit = tk.Button(root, text=' Exit X', command=root.destroy, bg='tomato', fg='white',
                                        font=('helvetica', 12, 'bold'))
Exit.pack(pady = 10)


canvas1.create_window(150, 60, window=input_message)
canvas1.create_window(150, 120, window=browseButton_Excel)
canvas1.create_window(250, 220, window=submit_button)
canvas1.create_window(50, 220, window=Exit)
canvas1.create_window(150, 270, window=try_again_button)

root.mainloop()




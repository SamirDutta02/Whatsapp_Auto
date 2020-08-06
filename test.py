import schedule
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog,QTimeEdit
import pandas as pd
import xlrd
from xlrd import XLRDError
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
import autoit

parser = argparse.ArgumentParser(description='PyWhatsapp Guide')
parser.add_argument('--chrome_driver_path', action='store', type=str, default=r'./driver/chromedriver.exe', help='chromedriver executable path')
parser.add_argument('--message', action='store', type=str, default='', help='Enter the msg you want to send')
parser.add_argument('--remove_cache', action='store', type=str, default='False', help='Remove Cache | Scan QR again or Not')
args = parser.parse_args()

if args.remove_cache == 'True':
    os.system('rm -rf User_Data/*')
browser = None
Link = "https://web.whatsapp.com/"
wait = None
element = None
wait_try = None
doc_wait=None
img_wait=None
doc_send = ''
img_send= ''
not_sent_contacts=[]
not_sent_contacts_try=[]
unsent_message=[]
#unsent_message_try=[]
status=[]
message_inp_box= ''
Schedule_msg =''
num = 0
row_head = 0
row_tail = 150
date_time = datetime.datetime.now().strftime("%H.%M.%S")



class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(657, 400)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        #label1
        self.Placeur_msg_label = QtWidgets.QLabel(self.centralwidget)
        self.Placeur_msg_label.setGeometry(QtCore.QRect(10, 10, 221, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.Placeur_msg_label.setFont(font)
        self.Placeur_msg_label.setObjectName("Placeur_msg_label")

        #inpbox
        self.input_text_box = QtWidgets.QTextEdit(self.centralwidget)
        self.input_text_box.setGeometry(QtCore.QRect(10, 30, 631, 161))
        self.input_text_box.setObjectName("input_text_box")
        self.input_text_box.setPlaceholderText('Place Your Text Here')

        #text_submit_btn
        self.text_submit = QtWidgets.QPushButton(self.centralwidget)
        self.text_submit.setGeometry(QtCore.QRect(560, 200, 75, 23))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.text_submit.setFont(font)
        self.text_submit.setObjectName("text_submit")

        

        #excelbtn
        self.Import_excel = QtWidgets.QPushButton(self.centralwidget)
        self.Import_excel.setGeometry(QtCore.QRect(210, 260, 100, 31))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.Import_excel.setFont(font)
        self.Import_excel.setObjectName("Import_excel")

        #send_now_btn
        self.Send_now = QtWidgets.QPushButton(self.centralwidget)
        self.Send_now.setGeometry(QtCore.QRect(530, 320, 91, 41))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.Send_now.setFont(font)
        self.Send_now.setObjectName("Send_now")

        #set time
        self.Submit = QtWidgets.QPushButton(self.centralwidget)
        self.Submit.setGeometry(QtCore.QRect(400, 280, 75, 23))
        self.Submit.setObjectName("Submit")


        #time_widget
        self.timeEdit_wid = QtWidgets.QTimeEdit(self.centralwidget)
        self.timeEdit_wid.setGeometry(QtCore.QRect(400, 245, 118, 22))
        self.timeEdit_wid.setObjectName("timeEdit_wid")
        
        #Schedule_msg_label
        self.Schedule_msg = QtWidgets.QLabel(self.centralwidget)
        self.Schedule_msg.setGeometry(QtCore.QRect(400, 220, 101, 16))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.Schedule_msg.setFont(font)
        self.Schedule_msg.setObjectName("Schedule_msg")

        #img btn
        self.Select_IMG = QtWidgets.QPushButton(self.centralwidget)
        self.Select_IMG.setGeometry(QtCore.QRect(50, 210, 81, 31))
        self.Select_IMG.setObjectName("Select_IMG")

        #doc btn
        self.select_PDF = QtWidgets.QPushButton(self.centralwidget)
        self.select_PDF.setGeometry(QtCore.QRect(50, 280, 81, 31))
        self.select_PDF.setObjectName("select_PDF")



        self.row_head_inp = QtWidgets.QSpinBox(self.centralwidget)
        self.row_head_inp.setGeometry(QtCore.QRect(193, 230, 42, 22))
        self.row_head_inp.setMaximum(999)
        self.row_head_inp.setObjectName("row_head_inp")
        self.row_tail_inp = QtWidgets.QSpinBox(self.centralwidget)
        self.row_tail_inp.setGeometry(QtCore.QRect(282, 230, 42, 22))
        self.row_tail_inp.setMaximum(999)
        self.row_tail_inp.setObjectName("row_tail_inp")
        self.row_tail_inp.setValue(150)

        
        self.Rows = QtWidgets.QLabel(self.centralwidget)
        self.Rows.setGeometry(QtCore.QRect(240, 230, 41, 20))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.Rows.setFont(font)
        self.Rows.setObjectName("Rows")
        
        self.IMG_Selected = QtWidgets.QLabel(self.centralwidget)
        self.IMG_Selected.setGeometry(QtCore.QRect(40, 240, 91, 20))
        self.IMG_Selected.setObjectName("IMG_Selected")
        
        self.Doc_selected = QtWidgets.QLabel(self.centralwidget)
        self.Doc_selected.setGeometry(QtCore.QRect(40, 310, 91, 20))
        self.Doc_selected.setObjectName("Doc_selected")
        
        self.Excel_selected = QtWidgets.QLabel(self.centralwidget)
        self.Excel_selected.setGeometry(QtCore.QRect(209, 290, 91, 20))
        self.Excel_selected.setObjectName("Excel_selected")
        
        self.Task_completed = QtWidgets.QLabel(self.centralwidget)
        self.Task_completed.setGeometry(QtCore.QRect(240, 380, 111, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.Task_completed.setFont(font)
        self.Task_completed.setObjectName("Task_completed")
        
        self.Try_again_btn = QtWidgets.QPushButton(self.centralwidget)
        self.Try_again_btn.setGeometry(QtCore.QRect(444, 340, 71, 23))
        self.Try_again_btn.setObjectName("Try_again_btn")

        self.Message_saved_label = QtWidgets.QLabel(self.centralwidget)
        self.Message_saved_label.setGeometry(QtCore.QRect(563, 225, 81, 20))
        self.Message_saved_label.setObjectName("Message_saved_label")
        self.Time_set = QtWidgets.QLabel(self.centralwidget)
        self.Time_set.setGeometry(QtCore.QRect(410, 299, 51, 20))


        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 657, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Auto Whatsapp"))
        self.Placeur_msg_label.setText(_translate("MainWindow", "Your Message :-"))
        self.Import_excel.setText(_translate("MainWindow", "Import Excel"))
        self.Send_now.setText(_translate("MainWindow", "SEND >"))
        self.Submit.setText(_translate("MainWindow", "Set Time"))
        self.Schedule_msg.setText(_translate("MainWindow", "Schedule Message"))
        self.Select_IMG.setText(_translate("MainWindow", "Select IMG"))
        self.select_PDF.setText(_translate("MainWindow", "Select PDF"))
        self.text_submit.setText(_translate("MainWindow", "Save"))
        self.Rows.setText(_translate("MainWindow", "-Rows-"))
        self.IMG_Selected.setText(_translate("MainWindow", " "))
        self.Doc_selected.setText(_translate("MainWindow", " "))
        self.Excel_selected.setText(_translate("MainWindow", " "))
        self.Task_completed.setText(_translate("MainWindow", ""))
        self.Try_again_btn.setText(_translate("MainWindow", "Try Again"))
        self.Message_saved_label.setText(_translate("MainWindow", ""))
        self.Time_set.setText(_translate("MainWindow", ""))


        #function_mapping
        self.Import_excel.clicked.connect(self.getExcel)
        self.text_submit.clicked.connect(self.getText)
        self.Send_now.clicked.connect(self.auto)
        self.Submit.clicked.connect(self.get_time)
        self.Select_IMG.clicked.connect(self.get_img_name)
        self.select_PDF.clicked.connect(self.get_doc_name)
        self.Try_again_btn.clicked.connect(self.try_again)


    
    def get_time(self):
        global Schedule_msg,send_time,set_time
        set_time = self.timeEdit_wid.time()
        print('Schedule at')
        send_time=(set_time.toString())
        send_time= send_time[:-3]
        print(send_time)
        Schedule_msg='Yes'
        self.Time_set.setText(' Time Set')
        print(Schedule_msg)

    
    def getText(self):
        global message_inp_box
        message_inp_box = self.input_text_box.toPlainText()
        print(message_inp_box)
        self.Message_saved_label.setText('Message Saved')

    def get_img_name(self):
        global imgname, img_send
        img_path_folder = QFileDialog.getOpenFileName()
        img_path = img_path_folder[0]
        print(img_path)
        imgname=os.path.basename(img_path)
        print(imgname)
        self.IMG_Selected.setText("     IMG Selected")
        img_send= 'yes'

    def get_doc_name(self):
        global doc_filename, doc_send
        doc_path_folder = QFileDialog.getOpenFileName()
        doc_path = doc_path_folder[0]
        print(doc_path)
        doc_filename=os.path.basename(doc_path)
        print(doc_filename)
        self.Doc_selected.setText("     DOC Selected")
        doc_send= 'yes'


    def getExcel(self):
        global message_inp_box,unsaved_Contacts,message,raw_user_name
        row_head=self.row_head_inp.value()
        row_tail=self.row_tail_inp.value()
        if (row_tail-row_head > 150):
            self.Excel_selected.setText("Max 150 Rows")
            return 
        
        print('Import Excel')
        path = QFileDialog.getOpenFileName()
        file_path = path[0]
        print(file_path)
        self.Excel_selected.setText("     Excel Selected")

        #read data from excel
        df1 = pd.read_excel(file_path,na_filter= False,dtype=str)
        raw_performa = list(df1.loc[row_head:row_tail, 'Performa Invoice No.'])
        raw_user_name  =list(df1.loc[row_head:row_tail, 'Client Name'])
        raw_unsaved_Contacts = list(df1.loc[row_head:row_tail, 'Contact No.'])
        raw_invoice_date = list(df1.loc[row_head:row_tail, 'Invoice Date'])
        raw_amount = list(df1.loc[row_head:row_tail, 'Amount'])
        raw_due_date = list(df1.loc[row_head:row_tail, 'Due Date'])
        raw_remark = list(df1.loc[row_head:row_tail, 'Remark'])
        raw_message = list(df1.loc[row_head:row_tail, 'Message'])
        
        
        
        
        performa=[]
        for per in raw_performa:
            performa.append(per.replace(' ', '%20').replace('\n', '%0D%0A').replace('&','%26').replace('+', '%2B'))

        user_name=[]
        for un in raw_user_name:
            user_name.append(un.replace(' ', '%20').replace('\n', '%0D%0A').replace('&','%26').replace('+', '%2B'))
        
        unsaved_Contacts=[]
        for names in raw_unsaved_Contacts:
            unsaved_Contacts.append(names.replace(' ', '').replace('+', ''))
                                           
        invoice_date=[]
        for id in raw_invoice_date:
            invoice_date.append(id .replace(' 00:00:00','').replace('\n', '%0D%0A').replace('&','%26').replace('+', '%2B')
                                   .replace(' ', '%20'))
    
        
        amount=[]
        for amt in raw_amount:
            amount.append(amt.replace(' ', '%20').replace('\n', '%0D%0A').replace('&','%26').replace('+', '%2B'))
        
        due_date=[]
        for due in raw_due_date:
            due_date.append(due.replace('\n', '%0D%0A').replace('&','%26').replace('+', '%2B').replace(' 00:00:00','')
                                .replace(' ', '%20'))
        
        remark=[]
        for rem in raw_remark:
            remark.append(rem.replace(' ', '%20').replace('\n', '%0D%0A').replace('&','%26').replace('+', '%2B'))
        
        #report work
        for n in range (0,len(user_name)):
            status.append('pending')
            
        if(message_inp_box== ''):
            message=[]
            for (rm,per,un,id,amt,due,rem) in zip(raw_message,performa,user_name,invoice_date,amount,due_date,remark):
                message.append(rm.replace(' ', '%20').replace('\n', '%0D%0A').replace('+', '%2B').replace('&','%26')
                                    .replace('[Client%20Name]', un).replace('[Performa%20Invoice%20No.]',per)
                                    .replace('[Invoice%20Date]',id).replace('[Amount]',amt).replace('[Due%20Date]',due)
                                    .replace('[Remark]',rem))
                #print(message)

        else:
            msg1 = message_inp_box.replace(' ', '%20').replace('\n', '%0D%0A').replace('&','%26').replace('+', '%2B')

                                    
            message = []
            for (per,un,id,amt,due,rem) in zip(performa,user_name,invoice_date,amount,due_date,remark):
                message.append(msg1.replace('[Performa%20Invoice%20No.]',per).replace('[Client%20Name]', un)
                                    .replace('[Invoice%20Date]',id).replace('[Amount]',amt).replace('[Due%20Date]',due)
                                    .replace('[Remark]',rem))
                #print(message)          



    def try_again(self):
        print("Inside try again")
        global not_sent_contacts_try,unsent_message_try,wait_try

        for (b, u) in zip(not_sent_contacts, unsent_message):
            link = "https://web.whatsapp.com/send?phone={}&text={}".format(b, u)
            print(link)
            browser.get(link)
            try:
                time.sleep(8)
                wait_try = WebDriverWait(browser, 60).until(
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



    def auto(self):
        
        def whatsapp_login(chrome_path):
            global  browser, Link, wait
            chrome_options = Options()
            chrome_options.add_argument('--user-data-dir=./User_Data')
            browser = webdriver.Chrome(executable_path=chrome_path, options=chrome_options)
            browser.get(Link)
            wait = WebDriverWait(browser, 120).until(
                            EC.presence_of_element_located((By.ID, "pane-side")))
            browser.maximize_window()
            print("QR scanned")

        def send_unsaved_contact_message():
            send = browser.find_element_by_xpath("//span[@data-testid='send']")
            send.click()
            return

        def scheduler():
            while True:
                schedule.run_pending()
                time.sleep(1)
        
        def send_img():
            global imgname,img_wait

            clipButton = browser.find_element_by_xpath('//*[@id="main"]/header/div[3]/div/div[2]/div/span')
            clipButton.click()
            time.sleep(1)

            mediaButton = browser.find_element_by_xpath('//*[@id="main"]/header/div[3]/div/div[2]/span/div/div/ul/li[1]/button')
            mediaButton.click()
            time.sleep(3)

            image_path = os.getcwd() +"\\Media\\" + imgname
            print('test of img path')
            print(image_path)

            autoit.control_focus("Open", "Edit1")
            autoit.control_set_text("Open", "Edit1", image_path)
            time.sleep(1)
            autoit.control_click("Open", "Button1")

            time.sleep(5)
            whatsapp_send_button = browser.find_element_by_xpath('//span[@data-testid="send"]')
            img_wait = WebDriverWait(browser, 15).until(
                EC.presence_of_element_located((By.XPATH, "//span[@data-testid='send']")))
            whatsapp_send_button.click()
            time.sleep(1)
        
        
        def send_doc():
            global doc_filename,doc_wait
            clipButton = browser.find_element_by_xpath('//*[@id="main"]/header/div[3]/div/div[2]/div/span')
            clipButton.click()
            time.sleep(1)

            docButton = browser.find_element_by_xpath('//*[@id="main"]/header/div[3]/div/div[2]/span/div/div/ul/li[3]/button')
            docButton.click()
            time.sleep(3)

            
            docPath = os.getcwd() + "\\Media\\" + doc_filename
            print('test of doc path')
            print(docPath)

            autoit.control_focus("Open", "Edit1")
            autoit.control_set_text("Open", "Edit1", docPath)
            time.sleep(1)
            autoit.control_click("Open", "Button1")
            
            time.sleep(4)
            whatsapp_send_button = browser.find_element_by_xpath('//span[@data-testid="send"]')
            doc_wait = WebDriverWait(browser, 15).until(
                EC.presence_of_element_located((By.XPATH, "//span[@data-testid='send']")))
            whatsapp_send_button.click()
            time.sleep(5)
              
        
        def sender():
            print('Sender was called')
            global message,num,element
            num = 0
            for(i,g) in zip(unsaved_Contacts,message):
                    link = "https://web.whatsapp.com/send?phone={}&text={}".format(i,g)
                    print(link)
                    browser.get(link)
                    try:
                        time.sleep(7)
                        if((img_send=='yes') or (doc_send=='yes')):

                            element = WebDriverWait(browser, 65).until(
                                EC.presence_of_element_located((By.XPATH, "//span[@data-testid='send']")))
                            send_unsaved_contact_message()
                            print("Page is ready!")
                        else:
                            element = WebDriverWait(browser, 30).until(
                                EC.presence_of_element_located((By.XPATH, "//span[@data-testid='send']")))
                            send_unsaved_contact_message()
                            print("Page is ready!")


                        if ((img_send=='yes') and (doc_send=='')):
                            print('sending img')
                            send_img()
                            
                        if ((img_send=='') and (doc_send=='yes')):
                            print('sending doc')
                            send_doc()

                        if ((img_send=='yes') and (doc_send=='yes')):
                            send_img()
                            send_doc()
                        else:
                            pass
                    
                        status.pop(num)
                        status.insert(num,'sent')
                        report_folder()
                        num = num + 1
                        time.sleep(2)

                    except:
                        print('Not sent')
                        status.pop(num)
                        status.insert(num,'not_sent')
                        report_folder()
                        num = num + 1
                        not_sent_contacts.append(i)
                        unsent_message.append(g)
                        #print(not_sent_contacts)

                    else:
                        continue
    
        def report_folder():
            def write_in_file():
                #create file & write
                name_of_file = date_time
                path_1 = r"./Whatsapp_Auto_Report/"
                path_2 = str(name_of_file)
                path_3 = '.xlsx'
                report_path = path_1 + path_2 + path_3
                df_rep = pd.DataFrame({'Names':raw_user_name,
                                    'Contact':unsaved_Contacts,
                                    'Status':status})
                df_rep.to_excel(report_path, sheet_name='Whatsapp_Report', index=False)


            try:
                #createfolder
                directory = "Whatsapp_Auto_Report"
                parent_dir = r"./"
                path = os.path.join(parent_dir, directory)
                os.mkdir(path)
                write_in_file()

            except:
                write_in_file()

        def first_func():

            #if __name__ == "__main__":
            print("Web Page Open")
            print("SCAN YOUR QR CODE FOR WHATSAPP WEB")
            whatsapp_login(args.chrome_driver_path)
            
            if(Schedule_msg == 'Yes'):
                
                print('Scheduling')
                schedule.every().day.at(send_time).do(sender)
                scheduler()
                print(not_sent_contacts)
                print("Task Completed")
                #self.Task_completed.setText('Task Completed')

            else:
                sender()
                print(not_sent_contacts)
                print("Task Completed")
                
            
                #self.Task_completed.setText('Task Completed')
        first_func()

if __name__ == "__main__":

    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())



from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QTimeEdit
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

#root = tk.Tk()
#message_var=tk.StringVar(value='Place Your Message Here')

parser = argparse.ArgumentParser(description='PyWhatsapp Guide')
parser.add_argument('--chrome_driver_path', action='store', type=str, default=r'./driver/chromedriver.exe', help='chromedriver executable path')
parser.add_argument('--message', action='store', type=str, default='', help='Enter the msg you want to send')
parser.add_argument('--remove_cache', action='store', type=str, default='False', help='Remove Cache | Scan QR again or Not')
args = parser.parse_args()

if args.remove_cache == 'True':
    os.system('rm -rf User_Data/*')
browser = None
#message = []
Link = "https://web.whatsapp.com/"
wait = None
#unsaved_Contacts = None
choice = None
docChoice = None
doc_filename = None
imgname= None
not_sent_contacts=[]
not_sent_contacts_try=[]
unsent_message=[]
#unsent_message_try=[]
message_inp_box= ''
Schedule_msg =''
date_time = datetime.datetime.now().strftime("%H.%M.%S-%d.%m.%y")



class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(657, 391)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.Placeur_msg_label = QtWidgets.QLabel(self.centralwidget)
        self.Placeur_msg_label.setGeometry(QtCore.QRect(10, 10, 221, 16))
        font = QtGui.QFont()
        font.setPointSize(10)

        #label1
        self.Placeur_msg_label.setFont(font)
        self.Placeur_msg_label.setObjectName("Placeur_msg_label")

        #inpbox
        self.input_text_box = QtWidgets.QTextEdit(self.centralwidget)
        self.input_text_box.setGeometry(QtCore.QRect(10, 30, 631, 161))
        self.input_text_box.setObjectName("input_text_box")

        #text_submit_btn
        self.text_submit = QtWidgets.QPushButton(self.centralwidget)
        self.text_submit.setGeometry(QtCore.QRect(150, 210, 81, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.text_submit.setFont(font)
        self.text_submit.setObjectName("text_submit")

        

        #excelbtn
        self.Import_excel = QtWidgets.QPushButton(self.centralwidget)
        self.Import_excel.setGeometry(QtCore.QRect(270, 200, 100, 41))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.Import_excel.setFont(font)
        self.Import_excel.setObjectName("Import_excel")

        #send_now_btn
        self.Send_now = QtWidgets.QPushButton(self.centralwidget)
        self.Send_now.setGeometry(QtCore.QRect(550, 280, 81, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.Send_now.setFont(font)
        self.Send_now.setObjectName("Send_now")

        #submit_for_time
        self.Submit = QtWidgets.QPushButton(self.centralwidget)
        self.Submit.setGeometry(QtCore.QRect(20, 320, 75, 23))
        self.Submit.setObjectName("Submit")

        #time_widget
        self.timeEdit_wid = QtWidgets.QTimeEdit(self.centralwidget)
        self.timeEdit_wid.setGeometry(QtCore.QRect(20, 290, 118, 22))
        self.timeEdit_wid.setObjectName("timeEdit_wid")
        
        #Schedule_msg_label
        self.Schedule_msg = QtWidgets.QLabel(self.centralwidget)
        self.Schedule_msg.setGeometry(QtCore.QRect(20, 270, 101, 16))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.Schedule_msg.setFont(font)
        self.Schedule_msg.setObjectName("Schedule_msg")

        self.Select_IMG = QtWidgets.QPushButton(self.centralwidget)
        self.Select_IMG.setGeometry(QtCore.QRect(200, 280, 81, 31))
        self.Select_IMG.setObjectName("Select_IMG")

        self.select_PDF = QtWidgets.QPushButton(self.centralwidget)
        self.select_PDF.setGeometry(QtCore.QRect(390, 280, 81, 31))
        self.select_PDF.setObjectName("select_PDF")


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
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.Placeur_msg_label.setText(_translate("MainWindow", "Place Your Message Here"))
        self.Import_excel.setText(_translate("MainWindow", "Import Excel"))
        self.Send_now.setText(_translate("MainWindow", "Send Now >"))
        self.Submit.setText(_translate("MainWindow", "Submit"))
        self.Schedule_msg.setText(_translate("MainWindow", "Schedule Message"))
        self.Select_IMG.setText(_translate("MainWindow", "Select IMG"))
        self.select_PDF.setText(_translate("MainWindow", "Select PDF"))
        self.text_submit.setText(_translate("MainWindow", "Submit Msg"))



        self.Import_excel.clicked.connect(self.getExcel)
        self.text_submit.clicked.connect(self.getText)
        self.Send_now.clicked.connect(self.auto)
        self.Submit.clicked.connect(self.get_time)



    def get_time(self):
        global Schedule_msg
        send_time = self.timeEdit_wid.time()
        print('time is ')
        print(send_time.toString())
        Schedule_msg='Yes'


    
    def getText(self):
        global message_inp_box
        message_inp_box = self.input_text_box.toPlainText()
        print(message_inp_box)





    def getExcel(self):
        global message_inp_box,unsaved_Contacts,message
        
        print('hello')
        path = QFileDialog.getOpenFileName()
        file_path = path[0]
        print(file_path)

        
        df1 = pd.read_excel(file_path, usecols='A')
        df2 = pd.read_excel(file_path, usecols='B')
        df3 = pd.read_excel(file_path, usecols='C')
        raw_user_name=list(df1['Names'])
        unsaved_Contacts = list(df2['Contact'])
        raw_message = list(df3['Message'])
        user_name=[]
        for s in raw_user_name:
            user_name.append(s.replace(' ', '%20').replace('&','%26').replace('?', '%3F'))
        
        if(message_inp_box== ''):
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



    
    def auto(self):

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

        def send_unsaved_contact_message(self):
            send = browser.find_element_by_xpath("//span[@data-testid='send']")
            send.click()
            return



        def scheduler():
            while True:
                schedule.run_pending()
                time.sleep(1)

        def sender():
            print('Sender was called')
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
            print("SCAN YOUR QR CODE FOR WHATSAPP WEB")
            whatsapp_login(args.chrome_driver_path)
        
        if(Schedule_msg == "Yes"):
            schedule.every().day.at(send_time).do(sender)
        else:
            sender()
            scheduler()
            print(not_sent_contacts)
            print("Task Completed")
 

if __name__ == "__main__":

    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
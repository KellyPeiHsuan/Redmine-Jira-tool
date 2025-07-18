#20240221 Tool_v1.3 Darren fixed the issue "Modify_code of Transfer status does not been change when origrnal status is wait"
#20240306 Tool_v1.4 Darren fixed the issue "Leader's comment can not been updated when "Transferred" status!"
#20240306 Tool_v1.4 Darren fixed the issue "Sync new component items of Jira"
#20240801 Tool_v2.0 Minse add function "Add new PIMS to Redmine"　ａｎｄ　fixed the crash caused by search syntax errors.
#20240809 Tool_v2.1 Minse 修復閃退、添加ｄｕｅ　ｄａｔｅ和進度條提醒
#20240815 Tool_v2.2 Minse 修復進度條提醒造成閃退問題
#20240823 Tool_v2.3 Minse 修正日期時區
#20240904 Tool_v2.4 Minse Jira憑證問題和Redmine Assigned添加 Assignee to jira Analyst 
#20241227 Tool_v2.4.1 Aslan添加下commond自動start功能 
#20241230 Tool_v2.4.2 Aslan添加下commond可選checkbox功能
#20250103 Tool_v2.4.3 Aslan添加下commond可秀cmd list功能
#20250107 Tool_v2.4.4 Aslan添加視窗自動RUN完可自動關閉功能
#20250107 Tool_v2.4.5 Aslan添加在LOG底部會秀視窗已經RUN完且自動關閉字串
#20250314 Tool_v2.4.6 Aslan添加新格子Solution Category for PQM
#20250408 Tool_v2.4.7 Aslan添加判斷Redmine Jira哪個連線FAIL機制
#20250612 Tool_v2.4.7 Aslan添加tool run 完自動寄信功能
#20250707 Kelly update for EC to Use
from redminelib import Redmine
import jira
import sys
from PyQt5.QtGui import QColor, QPixmap, QFont
from PyQt5.QtWidgets import QMainWindow, QWidget, QApplication, QListWidget, QPushButton, QGridLayout, QLabel, QProgressBar, QCheckBox, QVBoxLayout
from PyQt5.Qt import pyqtSignal
from PyQt5.QtCore import QThread, Qt, QRect
import configparser
import datetime
import os
import pandas as pd
import re
import configparser
import io
import requests
import time
import json
import pytz
import urllib3
import smtplib
import csv
from FunctionDetermine import RedmineOwner
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)
sys.stderr = open(os.devnull, 'w')

# 禁用 InsecureRequestWarning 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class MyWidget(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_GUI()

    def init_GUI(self):
        self.setWindowTitle('Redmine & Jira Sync Tool_v2.4.8')
        
        # 设置窗口大小
        self.screenRect = QApplication.desktop().screenGeometry()
        self.height = self.screenRect.height()
        self.width = self.screenRect.width()
        self.resize(800, 600)
        
        # 使用者資訊網格
        self.info_widget = QWidget()  # 創造function物件
        self.info_widget.setObjectName('itemwidget')
        self.info_layout = QGridLayout()  
        self.info_widget.setLayout(self.info_layout) 

        # 使用者資訊物件
        self.user_info_icon = QLabel(self)
        self.user_info_icon.setObjectName('label_1')
        self.user_info_icon.setPixmap(QPixmap(r'.\lib\image\icon_user.png'))
        self.user_info_icon.setMaximumHeight(150)  # 设置图片最大高度
        self.user_info_icon.setMaximumWidth(150)  # 设置图片最大寬度
        self.user_info_icon.setAlignment(Qt.AlignCenter)
        self.user_info_icon.setAutoFillBackground(True)
        self.user_info_icon.setScaledContents(True)

        # 读取配置文件
        config = configparser.ConfigParser()
        new_path = os.getcwd()
        config.read(new_path + "\\Parameter.ini")
        weburl = config.get("Web", "Redmine_url")
        project_id_0 = config.get("Filter", "Project")
        project_id_1 = config.get("Filter", "Project_1")
        project_id_2 = config.get("Filter", "Project_2")
        project_id = ''
        if project_id_0:
            project_id = project_id + project_id_0
        if project_id_1:
            project_id = project_id + ', ' + project_id_1
        if project_id_2:
            project_id = project_id + ', ' + project_id_2
        jira_account = config.get("Jira", "Account")
        jira_email = config.get("Jira", "Email")
        jira_web = config.get("Web", "Jira_url")
        jira_verify_str = config.get("Web", "Jira_verify")
        redmine_account = config.get("Redmine", "Account")
        
        # 设置显示标签
        self.web_url = QLabel("Redmine URL : " + weburl)
        self.project_name = QLabel("Project : " + project_id)
        self.jiraweb_url = QLabel("Jira URL : " + jira_web + " ; SSL Verification : " + jira_verify_str)
        self.Jira_Account = QLabel("Jira Account : " + jira_account)
        self.Jira_Email = QLabel("Jira Email : " + jira_email)
        self.Redmine_Account = QLabel("Redmine Account : " + redmine_account)

        # 功能列表網格
        self.function_widget = QWidget()
        self.function_widget.setObjectName('itemwidget')
        self.function_layout = QGridLayout()  
        self.function_widget.setLayout(self.function_layout)

        # 功能列表物件
        self.function = QLabel('Select Features', self)
        self.function.setObjectName('title')
        self.checkbox = QCheckBox('   Add new PIMS to Redmine', self)
        self.checkbox.setChecked(False)
        self.checkbox2 = QCheckBox('   Upload PIMS information of Redmine to Jira', self)
        self.checkbox2.setChecked(False)

        # 根据命令行参数设置复选框状态
        #self.set_checkboxes_based_on_cmd()
        '''
    def set_checkboxes_based_on_cmd(self):
        # 获取命令行参数并设置复选框状态
        if len(sys.argv) > 1:  # 确保有命令行参数
            command = sys.argv[1]

            if command == 's1':
                self.checkbox.setChecked(True)  
                self.checkbox2.setChecked(False)
            elif command == 's2':
                self.checkbox.setChecked(False)
                self.checkbox2.setChecked(True)  
            elif command == 'sa':
                self.checkbox.setChecked(True)  
                self.checkbox2.setChecked(True)
            elif  command =='-l':
                        print("Please check out the following command prompts::")
                        print("s1 -> adds a new PIMS to Redmine")
                        print("s2 -> uploads the PIMS information to Jira")
                        print("sa -> Both are performed")
                        sys.exit(0)
            else:
                print("Please enter -l to view the command prompt")
                self.print_commands()
                sys.exit(0)
        '''      


        # 訊息框
        self.textbox = QListWidget(self)
        self.textbox.setFixedHeight(300)
        self.textbox.addItem(" Message") 
        font = QFont()
        font.setPointSize(14)
        font.setBold(True)
        font.setFamily('Calibri')
        self.textbox.item(0).setFont(font)

        # 進度條
        self.progressBar = QProgressBar()
        self.progressBar.setGeometry(QRect(100, 80, 118, 23))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        
        # 功能按鈕
        self.btnStart = QPushButton('Start',self)
        self.btn_over = QPushButton('Stop',self)
        # 功能連結
        self.btnStart.clicked.connect(self.execute_action)
        self.btn_over.clicked.connect(self.slot_btn_over)
        
        # 圖片
        self.se_icon = QLabel(self)
        self.se_icon.setObjectName('label_1')
        self.se_icon.setPixmap(QPixmap(r'.\lib\image\logo.png'))
        self.se_icon.setFixedSize(600, 200)  # 設置圖標大小
        self.se_icon.setAlignment(Qt.AlignCenter)
        self.se_icon.setAutoFillBackground(True)
        self.se_icon.setScaledContents(True)
        
        # 圖片容器
        self.se_icon_container = QWidget(self)
        self.se_icon_layout = QVBoxLayout(self.se_icon_container)
        self.se_icon_layout.addWidget(self.se_icon)
        self.se_icon_layout.setAlignment(Qt.AlignCenter)
        
        # 把控件放置在栅格布局中
        self.main_widget = QWidget() #創造窗口
        self.main_layout = QGridLayout() #創造物件布局
        self.main_widget.setObjectName('main_widget')
        self.main_widget.setLayout(self.main_layout) #設置窗口主物件布局做為網格
        self.setCentralWidget(self.main_widget) # 设置窗口主部件
        
        self.main_layout.addWidget(self.info_widget, 0, 0, 1, 2)
        self.info_layout.addWidget(self.user_info_icon, 0, 0, 4, 1)
        self.info_layout.addWidget(self.web_url, 0, 1, 1, 1)
        self.info_layout.addWidget(self.project_name, 1, 1, 1, 1)
        self.info_layout.addWidget(self.jiraweb_url, 2, 1, 1, 1)
        self.info_layout.addWidget(self.Jira_Account, 3, 1, 1, 1)
        self.info_layout.addWidget(self.Jira_Email, 4, 1, 1, 1)
        self.info_layout.addWidget(self.Redmine_Account, 5, 1, 1, 1)
        self.info_layout.setHorizontalSpacing(40)
        
        self.main_layout.addWidget(self.function_widget, 1, 0, 1, 2)
        self.function_layout.addWidget(self.function, 0, 0)
        self.function_layout.addWidget(self.checkbox, 1, 0)
        self.function_layout.addWidget(self.checkbox2, 2, 0)
        
        self.main_layout.addWidget(self.textbox, 2, 0, 1, 2)
        self.main_layout.addWidget(self.progressBar,3, 0, 1, 2)
        self.main_layout.addWidget(self.btnStart, 4, 0)
        self.main_layout.addWidget(self.btn_over, 4, 1)
        self.main_layout.addWidget(self.se_icon_container, 5, 0, 1, 2)
      
        
        #self.execute_action()  # 自动执行 Start 按钮的操作

        
        # log
        self.update_log_path()
        # CSS
        self.main_widget.setStyleSheet('''
            QWidget#main_widget{
                background:#D2E9FF;
                }
            QLabel#label_1{
                background:#FFFFFF; 
                font-family: Calibri;
                }                      
            ''')  
        self.info_widget.setStyleSheet('''
            QWidget{
                border-radius:1px;
                }
            QWidget#itemwidget{
                background:#FFFFFF;
                border-top:1px solid white;
                border-bottom:1px solid white;
                border:1px solid white;
                border-top-left-radius:10px;
                border-radius:10px;
                font-family: Calibri;
                }
            QLabel#title{
                font-weight: bold;
                font-size: 30px;  
                font-family: Calibri;
                }
            QLabel{
                font-weight: bold;
                font-size: 20px;    
                font-family: Calibri;
                }
            ''')
        self.function_widget.setStyleSheet('''
            QWidget{
                border-radius:1px;
                }
            QWidget#itemwidget{
                background:#ffffff;
                border-top:1px solid white;
                border-bottom:1px solid white;
                border:1px solid white;
                border-top-left-radius:10px;
                border-radius:10px;
                font-family: Calibri;
                }
            QLabel#title{
                font-family: Calibri;
                font-weight: bold;
                font-size: 35px;          
                }
            QCheckBox{
                font-family: Calibri;
                font-size: 20px;          
                }
            ''')
        self.btnStart.setStyleSheet('''
           QPushButton{
               background: #FFFAFA;
               border-radius:10px;
               border-style:ridge;
               border-width:2px;
               border-color:#00000;
               color: #000000;
               font-size: 30px;
               font-family: Calibri;
               }
           QPushButton:hover{
               background: #EEE9E9;
               }
           ''')
        self.btn_over.setStyleSheet('''
            QPushButton{
                background: #FFFAFA;
                border-radius:10px;
                border-style:ridge;
                border-width:2px;
                border-color:#00000;
                color: #000000;
                font-size: 30px;
                font-family: Calibri;
                }
            QPushButton:hover{
                background: #EEE9E9;
                }
            ''')
        self.textbox.setStyleSheet('''
            QListWidget{
                border-radius:30px;
                background: #ffffff;
                border-style:ridge;
                border-width:5px;
                border-color:#ffffff;
                font-size: 20px;
                font-family: Calibri;
                }
            QScrollBar{
                width:2px;
                border-width:1px;
                }
            QScrollBar:vertical{   
                border-radius:2px;
                border-style:transparent;
                }
            QScrollBar::add-page:vertical{background-color:transparent;}
            QScrollBar::sub-page:vertical{background-color:transparent;}
            QScrollBar::sub-line:vertical{background-color:transparent;}
            QScrollBar::add-line:vertical{background-color:transparent;}
            QScrollBar::handle:vertical{background:#000000;}
            
            ''')
        self.progressBar.setStyleSheet('''
                border: 2px solid #000;
                text-align:center;
                font-family:Calibri;
                background:#aaa;
                color:#fff;
                height: 15px;
                border-radius: 10px;
                width:150px;
                font-size: 20px; 
                QProgressBar::chunk {
                    background: #333;
                    width:1px;
            }
            ''')
        
    # 生成log
    def update_log_path(self):
        date_now = datetime.datetime.now()
        self.today = date_now.strftime("%Y_%m_%d_%H_%M_%S")
        #上一層資訊
        #self.new_path = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
        self.new_path = os.getcwd()
        self.log_path = self.new_path + '\\Log\\' + str(self.today) + '.txt'
    # 生成Error log
    def error_log(self):
        with open(self.log_path, 'r') as file:
            log_content = file.read()
        # 按 '------------------------------------------------' 分割成多個部分
        log_entries = log_content.split('------------------------------------------------')

        # 找出包含 'Error message' 的部分
        error_entries = [entry.strip() for entry in log_entries if 'Error message' in entry]

        # 排序的 key函數，提取Category的值
        def get_category(record):
            match = re.search(r"Category : (\w+)", record)
            return match.group(1) if match else ''

        # 根据 Category 的值排序
        sorted_records = sorted(error_entries, key=get_category)
        
        # 將找到的 'Error message' 部分寫入新的文本檔
        with open(self.new_path + '\\Log\\' + self.today + '_error_messages.txt', 'w') as error_file:
            for entry in sorted_records:
                error_file.write(entry + '\n' + '-' * 50 + '\n')
    # 提示           
    def pims_display(self,str):
    # 由於自定義訊號時自動傳遞一個字串引數，所以在這個槽函式中要接受一個引數
        self.textbox.addItem(str)
        self.textbox.scrollToBottom()
        # log
        file=open(self.log_path,"a")
        if str == '------------------------------------------------' or str == '================================================':
            file.write( str + '\n')
        else:
            time_now = datetime.datetime.now()
            self.time_out = time_now.strftime("%Y/%m/%d")
            file.write('['+self.time_out+'] '+ str + '\n')
        file.close()
    # 成功提示    
    def display(self,str):
    # 由於自定義訊號時自動傳遞一個字串引數，所以在這個槽函式中要接受一個引數
        self.textbox.addItem(str)
        index = int(len(self.textbox))-1
        self.textbox.item(index).setForeground(QColor("green"))
        self.textbox.scrollToBottom()
        # log
        file=open(self.log_path,"a")
        time_now = datetime.datetime.now()
        self.time_out = time_now.strftime("%Y/%m/%d %H:%M:%S")
        file.write('['+self.time_out+'] '+ str + '\n')
        file.close()
    # 錯誤提示   
    def display_Error(self,str):
        self.textbox.addItem(str)
        index = int(len(self.textbox))-1
        self.textbox.item(index).setForeground(QColor("red"))
        self.textbox.scrollToBottom()
        # log
        file = open(self.log_path,"a")
        time_now = datetime.datetime.now()
        self.time_out = time_now.strftime("%Y/%m/%d %H:%M:%S")
        file.write('['+self.time_out+'] '+ str + '\n')
        file.close()
    # 警告提示
    def display_wait_connection(self,str):
    # 由於自定義訊號時自動傳遞一個字串引數，所以在這個槽函式中要接受一個引數
        self.textbox.addItem(str)
        index = int(len(self.textbox))-1
        self.textbox.item(index).setForeground(QColor("orange"))
        self.textbox.scrollToBottom() 
    #  Add new PIMS to Redmine
    def comply_work(self):
        self.work = Work()
        self.work.trigger_pims.connect(self.pims_display)
        self.work.trigger.connect(self.display)
        self.work.trigger_fail.connect(self.display_Error)
        self.work.trigger_wait.connect(self.display_wait_connection)
        self.work.countChanged.connect(self.onCountChanged)
        #self.work.trigger_bt_on.connect(self.start_bt_on) # 解鎖功能
        self.work.start()
    # Upload PIMS information of Redmine to Jira
    def comply_work2(self):
        self.work2 = Work2()
        self.work2.trigger_pims.connect(self.pims_display)
        self.work2.trigger.connect(self.display)
        self.work2.trigger_fail.connect(self.display_Error)
        self.work2.trigger_wait.connect(self.display_wait_connection)
        self.work2.countChanged.connect(self.onCountChanged)
        self.work2.trigger_bt_on.connect(self.start_bt_on) # 解鎖功能
        self.work2.start()
    # 進度條
    def onCountChanged(self, value):
        self.progressBar.setValue(value)
    # 執行按鈕
    def execute_action(self):
        self.progressBar.setValue(0)
        self.btnStart.setEnabled(False)
        self.update_log_path()  # 生成新的 log_path
        if self.checkbox.isChecked() and self.checkbox2.isChecked():
            self.comply_work()  # 先執行第一個
            self.work.finished.connect(self.comply_work2)  # 完成後執行第二個
        elif self.checkbox.isChecked():
            self.comply_work()
        elif self.checkbox2.isChecked():
            self.comply_work2()
    # 解鎖Start鍵  setEnabled(True)       
    def start_bt_on(self): 
        self.btnStart.setEnabled(True)
        self.error_log()
        #send_email()   
    # 停止按鈕
    def slot_btn_over(self):
        try:
            self.work.terminate()
        except:
            pass
        try:
            self.work2.terminate()
        except:
            pass
        self.btnStart.setEnabled(True)
        self.textbox.addItem('Stop')
        # log
        file=open(self.log_path,"a")
        time_now = datetime.datetime.now()
        self.time_out = time_now.strftime("%Y/%m/%d")
        file.write('['+self.time_out+'] '+ 'Stop\n')
        file.close()
        self.error_log()
        


if getattr(sys, 'frozen', False):
    base_path = os.path.dirname(sys.executable)
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

log_folder_path = os.path.join(base_path, "Log")

parameter = configparser.ConfigParser()
parameter.read(os.path.join(base_path, "Parameter.ini"))

sender_email = parameter.get("MAIL", "sender")
password = parameter.get("MAIL", "password")
receiver_email = parameter.get("MAIL", "receiver")
receiver_list = [email.strip() for email in receiver_email.split(",")]

def get_latest_log_files(folder_path, n=2):
    try:
        files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
        files_sorted = sorted(files, key=os.path.getmtime, reverse=True)
        return files_sorted[:n]
    except Exception as e:
        print(f"找尋最新檔案失敗：{e}")
        return []

def send_email():
    latest_files = get_latest_log_files(log_folder_path, 2)
    if not latest_files:
        print("找不到最新的 log 檔案，停止寄信。")
        return

    message = MIMEMultipart()
    message["Subject"] = "Jira & Redmine Tool execution completed - log & Error log"
    message["From"] = sender_email
    message["To"] = receiver_email

    body = MIMEText("此為Tool自動通知信請勿直接回覆此mail", "plain")
    message.attach(body)

    for filepath in latest_files:
        try:
            with open(filepath, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f'attachment; filename="{os.path.basename(filepath)}"',
            )
            message.attach(part)
        except Exception as e:
            print(f"附加檔案失敗 {os.path.basename(filepath)}：{e}")

    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_list, message.as_string())
        print("郵件已成功寄出！")
    except Exception as e:
        print(f"寄送郵件失敗：{e}")
class Work(QThread):
    trigger_pims = pyqtSignal(str)
    trigger = pyqtSignal(str)
    trigger_fail = pyqtSignal(str)
    trigger_wait = pyqtSignal(str)
    countChanged = pyqtSignal(int)
    trigger_bt_on = pyqtSignal()
    def __int__(self):
        # 初始化函数
        super(Work, self).__init__()
       
    def run(self):
        self.trigger.emit('Start executing Add new PIMS to Redmine')
        #載入帳號密碼和Redmine網址
        config = configparser.ConfigParser()
        #上一層資訊
        #new_path = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
        new_path = os.getcwd()
        config.read(new_path + "\\Parameter.ini")
        project_id = config.get("Filter", "Project")
        #project_id_1 = config.get("Filter", "Project_1")
        #project_id_2 = config.get("Filter", "Project_2")
        project_id_filter = config.get("Jira Filter", "Project")
        #project_id_1_filter = config.get("Jira Filter", "Project_1")
        #project_id_2_filter = config.get("Jira Filter", "Project_2")
        jira_web = config.get("Web", "Jira_url")
        jira_verify_str = config.get("Web", "Jira_verify")
        str_to_bool = {'true': True, 'false': False}
        jira_verify = str_to_bool[jira_verify_str.lower()]
        jira_account = config.get("Jira", "Account")
        jira_password = config.get("Jira", "Password")
        redmine_account = config.get("Redmine", "Account")
        redmine_password = config.get("Redmine", "Password")
        redmine_web = config.get("Web", "Redmine_url")
        taipei_tz = pytz.timezone('Asia/Taipei')
        # 各別project對應的篩選語法, only X17 on BIOS Redmine
        project_filter_list = [[project_id, project_id_filter]]
        try:
            # 登入Jira
            options = {'server': jira_web, 'verify': jira_verify}
            jira_client = jira.JIRA(options=options, basic_auth=(str(jira_account),str(jira_password)))
            connection = 1
        except:
            connection = 0
            self.trigger_wait.emit("jira fail")
            self.trigger_bt_on.emit()
        try:
            # 登入redmine
            redmine = Redmine(str(redmine_web), username=str(redmine_account), password=str(redmine_password))
            projects = redmine.project.all()
            len(projects) # 判斷redmine是否有連接成功
            connection = 1
        except:
            connection = 0
            self.trigger_wait.emit("redmind fail")
            self.trigger_bt_on.emit()
        #宣告進度條0%
        count_bar = 0
        if connection == 1:
            for i in range(0,len(project_filter_list)):
                # 如果Filter和Jiar filter有填寫
                if project_filter_list[i][0] != '' and project_filter_list[i][1] != '':
                    # 顯示目前執行Project
                    self.trigger_pims.emit("------------------------------------------------")
                    self.trigger.emit('Check the ' + project_filter_list[i][0])
                    self.trigger_pims.emit("------------------------------------------------")
                    # 查詢JIRA語法
                    jql_query = project_filter_list[i][1]
                    self.trigger.emit('Searching with filters...')
                    #print('Searching with filters...')
                    # 執行語法, 最多搜索到1000筆資料
                    try:
                        jira_issues_search = jira_client.search_issues(jql_query, maxResults=1000)
                        self.trigger.emit(f'There are a total of {len(jira_issues_search)} records, starting comparison.')
                        self.trigger_pims.emit("------------------------------------------------")
                        search_pass = 1
                    except:
                        self.trigger_fail.emit('Error message : The '+ project_filter_list[i][0] +' query syntax is incorrect. Please verify.')
                        search_pass = 0
                    #print(f'There are a total of {len(jira_issues_search)} records, starting comparison.')
                    if search_pass == 1:
                        # bar 進度條
                        if len(jira_issues_search) != 0:
                            count_pims = 100 / int(len(jira_issues_search))
                        count_bar = 0
                        # 逐條查詢是否在Redmine是否有資訊
                        for issue in jira_issues_search:
                            attempts = 0
                            retry_attempts = 10
                            while attempts < retry_attempts:
                                try:
                                    # 篩選器
                                    search_results = redmine.issue.filter(
                                                status_id='!*', # 所有狀態
                                                cf_1=issue.key  # PIMS查詢
                                            )
                                    attempts += 11
                                # 連線失敗再重新來過
                                except ConnectionError:
                                    attempts += 1
                                    self.trigger_wait.emit(f"ConnectionError encountered. Retrying {attempts}/{retry_attempts}...")
                                    time.sleep(10)
                            # 如果Redmine搜尋不到資料，開始添加ｉｄ動作
                            if len(search_results) == 0:
                                # 初始化 id 建立許可
                                create_id = 1
                                self.trigger_pims.emit(f"Issue Key: {issue.key}")
                                #print(f"Issue Key: {issue.key}")
                                pims_issue = jira_client.issue(issue.key)
                                
                                # Issue Status
                                pims_status = pims_issue.fields.status 
                                
                                # Issue Summary
                                pims_summary = pims_issue.fields.summary
                                
                                # Issue Type
                                try:
                                    pims_issue_type = pims_issue.fields.issuetype.name
                                    if pims_issue_type == 'Defect':
                                        pims_issue_type_value = 4
                                    elif pims_issue_type == 'Change Request':
                                        pims_issue_type_value = 3
                                except:
                                    self.trigger_fail.emit('Add New PIMS Error message : The Issue Type field is empty.')
                                    #print('Error message : The Issue Type field is empty.')
                                    create_id = 0
                                
                                # Issue Severity
                                try:
                                    pims_issue_severity = pims_issue.fields.customfield_28162
                                except:
                                    self.trigger_fail.emit('Add New PIMS Error message : The Issue Severity field is empty.')
                                    #print('Error message : The Issue Severity field is empty.')
                                    create_id = 0
                                
                                # Issue Subsystem (Component / Subcomponent:)
                                try:
                                    pims_issue_subsystem = pims_issue.fields.customfield_28436
                                    pims_issue_subsystem_value = ' - '.join(item.value.lstrip() for item in pims_issue_subsystem)
                                except:
                                    self.trigger_fail.emit('Add New PIMS Error message : The Component / Subcomponent field is empty.')
                                    #print('Error message : The Component / Subcomponent field is empty.')
                                    create_id = 0
                                
                                # Submitted Date
                                try:
                                    pims_submitted_date = pims_issue.fields.customfield_28429
                                    date_obj = datetime.datetime.strptime(pims_submitted_date, '%Y-%m-%dT%H:%M:%S.%f%z')
                                    taipei_time = date_obj.astimezone(taipei_tz)
                                    submitted_date = taipei_time.strftime('%Y-%m-%d')
                                    due_date_14day = (taipei_time + datetime.timedelta(days=14)).strftime('%Y-%m-%d')
                                except:
                                    self.trigger_fail.emit('Add New PIMS Error message : Submitted Date field is empty.')
                                    #print('Error message : Submitted Date field is empty.')
                                    create_id = 0
                                
                               # Platform Found (PF):
                                try:
                                    pims_platform_found = pims_issue.fields.customfield_28470
                                      # 字串轉json
                                    pims_list = json.loads(pims_platform_found)
                                    # json拿出value轉str
                                    pims_platform_found_result = ', '.join(item['value'] for item in pims_list)
                                except:
                                    self.trigger_fail.emit('Add New PIMS Error message : Platform Found (PF) field is empty.')
                                    #print('Error message : Platform Found (PF) field is empty.')
                                    create_id = 0
                                
                                # 如果資料都齊全就將id建立     
                                if create_id == 1:
                                    try:
                                        # 建立新id
                                        new_id = redmine.issue.create(
                                         project_id = project_filter_list[i][0],  # 替換project的项目 ID
                                         subject = pims_summary, # 標題
                                         tracker_id = 10,  # 替換Tracker ID : PIMS/BITS (trackers = redmine.tracker.all())
                                         status_id = 1,  # 替換狀態 ID : New
                                         priority_id = 2,  # 替換優先級 ID : Medium (redmine.enumeration.filter(resource='issue_priorities'))
                                         #assigned_to_id = 7, # assigned : A31_Assigee System
                                         start_date = submitted_date, # start日期
                                         due_date = due_date_14day, # due日期
                                         custom_fields=[
                                             {'id': 122, 'value': issue.key},   #PSR Number
                                             {'id': 127, 'value': 'Automatic_System'}, # utomatic_System
                                             {'id': 121, 'value': pims_issue_subsystem_value}, #PQM info
                                             {'id': 123, 'value': pims_platform_found_result}, 
                                             {'id': 124, 'value': pims_issue_type_value}, #A31 Category
                                         ])
                                        self.trigger.emit(f"Created Issue ID: {new_id.id}")
                                        #print(f"Created Issue ID: {new_id.id}")
                                        time.sleep(1)
                                    except:
                                        self.trigger_fail.emit(f"Add New PIMS Error message : There was an error during creation. Please manually add {issue.key} to Redmine.")
                                        #print(f"Error message : There was an error during creation. Please manually add {issue.key} to Redmine.")
                                        time.sleep(1)
                                    self.trigger_pims.emit("------------------------------------------------")
                            count_bar += count_pims
                            self.countChanged.emit(int(round(count_bar,0)))
                    if count_bar >= 99 and i + 1 < len(project_filter_list) and (project_filter_list[i+1][0] != '' and project_filter_list[i+1][1] != ''):
                        self.countChanged.emit(int(round(0,0)))
                        self.trigger.emit('The ' + project_filter_list[i][0] + ' has been completed. Switching to the ' + project_filter_list[i+1][0] + ". Please don't close window.")
                    elif count_bar >= 99 and i == 0  and (project_filter_list[i+2][0] != '' and project_filter_list[i+2][1] != ''):
                        self.countChanged.emit(int(round(0,0)))
                        self.trigger.emit('The ' + project_filter_list[i][0] + ' has been completed. Switching to the ' + project_filter_list[i+2][0] + ". Please don't close window.")
                    elif count_bar >= 99:
                        self.trigger.emit('The ' + project_filter_list[i][0] + ' has been completed.')
             
            self.trigger_bt_on.emit()
            self.trigger.emit('Execution completed !')
            
# Upload EC PIMS information of Redmine to Jira
class Work2(QThread):
    trigger_pims = pyqtSignal(str)
    trigger = pyqtSignal(str)
    trigger_fail = pyqtSignal(str)
    trigger_wait = pyqtSignal(str)
    countChanged = pyqtSignal(int)
    trigger_bt_on = pyqtSignal()
    def __int__(self):
        # 初始化函数
        super(Work2, self).__init__()
       
    def run(self):
        self.countChanged.emit(int(round(0,0)))
        self.trigger_pims.emit("================================================")
        self.trigger_pims.emit("================================================")
        self.trigger.emit('Start executing Upload PIMS information of Redmine to Jira')
        #載入帳號密碼和Redmine網址
        config = configparser.ConfigParser()
        #上一層資訊
        #new_path = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
        new_path = os.getcwd()
        config.read(new_path + "\\Parameter.ini")
        #project_id = config.get("Filter", "Project")
        jira_web = config.get("Web", "Jira_url")
        jira_verify_str = config.get("Web", "Jira_verify")
        str_to_bool = {'true': True, 'false': False}
        jira_verify = str_to_bool[jira_verify_str.lower()]
        jira_account = config.get("Jira", "Account")
        jira_password = config.get("Jira", "Password")
        jira_email = config.get("Jira", "Email")
        redmine_account = config.get("Redmine", "Account")
        redmine_password = config.get("Redmine", "Password")
        redmine_web = config.get("Web", "Redmine_url")
        ec_member_ids = list(RedmineOwner.EC_Member.values())
        #bios_member_ids = list(RedmineOwner.BIOS_Member.values())
        project_gen = config.get("Platform","X17_platform")

        try:
            
            # 登入Jira
            options = {'server': jira_web, 'verify': jira_verify}
            jira_client = jira.JIRA(options=options, basic_auth=(str(jira_account),str(jira_password)))
            #當前帳號名稱
            #jira_account_email = jira_account + '@compal.com'
            account_user_info = jira_client.search_users(jira_email)[0].name
            self.trigger.emit(f'Log in Jira successfully,current user is {account_user_info}.')
            
            # 登入redmine
            redmine = Redmine(str(redmine_web), username=str(redmine_account), password=str(redmine_password))
            self.trigger.emit('Log in Redmine successfully.')
            # project數量限制         
            ''' redmine = Redmine(str(redmine_web), username=str(redmine_account), password=str(redmine_password))
            issue_statuses = redmine.issue_status.all()
            
           #BIOS Redmine
            Status ID: 1, Status Name: New
            Status ID: 2, Status Name: Open
            Status ID: 3, Status Name: Resolved
            Status ID: 5, Status Name: Closed
            Status ID: 6, Status Name: Waive
            Status ID: 7, Status Name: Pending
            Status ID: 8, Status Name: Tracking
            Status ID: 11, Status Name: Waiting
            Status ID: 12, Status Name: Verify
            Status ID: 13, Status Name: Tranferred
            Status ID: 14, Status Name: Transfferred-Closed
            Status ID: 15, Status Name: ATS/WAD - Can
            Status ID: 16, Status Name: ATS P1/P2/P3
            Status ID: 17, Status Name: ATS P4/P5
            Status ID: 18, Status Name: ATS P1/P2/P3 - Verify
            Status ID: 19, Status Name: ATS P1/P2/P3 - Closed
            Status ID: 20, Status Name: Review
            '''
            statuses = [1,2,11,3,20,12,13,15,18,19]  #狀態標識符列表
            issues = []
            ec_member_ids = list(RedmineOwner.EC_Member.values())
            ec_member = '|'.join(str(i) for i in ec_member_ids)
            
            tracker_id = 10  #BIOS Redmine Tracker ID: 10, Tracker Name: BITS/PIMS
            project_id = 166 #Project ID: 166, Project Name: A31_BC_Projects
            
            for status in statuses:
                filter_params = {
                    'project_id': project_id,
                    'status_id': status,
                    'tracker_id': tracker_id,
                    'cf_123': project_gen,         # X17的Platform,
                    'cf_133': 1668           # EC Member   
                }
            
                filtered_issues = redmine.issue.filter(**filter_params)
                issues.extend(filtered_issues)
            
            '''
            #Check if the issue list is what we needed.
            #先輸出CSV檔案確認撈出的資料
            with open('filtered_issues.csv', 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['ID', 'Subject', 'Status', 'Priority'])  # 欄位名稱
                for issue in issues:
                    writer.writerow([
                        issue.id,
                        issue.subject,
                        issue.status.name if hasattr(issue, 'status') else '',
                        issue.priority.name if hasattr(issue, 'priority') else ''
                    ])
            self.trigger.emit("已匯出 filtered_issues.csv，可用 Excel 檢查資料。")
            
            '''
            connect = 1    
        except:                
            self.trigger_wait.emit("Please check your network connection or verify that your account credentials are correct.")
            self.trigger_bt_on.emit()
            #print("Please check your network connection or verify that your account credentials are correct.")
        
        if connect == 1:
            # bar 進度條
            count_pims = 100 / int(len(issues))
            count_bar = 0
            
            # DT select 列表
            option_mapping = {
                '': '-1',
                'Deferred Candidate': '90054',
                'Existing Fix Provided': '90055',
                'Existing Info Provided': '90056',
                'Fixed - Build Environment Change': '90057',
                'Fixed - Code Change': '90058',
                'Fixed - Documentation Change': '90059',
                'Fixed - Duplicate Root Cause': '90060',
                'Fixed - Hardware Change': '90061',
                'Fixed - Specification Change': '90062',
                'Invalid - Duplicate Issue in Tool': '90063',
                'Invalid - Expected Behavior': '90064',
                'Invalid - Invalid Config': '90065',
                'Invalid - Invalid Test Case': '90066',
                'Invalid - User Misunderstanding': '90067',
                'No Fix - Can Not Duplicate (CND)': '90068',
                'No Fix - Platform EOL': '90069',
                'No Fix - Single Unit Failure': '90070',
                'Obsolete - Business Reason': '90071',
                'Obsolete - Project Cancelled': '90072',
                'Obsolete - Technology Not Supported': '90073',
                'Fixed- Duplicate Root Cause': '132312',
                'Fixed-Code Change': '132314',
                'No Fix- CND': '132315',
                'Fixed- Hardware Change': '132316',
                'Fixed- Build Environment Change': '132317',
                'No Fix- Single Unit Failure': '132318',
                'Fixed- Documentation Change': '132322'
            }
            
            #Solution
            solution_mapping={
            'Add-in Cards': '171611',
            'Audio': '171258',
            'Battery': '171276',
            'BIOS': '171283',
            'Camera': '171749',
            'Catalog': '171758',
            'ChipSet': '171759',
            'CPU': '171769',
            'Docks/Stands': '171770',
            'EC(Embedded Controller)': '171791',
            'Factory': '171792',
            'Human Interface Device (HID)': '171798',
            'I/O Ports': '171818',
            'Intel MSR': '171848',
            'LCD Panel (touch/non-touch)': '171860',
            'Lights/LEDs': '171874',
            'Manageability and Security': '171878',
            'Mechanicals': '171892',
            'Memory': '171912',
            'Miscellaneous': '171922',
            'Monitors': '171923',
            'Network': '171924',
            'Operating System': '172083',
            'Peripherals/Externals': '172096',
            'Platform': '172152',
            'Power Adapters': '172153',
            'Power Supply/PSU': '172154',
            'Process (Internal/External)': '172155',
            'System Board': '172261',
            'Thermals/Acoustics': '172262',
            'Video': '172263',
           }
            # Problem Category 列表
            problem_mapping = {
             'Battery': '120759',
             'BIOS': '120181',
             'Chassis': '120988',
             'EE': '128414',
             'Firmware': '120185',
             'Input': '120157',
             'Keypart': '128415',
             'Mechanicals': '120653',
             'NAS': '120173',
             'Operating System': '120228',
             'Network': '121195',
             'Power': '120152',
             'Security Encryption': '121061',
             'SW': '128416',
             'Systems Management': '120341',
             'Video': '120926',
             'Other': '128417',
             'BIOS|CPU:General': '132319',
             'BIOS|Setup:Main': '132321',
             'Operating System|Driver': '132325',
             'BIOS|Setup:Security': '132330',
             'Power|Power:Speed/Function': '132331'}
            # Component / Subcomponent 列表 #20240306 Update
            component_mapping = {
            'None':'-1',
            'Power Adapters':'120042',
            'SW Apps & Technologies':'120050',
            'Audio':'120072',
            'Battery':'119968',
            'BIOS':'120032',
            'CPU':'119953',
            'Monitors':'119984',
            'Docks/Stands':'119903',
            'I/O Ports':'165897',
            'LCD Panel':'127316',
            'Platform':'119901',
            'Power Supply/PSU':'120062',
            'System Board':'120082',
            'Thermals/Acoustics':'120001',
            'USB Power Delivery (PD) Controller':'120044',
            'Video':'119979',
            'Operating System':'120088',
            'Human Interface Device (HID)':'119950',
            'Peripherals/Externals':'159877',
            'Camera':'165896',
            'ChipSet':'119937',
            'Factory':'165410',
            'Lights/LEDs':'165625',
            'Mechanicals':'119947',
            'Regulatory Compliance':'165828',
            'Sensors':'165850',
            'Storage':'120022'
            }
            
            for reissue in issues:
                self.trigger_pims.emit(f"Currently processing Redmine id: {reissue.id}")
                # 重新整理資訊
                lc = ''
                dt_pqm = ''
                dd_pqm = ''
                sc_pqm = ''
                trc_pqm = ''
                pqm_sub = ''
                pc_pqm = ''
                issue_status_pqm = ''
                transferred_assignee = ''
                redmine_document = ''
                pims = ''
                # 執行所有id對比pims
                connection_server = 0
                while True:
                    try:
                        # 登入redmine
                        redmine = Redmine(str(redmine_web), username=str(redmine_account), password=str(redmine_password))
                        re_issue_id = reissue.id
                        re_issue = redmine.issue.get(re_issue_id)
                        break
                    except:
                        if connection_server == 0:
                            self.trigger_wait.emit("Connection abnormal, waiting for the network to reconnect.")
                            connection_server+=1
                # 狀態
                re_status = re_issue.status.name
                '''
                try:
                    re_category = re_issue.category.name
                except:
                    re_category = '-'
                
                # Assignee
                try:
                    re_assigned_to = re_issue.assigned_to.name
                    re_assigned_to = re_assigned_to.replace(' ', '_') + '@compal.com'
                    re_assigned_to_1 = re_issue.assigned_to.name
                    #範例抓取出來:minse_yang
                    re_user_info = jira_client.search_users(re_assigned_to)[0].name  
                except:
                    re_assigned_to = ''
                    re_assigned_to_1 = '-' #用來顯示
                    re_user_info = ''
                '''    
                # 遍歷資料訊息
                for field in re_issue.custom_fields:
                    if field.name == "Feature":
                        re_category = field.value
                    if field.name == "Leader Description":
                        lc = field.value
                    if field.name == "Disposition Type":
                        dt_pqm = field.value
                    if field.name == "A31_Task Owner":
                        user_id = field.value
                        user_obj = redmine.user.get(user_id)
                        user_name = user_obj.login  # 例如 'kellyph_chen'
                        re_assigned_to = user_name + '@compal.com'
                        #print("re_assigned_to:", re_assigned_to)  # 印出 re_assigned_to
                        re_user_info = jira_client.search_users(re_assigned_to)[0].name    
                    '''  
                    #not use 
                    if field.name == "Solution Category for PQM":
                        sc_pqm = field.value
                    ''' 
                    if field.name == "Disposition Details":
                        dd_pqm = field.value                                      
                    if field.name == "Tech Root Cause":
                        trc_pqm = field.value
                    '''
                    #not use
                    if field.name == "Problem Category for PQM":
                        pc_pqm =  field.value
                    '''    
                    if field.name == "Component":
                        pqm_sub = field.value
                    if field.name == "PSR Number":
                        pims = field.value
                        pims = ''.join(pims.split())
                    if field.name == "Assignee":
                        transferred_assignee = field.value
                        try:
                            #範例抓取出來:minse_yang
                            transferred_user_info = jira_client.search_users(transferred_assignee)[0].name
                        except:
                            transferred_user_info = ''
                    if field.name == "Attach file to Jira":
                        redmine_document = field.value
                    if field.name == "Status":
                        issue_status_pqm = field.value
                    if field.name == "Priority":
                        issue_severity_pqm = field.value
                self.trigger_pims.emit("Redmine id : " + str(re_issue_id) + " / Status : " + re_status )
                self.trigger_pims.emit("Redmine Assignee : " + str(re_assigned_to) + " / Feature : " + str(re_category) )
                #print("Redmine id : " + str(re_issue_id) + " / Status : " + re_status)
                try:
                    # 抓取PIMS資訊
                    if pims != '':
                        connection_count = 0
                        while True:
                            try: 
                                response = requests.get(jira_web, auth=(str(jira_account),str(jira_password)), timeout=10, verify=jira_verify)
                                #self.trigger_wait.emit("Retrieving information...")
                                if response.status_code // 100 == 2:
                                    #pims = 'CEP-14875' 
                                    #pims = 'PIMS-284193'
                                    pims_issue = jira_client.issue(pims)
                                    #print(dir(pims_issue.fields))
                                    comment = jira_client.comments(pims)
                                    pims_status = pims_issue.fields.status 
                                    try:
                                        pims_dt_pqm = pims_issue.fields.customfield_18716.value
                                    except:
                                        pims_dt_pqm = ""
                                    try:
                                        pims_dd_pqm = pims_issue.fields.customfield_18715
                                    except:
                                        pims_dd_pqm = ""
                                    try:
                                        pims_trc_pqm = pims_issue.fields.customfield_26700
                                    except:
                                        pims_trc_pqm = ""
                                    executive_summary = pims_issue.fields.customfield_28473
                                    try:
                                        pims_pqm_sub = pims_issue.fields.customfield_28436[0].value
                                    except:
                                        pims_pqm_sub = ""
                                    try:
                                        discretionary_field2 = pims_issue.fields.customfield_28468.value
                                    except:
                                        discretionary_field2 = ""
                                    try:
                                        pims_issue_severity = pims_issue.fields.customfield_28162
                                    except:
                                        pims_issue_severity = ""
                                    connection_ok = 1
                                    break  # 連接成功，退出迴圈
                            except:
                                if connection_count < 3:
                                    self.trigger_wait.emit("Error message : Redmine id : " + str(re_issue_id) + "/ Retry connecting: Attempt " + str(connection_count + 1))
                                    connection_count += 1
                                elif connection_count == 3:
                                    self.trigger_wait.emit("Connection abnormal, waiting for the network to reconnect.")
                                    connection_count += 1
                        
                        if connection_ok == 1:
                            update_error = 0
                            while True:
                                try:
                                    # comment
                                    if re_status not in ['New', 'Closed', 'Transferred-Closed', 'ATS (P1/P2/P3)-Closed', 'ATS (P4/P5) / WAD', 'Transferred']:
                                        comsun = 0
                                        try:
                                            if lc != "" and '[20xxxxxx EC xxx]' not in lc:
                                                # 使用正規表達式找到 [EC status start] 到 [EC status end] 之間的字串並刪除
                                                pattern = re.compile(r'\[EC status start\].*?\[EC status end\]', re.DOTALL)
                                                # 用來加入comment的字串
                                                lc_new = pattern.sub('', lc)
                                                # 如果沒有評論
                                                if not comment:
                                                    jira_client.add_comment(pims, str(lc_new))
                                                    self.trigger.emit('Update the comment as completed.')
                                                # 如果有評論檢查最後一則評論是否一樣
                                                elif len(comment) > 0:
                                                    for i in range(0,len(comment)):
                                                        if comment[i].body == lc_new:
                                                            comsun += 1
                                                    if comsun == 0:
                                                        jira_client.add_comment(pims, str(lc_new))
                                                        self.trigger.emit('Update the comment as completed.')
                                                    else:
                                                        self.trigger.emit("No update needed as the comment remains the same.")
                                        except:
                                            pass
                                        try:
                                            # 使用 re.search() 提取 [EC status start] 和 [EC status end] 之間的部分
                                            redmine_match = re.search(r'\[EC status start\](.*?)\[EC status end\]', lc, re.DOTALL)
                                            if redmine_match:
                                                #用判斷裡面是不是空值
                                                r_extracted_part_judge = redmine_match.group(1).strip()
                                                #要加入的資料
                                                r_extracted_part = redmine_match.group(0).strip()
                                            else:
                                                r_extracted_part_judge = ''
                                                r_extracted_part = ''
                                            # 如果jira executive summary不是空值
                                            if executive_summary != None:
                                                # 使用正規表達式找到 [EC status start] 到 [EC status end] 之間的字串並刪除
                                                pattern = re.compile(r'\[EC status start\].*?\[EC status end\]', re.DOTALL)
                                                executive_summary_new = pattern.sub('', executive_summary)
                                                #將jira 裡面的EC資訊抓取出來
                                                jira_match = re.search(r'\[EC status start\](.*?)\[EC status end\]', executive_summary, re.DOTALL)
                                                if jira_match:
                                                    j_extracted_part = jira_match.group(0).strip()
                                                else:
                                                    j_extracted_part = ''
                                                # 如果Redmine上不是空值
                                                if r_extracted_part_judge != "":
                                                    # 判斷j_extracted_part是否和r_extracted_part
                                                    if j_extracted_part != r_extracted_part:
                                                        new_es = r_extracted_part + "\n" + executive_summary_new
                                                        pims_issue.update(customfield_28473 = str(new_es))
                                                        self.trigger.emit('Update the executive summary as completed.')
                                            elif r_extracted_part_judge != "":
                                                pims_issue.update(customfield_28473 = str(r_extracted_part))
                                                self.trigger.emit('Update the executive summary as completed.')
                                        except:
                                            pass
                                    # New方法
                                    if re_status == 'New':
                                        if pims_status.name == 'Submitted' or pims_status.name == 'Analyze':
                                            self.trigger.emit("Do nothing!")
                                        else:
                                            self.trigger_fail.emit("Error message : (Redmine id : " + str(re_issue_id) + " / Status : " + re_status + ");(" + pims + " / Status : " + pims_status.name + ")")
                      
                                    # Open方法
                                    elif re_status == 'Open':
                                        #new_path = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
                                        new_path = os.getcwd()
                                        #檢查CSV有沒有id資訊,帶出來
                                        df = pd.read_csv(new_path + '\\marker.csv')
                                        try:
                                            #如果csv找尋到id就把modify_code帶出,代表有跑過一次
                                            index_of_id_name = int(df.index[df['Redmine_id'] == re_issue_id].values[0])
                                            #刪除狀態
                                            df = df.drop(index_of_id_name)
                                            print(f"Remove Redmine id {re_issue_id} from marker.csv")
                                            df.to_csv(new_path + '\\marker.csv', index=False)
                                        except:
                                            pass       
                                        # 當前PIMS不等於Redmine人員時修改為Redmine Assignee的人
                                        if pims_status.name == 'Analyze':
                                            if re_user_info != "":
                                                try:
                                                    # 抓取當前PIMS Assignee 人員 (minse_yang)
                                                    assignee_name_now = pims_issue.fields.assignee.name
                                                    if assignee_name_now != re_user_info:
                                                        if assignee_name_now != account_user_info:
                                                            #先修改成自己在修改Redmine Assignee的人
                                                            jira_client.assign_issue(pims, account_user_info)
                                                        jira_client.assign_issue(pims, re_user_info)
                                                        self.trigger.emit('Change the Assignee to "' + str(re_user_info) + '" successfully.')
                                                    else:
                                                        self.trigger.emit('The Assignee is the same person, do nothing.')
                                                    if pims_issue.fields.customfield_28471.name != transferred_user_info:
                                                        jira_client.transition_issue(pims, 371, fields={'customfield_28471': {'name': re_user_info}})
                                                except:
                                                    self.trigger_fail.emit("Error message : PIMS cannot locate the original assignee.")
                                            else:
                                                self.trigger_fail.emit('Error message : Redmine id : '+ str(re_issue_id)  +' has not been assigned to anyone.')           
                                        elif pims_status.name != 'Analyze':
                                            self.trigger_fail.emit("Error message : (Redmine id : " + str(re_issue_id) + " / Status : " + re_status + ");(" + pims + " / Status : " + pims_status.name + ")")
                                    
                                    # Resolved方法
                                    elif re_status == 'Resolved':
                                        id_value = option_mapping.get(str(dt_pqm))
                                        if pims_status.name == 'Analyze':
                                            try:
                                                if pims_dt_pqm != dt_pqm and id_value != None:
                                                    pims_issue.update(customfield_18716 = {'id': id_value})
                                                    self.trigger.emit('Change the Disposition type to "' + str(dt_pqm) + '" successfully.')
                                                elif id_value == None:
                                                    self.trigger_fail.emit("Error message : Check if the Disposition Type is inconsistent. Redmine : " + str(dt_pqm) + ", Jira : " + str(pims_dt_pqm))
                                                if pims_dd_pqm != dd_pqm:
                                                    pims_issue.update(customfield_18715 = str(dd_pqm))
                                                    self.trigger.emit('Change the Disposition Details to "' + str(dd_pqm) + '" successfully.')
                                                if pims_trc_pqm != trc_pqm:
                                                    pims_issue.update(customfield_26700 = str(trc_pqm))
                                                    self.trigger.emit('Change the Technical root cause to "' + str(trc_pqm) + '" successfully.')
                                                self.trigger.emit('Check completed.')
                                            except:
                                                self.trigger_fail.emit("Error message : During the status transition process, check for any issues with mandatory fields.")
                                        elif pims_status.name != 'Analyze':
                                            self.trigger_fail.emit("Error message : (Redmine id : " + str(re_issue_id) + " / Status : " + re_status + ");(" + pims + " / Status : " + pims_status.name + ")")
                                    
                                    # Review方法
                                    elif re_status == 'Review' or  re_status == 'ATS (P1/P2/P3)':
                                        if re_status == 'ATS (P1/P2/P3)':
                                            #new_path = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
                                            new_path = os.getcwd()
                                            #檢查CSV有沒有id資訊,帶出來
                                            df = pd.read_csv(new_path + '\\marker.csv')
                                            try:
                                                #如果csv找尋到id就把modify_code帶出,代表有跑過一次
                                                index_of_id_name = int(df.index[df['Redmine_id'] == re_issue_id].values[0])
                                                #刪除狀態
                                                df = df.drop(index_of_id_name)
                                                df.to_csv(new_path + '\\marker.csv', index=False)
                                            except:
                                                pass
                                           
                                        # 檔案上傳
                                        if (pims_status.name == 'Review' or pims_status.name == 'Analyze') and (re_status == 'ATS (P1/P2/P3)' and redmine_document != ""):
                                            redmine_attachment = redmine.attachment.get(redmine_document)
                                            attachment_content = redmine_attachment.download().content
                                            # 包裝成模擬文件
                                            with io.BytesIO(attachment_content) as attachment_file:
                                                # 上傳資料
                                                jira_attachments = jira_client.issue(pims).fields.attachment
                                                jira_attachment_list = [str(attachment.filename) for attachment in jira_attachments]
                                        
                                                if str(redmine_attachment.filename) not in jira_attachment_list:
                                                    jira_client.add_attachment(issue=pims, attachment=attachment_file, filename=str(redmine_attachment.filename))
                                                    self.trigger.emit('The file "' + str(redmine_attachment.filename) + '" has been successfully uploaded.')
                                                else:
                                                    self.trigger.emit('This file "' + str(redmine_attachment.filename) + '" already exists and will not be uploaded.')
                                        # 如果PIMS為Analyze : 修改Ａｓｓｉｇｎｅｅ、　Disposition typｅ、　Disposition Details、 Technical root cause
                                        if pims_status.name == 'Analyze':
                                            try:
                                                # 抓取當前PIMS Assignee 人員 (minse_yang)
                                                assignee_name_now = pims_issue.fields.assignee.name
                                                self.trigger.emit('Current Jira assignee is"' + str(assignee_name_now) +  '" .')
                                                try:
                                                    #判斷ATS是否正確
                                                    if re_status == 'ATS (P1/P2/P3)' and (discretionary_field2 != 'ATS P1' and discretionary_field2 != 'ATS P2' and discretionary_field2 != 'ATS P3'):
                                                        redmine.issue.update(re_issue_id, notes='DF2 is not expected!')
                                                        self.trigger_fail.emit("Error message : DF2 is not expected!")
                                                    id_value = option_mapping.get(str(dt_pqm))
                                                    #sc_value = solution_mapping.get(str(sc_pqm)) #not use
                                                    #if isinstance(sc_value, str):
                                                        #sc_value = [sc_value] 
                                                    # 先修改成Assign自己
                                                    jira_client.assign_issue(pims, account_user_info)
                                                    self.trigger.emit('Change Jira assignee to"' + str(account_user_info) +  '" .')
                                                    if pims_issue.fields.issuetype.name == 'Change Request':
                                                        #pc_value = problem_mapping.get(str(pc_pqm)) #not use
                                                        jira_client.transition_issue(pims, 61,fields={'customfield_18716': {"id":id_value},
                                                                                                      'customfield_18715': str(dd_pqm),
                                                                                                      'customfield_26700': str(trc_pqm),
                                                                                                      #'customfield_28480': {"id":pc_value},
                                                                                                      #'customfield_45100': [{"id": value} for value in sc_value]
                                                                                                      })
                                                        # 先修改成Assign自己
                                                        jira_client.assign_issue(pims, account_user_info)

                                                        # 改回來本人
                                                        jira_client.assign_issue(pims, assignee_name_now)
                                                        self.trigger.emit('Change Jira assignee back to"' + str(assignee_name_now) +  '" .')
                                                        self.trigger.emit('Change the status of PIMS to "Review" successfully.')
                                                        self.trigger.emit('Change the Disposition type to "' + str(dt_pqm) + '" successfully.')
                                                        self.trigger.emit('Change the Disposition Details to "' + str(dd_pqm) + '" successfully.')
                                                        self.trigger.emit('Change the Technical root cause to "' + str(trc_pqm) + '" successfully.')
                                                        #self.trigger.emit('Change the Problem Category to "' + str(pc_pqm) + '" successfully.')  #not use
                                                    else:
                                                        jira_client.transition_issue(pims, 61,fields={'customfield_18716': {"id":id_value},
                                                                                                      'customfield_18715': str(dd_pqm),
                                                                                                      'customfield_26700': str(trc_pqm),
                                                                                                      #'customfield_45100': [{"id": value} for value in sc_value] #not use
                                                                                                      })
                                                        # 先修改成Assign自己
                                                        jira_client.assign_issue(pims, account_user_info)
                                                        # 改回來本人
                                                        jira_client.assign_issue(pims, assignee_name_now)
                                                        self.trigger.emit('Change Jira assignee back to"' + str(assignee_name_now) +  '" .')
                                                        self.trigger.emit('Change the status of PIMS to "Review" successfully.')
                                                        self.trigger.emit('Change the Disposition type to "' + str(dt_pqm) + '" successfully.')
                                                        self.trigger.emit('Change the Disposition Details to "' + str(dd_pqm) + '" successfully.')
                                                        self.trigger.emit('Change the Technical root cause to "' + str(trc_pqm) + '" successfully.')     
                                                except:
                                                    jira_client.assign_issue(pims, assignee_name_now)
                                                    self.trigger_fail.emit("Error message : During the status transition process, check for any issues with mandatory fields.")
                                            except:
                                                self.trigger_fail.emit("Error message : PIMS cannot locate the original assignee.")
                                        # 修改Disposition typｅ、　Disposition Details、 Technical root cause
                                        else:
                                            try:
                                                #判斷ATS是否正確
                                                if re_status == 'ATS (P1/P2/P3)' and (discretionary_field2 != 'ATS P1' and discretionary_field2 != 'ATS P2' and discretionary_field2 != 'ATS P3'):
                                                    redmine.issue.update(re_issue_id, notes='DF2 is not expected!')
                                                    self.trigger_fail.emit("Error message : DF2 is not expected!")
                                                id_value = option_mapping.get(str(dt_pqm))
                                                if pims_dt_pqm != dt_pqm and id_value != None:
                                                    pims_issue.update(customfield_18716 = {'id': id_value})
                                                    self.trigger.emit('Change the Disposition type to "' + str(dt_pqm) + '" successfully.')
                                                elif id_value == None:
                                                    self.trigger_fail.emit("Error message : Check if the Disposition Type is inconsistent. Redmine : " + str(dt_pqm) + ", Jira : " + str(pims_dt_pqm))
                                                if pims_dd_pqm != dd_pqm:
                                                    pims_issue.update(customfield_18715 = str(dd_pqm))
                                                    self.trigger.emit('Change the Disposition Details to "' + str(dd_pqm) + '" successfully.')
                                                if pims_trc_pqm != trc_pqm:
                                                    pims_issue.update(customfield_26700 = str(trc_pqm))
                                                    self.trigger.emit('Change the Technical root cause to "' + str(trc_pqm) + '" successfully.')
                                                self.trigger.emit('Check completed.')
                                            except:
                                                self.trigger_fail.emit("Error message : During the status transition process, check for any issues with mandatory fields.")
                                            # 如果PIMS不為Review, Show error 
                                            if pims_status.name != 'Review':
                                                self.trigger_fail.emit("Error message : (Redmine id : " + str(re_issue_id) + " / Status : " + re_status + ");(" + pims + " / Status : " + pims_status.name + ")")
                                    
                                    # Verify方法
                                    elif re_status == 'Verify' or re_status == 'ATS (P1/P2/P3)-Verify':
                                        #new_path = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
                                        new_path = os.getcwd()
                                        #檢查CSV有沒有id資訊,帶出來
                                        df = pd.read_csv(new_path + '\\marker.csv')
                                        try:
                                            #如果csv找尋到id就把modify_code帶出,代表有跑過一次
                                            index_of_id_name = int(df.index[df['Redmine_id'] == re_issue_id].values[0])
                                            modify_code = df['modify_code'][index_of_id_name]
                                            if modify_code !=1:
                                                try:
                                                    #刪除狀態
                                                    df = df.drop(index_of_id_name)
                                                    df.to_csv(new_path + '\\marker.csv', index=False)
                                                    #再添加一次進CSV
                                                    df.loc[df.index.max()+1] = [re_issue_id, 1]
                                                    df.to_csv(new_path + '\\marker.csv', index=False)
                                                    modify_code = 0
                                                except:
                                                    pass                                                  
                                        except:
                                            #如果csv沒有id資訊就添加進去
                                            df.loc[df.index.max()+1] = [re_issue_id, 1]
                                            df.to_csv(new_path + '\\marker.csv', index=False)
                                            modify_code = 0
                                        
                                        if modify_code == 1:
                                            #已更新過狀態為Verify而且PIMS狀態也還在Verify
                                            if pims_status.name == 'Verify':
                                                try:
                                                    id_value = option_mapping.get(str(dt_pqm))
                                                    if pims_dt_pqm != dt_pqm and id_value != None:
                                                        pims_issue.update(customfield_18716 = {'id': id_value})
                                                        self.trigger.emit('Change the Disposition type to "' + str(dt_pqm) + '" successfully.')
                                                    elif id_value == None:
                                                        self.trigger_fail.emit("Error message : Check if the Disposition Type is inconsistent. Redmine : " + str(dt_pqm) + ", Jira : " + str(pims_dt_pqm))
                                                    if pims_dd_pqm != dd_pqm:
                                                        pims_issue.update(customfield_18715 = str(dd_pqm))
                                                        self.trigger.emit('Change the Disposition Details to "' + str(dd_pqm) + '" successfully.')
                                                    if pims_trc_pqm != trc_pqm:
                                                        pims_issue.update(customfield_26700 = str(trc_pqm))
                                                        self.trigger.emit('Change the Technical root cause to "' + str(trc_pqm) + '" successfully.')
                                                    self.trigger.emit('Check completed.')
                                                    if re_status == 'ATS (P1/P2/P3)-Verify':
                                                        # 檔案上傳
                                                        if redmine_document != "":
                                                            redmine_attachment = redmine.attachment.get(redmine_document)
                                                            attachment_content = redmine_attachment.download().content
                                                            # 包裝成模擬文件
                                                            with io.BytesIO(attachment_content) as attachment_file:
                                                                # 上傳資料
                                                                jira_attachments = jira_client.issue(pims).fields.attachment
                                                                jira_attachment_list = [str(attachment.filename) for attachment in jira_attachments]
                                                        
                                                                if str(redmine_attachment.filename) not in jira_attachment_list:
                                                                    jira_client.add_attachment(issue=pims, attachment=attachment_file, filename=str(redmine_attachment.filename))
                                                                    self.trigger.emit('The file "' + str(redmine_attachment.filename) + '" has been successfully uploaded.')
                                                                else:
                                                                    self.trigger.emit('This file "' + str(redmine_attachment.filename) + '" already exists and will not be uploaded.')
                                                except:
                                                    self.trigger_fail.emit("Error message : During the status transition process, check for any issues with mandatory fields.")
                                            
                                            elif pims_status.name == 'Closed' or pims_status.name == 'Draft' or pims_status.name == 'cancel':
                                                if re_status == 'ATS (P1/P2/P3)-Verify':
                                                    #更改Redmine狀態為ATS (P1/P2/P3)-Closed (Status ID: 19, Status Name: ATS P1/P2/P3 - Closed)
                                                    redmine.issue.update(re_issue_id, status_id = 19, custom_fields=[{'id': 157, 'value': 'Closed'}]) #cf_157=Jira Status
                                                    self.trigger.emit('Due to the Jira status being "' + str( pims_status.name) + '", change the Redmine status to "ATS (P1/P2/P3)-Closed" successfully.')
                                                else:
                                                    #更改Redmine狀態為Closed (Status ID: 5, Status Name: Closed)
                                                    redmine.issue.update(re_issue_id, status_id = 5, custom_fields=[{'id': 157, 'value': 'Closed'}]) #cf_157=Jira Status
                                                    self.trigger.emit('Due to the Jira status being "' + str( pims_status.name) + '", change the Redmine status to "Closed" successfully.')
                                                #刪除狀態
                                                df = df.drop(index_of_id_name)
                                                #檢查CSV有沒有id資訊,帶出來
                                                df.to_csv(new_path + '\\marker.csv', index=False)
                                            
                                            else:
                                                #已更新過一次Update, 但狀態不是Verify需要檢查 
                                                if pims_status.name == 'Analyze':
                                                    if re_status == 'ATS (P1/P2/P3)-Verify':
                                                        redmine.issue.update(re_issue_id, status_id = 16)   #Status ID: 16, Status Name: ATS P1/P2/P3
                                                        redmine.issue.update(re_issue_id, notes='Due to the PIMS status being "Analyze", the status of Redmine ID '+ str(re_issue_id)  +' has been corrected back to "ATS (P1/P2/P3)". ')
                                                        self.trigger.emit('Due to the PIMS status being "Analyze", the status of Redmine ID '+ str(re_issue_id)  +' has been corrected back to "ATS (P1/P2/P3)". ')
                                                    else:
                                                        redmine.issue.update(re_issue_id, status_id = 2)    #Status ID: 2, Status Name: Open
                                                        redmine.issue.update(re_issue_id, notes='Due to the PIMS status being "Analyze", the status of Redmine ID '+ str(re_issue_id)  +' has been corrected back to "Open". ')
                                                        self.trigger.emit('Due to the PIMS status being "Analyze", the status of Redmine ID '+ str(re_issue_id)  +' has been corrected back to "Assigned". ')
                                                else:
                                                    self.trigger_fail.emit("Error message : The Redmine id : " + str(re_issue_id) + " has been updated once, and the " + pims + " status is reverted to " + pims_status.name + ".")
                                                #刪除狀態
                                                df = df.drop(index_of_id_name)
                                                #檢查CSV有沒有id資訊,帶出來
                                                df.to_csv(new_path + '\\marker.csv', index=False)
                                        else:
                                            if pims_status.name == 'Analyze':
                                                id_value = option_mapping.get(str(dt_pqm))
                                                try:
                                                    # 抓取當前PIMS Assignee 人員 (minse_yang)
                                                    assignee_name_now = pims_issue.fields.assignee.name
                                                    # 先修改成Assign自己
                                                    jira_client.assign_issue(pims, account_user_info)
                                                    try:

                                                        if pims_issue.fields.issuetype.name == 'Change Request':
                                                            #pc_value = problem_mapping.get(str(pc_pqm)) not use
                                                            jira_client.transition_issue(pims, 61,fields={'customfield_18716': {"id":id_value},
                                                                                                          'customfield_18715': str(dd_pqm),
                                                                                                          'customfield_26700': str(trc_pqm),
                                                                                                          #'customfield_28480': {"id":pc_value}, not use
                                                                                                          })
                                                            self.trigger.emit('Change the status of PIMS to "Review" successfully.')
                                                            self.trigger.emit('Change the Disposition type to "' + str(dt_pqm) + '" successfully.')
                                                            self.trigger.emit('Change the Disposition Details to "' + str(dd_pqm) + '" successfully.')
                                                            self.trigger.emit('Change the Technical root cause to "' + str(trc_pqm) + '" successfully.')
                                                            #self.trigger.emit('Change the Problem Category to "' + str(pc_pqm) + '" successfully.')  
                                                            # 轉換完後會讓assign變成Reporter的人,所以要再轉換一次
                                                            jira_client.assign_issue(pims, account_user_info)
                                                            # 狀態修改為Verify
                                                            jira_client.transition_issue(pims, 71)
                                                        else:
                                                            jira_client.transition_issue(pims, 61,fields={'customfield_18716': {"id":id_value},
                                                                                                          'customfield_18715': str(dd_pqm),
                                                                                                          'customfield_26700': str(trc_pqm),
                                                                                                          })
                                                            self.trigger.emit('Change the status of PIMS to "Review" successfully.')
                                                            self.trigger.emit('Change the Disposition type to "' + str(dt_pqm) + '" successfully.')
                                                            self.trigger.emit('Change the Disposition Details to "' + str(dd_pqm) + '" successfully.')
                                                            self.trigger.emit('Change the Technical root cause to "' + str(trc_pqm) + '" successfully.')
                                                            # 轉換完後會讓assign變成Reporter的人,所以要再轉換一次
                                                            jira_client.assign_issue(pims, account_user_info)
                                                            # 狀態修改為Verify
                                                            jira_client.transition_issue(pims, 71)
                                                        self.trigger.emit('Change the status of PIMS to "Verify" successfully.')
                                                    except:
                                                        jira_client.assign_issue(pims, assignee_name_now)
                                                        self.trigger_fail.emit("Error message : During the status transition process, check for any issues with mandatory fields.")
                                                except:
                                                    self.trigger_fail.emit("Error message : PIMS cannot locate the original assignee.")

                                            elif pims_status.name == 'Review':
                                                try:
                                                    # 抓取當前PIMS Assignee 人員 (minse_yang)
                                                    assignee_name_now = pims_issue.fields.assignee.name
                                                    try:
                                                        id_value = option_mapping.get(str(dt_pqm))
                                                        if pims_dt_pqm != dt_pqm and id_value != None:
                                                            pims_issue.update(customfield_18716 = {'id': id_value})
                                                            self.trigger.emit('Change the Disposition type to "' + str(dt_pqm) + '" successfully.')
                                                        elif id_value == None:
                                                            self.trigger_fail.emit("Error message : Check if the Disposition Type is inconsistent. Redmine : " + str(dt_pqm) + ", Jira : " + str(pims_dt_pqm))
                                                        if pims_dd_pqm != dd_pqm:
                                                            pims_issue.update(customfield_18715 = str(dd_pqm))
                                                            self.trigger.emit('Change the Disposition Details to "' + str(dd_pqm) + '" successfully.')
                                                        if pims_trc_pqm != trc_pqm:
                                                            pims_issue.update(customfield_26700 = str(trc_pqm))
                                                            self.trigger.emit('Change the Technical root cause to "' + str(trc_pqm) + '" successfully.')
                                                        # 先修改成Assign自己
                                                        jira_client.assign_issue(pims, account_user_info)
                                                        # 狀態修改為Verify
                                                        jira_client.transition_issue(pims, 71)
                                                        self.trigger.emit('Change the status of PIMS to "Verify" successfully.')
                                                    except:
                                                        jira_client.assign_issue(pims, assignee_name_now)
                                                        self.trigger_fail.emit("Error message : During the status transition process, check for any issues with mandatory fields.")
                                                except:
                                                    self.trigger_fail.emit("Error message : PIMS cannot locate the original assignee.")
                                            elif pims_status.name == 'Closed' or pims_status.name == 'Draft' or pims_status.name == 'cancel':
                                                #更改Redmine狀態為Closed
                                                redmine.issue.update(re_issue_id, status_id = 5, custom_fields=[{'id': 157, 'value': 'Closed'}]) #cf_157=Jira Status
                                                self.trigger.emit('Due to the Jira status being "' + str( pims_status.name) + '", change the Redmine status to "Closed" successfully.')
                                                #刪除狀態
                                                df = pd.read_csv(new_path + '\\marker.csv')
                                                index_of_id_name = int(df.index[df['Redmine_id'] == re_issue_id].values[0])
                                                df = df.drop(index_of_id_name)
                                                df.to_csv(new_path + '\\marker.csv', index=False)                                           
                                            else:
                                                try:
                                                    id_value = option_mapping.get(str(dt_pqm))
                                                    if pims_dt_pqm != dt_pqm and id_value != None:
                                                        pims_issue.update(customfield_18716 = {'id': id_value})
                                                        self.trigger.emit('Change the Disposition type to "' + str(dt_pqm) + '" successfully.')
                                                    elif id_value == None:
                                                        self.trigger_fail.emit("Error message : Check if the Disposition Type is inconsistent. Redmine : " + str(dt_pqm) + ", Jira : " + str(pims_dt_pqm))
                                                    if pims_dd_pqm != dd_pqm:
                                                        pims_issue.update(customfield_18715 = str(dd_pqm))
                                                        self.trigger.emit('Change the Disposition Details to "' + str(dd_pqm) + '" successfully.')
                                                    if pims_trc_pqm != trc_pqm:
                                                        pims_issue.update(customfield_26700 = str(trc_pqm))
                                                        self.trigger.emit('Change the Technical root cause to "' + str(trc_pqm) + '" successfully.')
                                                    self.trigger.emit('Check completed.')
                                                except:
                                                    self.trigger_fail.emit("Error message : During the status transition process, check for any issues with mandatory fields.")
                                                if pims_status.name != 'Verify':
                                                    self.trigger_fail.emit("Error message : (Redmine id : " + str(re_issue_id) + " / Status : " + re_status + ");(" + pims + " / Status : " + pims_status.name + ")")
                                            #上傳檔案
                                            if re_status == 'ATS (P1/P2/P3)-Verify':
                                              # 檔案上傳
                                              if (pims_status.name == 'Review' or pims_status.name == 'Analyze' or pims_status.name == 'Verify') and redmine_document != "":
                                                  redmine_attachment = redmine.attachment.get(redmine_document)
                                                  attachment_content = redmine_attachment.download().content
                                                  # 包裝成模擬文件
                                                  with io.BytesIO(attachment_content) as attachment_file:
                                                      # 上傳資料
                                                      jira_attachments = jira_client.issue(pims).fields.attachment
                                                      jira_attachment_list = [str(attachment.filename) for attachment in jira_attachments]
                                              
                                                      if str(redmine_attachment.filename) not in jira_attachment_list:
                                                          jira_client.add_attachment(issue=pims, attachment=attachment_file, filename=str(redmine_attachment.filename))
                                                          self.trigger.emit('The file "' + str(redmine_attachment.filename) + '" has been successfully uploaded.')
                                                      else:
                                                          self.trigger.emit('This file "' + str(redmine_attachment.filename) + '" already exists and will not be uploaded.')
                                    
                                    # Closed方法
                                    elif re_status == 'Closed':
                                        #檢查CSV有沒有id資訊,帶出來
                                        #new_path = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
                                        new_path = os.getcwd()
                                        df = pd.read_csv(new_path + '\\marker.csv')
                                        try:
                                            #如果csv找尋到id就把modify_code帶出,代表有跑過一次
                                            index_of_id_name = int(df.index[df['Redmine_id'] == re_issue_id].values[0])
                                            #刪除狀態
                                            df = df.drop(index_of_id_name)
                                            df.to_csv(new_path + '\\marker.csv', index=False)
                                        except:
                                            pass
                                        if pims_status.name == 'Closed' or pims_status.name == 'Draft' or pims_status.name == 'cancel':
                                            self.trigger.emit('This issue has been closed.')
                                        else:
                                            self.trigger_fail.emit("Error message : (Redmine id : " + str(re_issue_id) + " / Status : " + re_status + ");(" + pims + " / Status : " + pims_status.name + ")")
                                    
                                    # Transferred
                                    elif re_status == 'Tranferred':
                                        #檢查CSV有沒有id資訊,帶出來
                                        #new_path = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
                                        new_path = os.getcwd()
                                        df = pd.read_csv(new_path + '\\marker.csv')
                                        try:
                                            #如果csv找尋到id就把modify_code帶出,代表有跑過一次
                                            index_of_id_name = int(df.index[df['Redmine_id'] == re_issue_id].values[0])
                                            modify_code = df['modify_code'][index_of_id_name]
                                            if modify_code !=2:
                                                try:
                                                    #刪除狀態
                                                    df = df.drop(index_of_id_name)
                                                    df.to_csv(new_path + '\\marker.csv', index=False)
                                                    #再添加一次進CSV
                                                    df.loc[df.index.max()+1] = [re_issue_id, 2]
                                                    df.to_csv(new_path + '\\marker.csv', index=False)
                                                    modify_code = 0
                                                except:
                                                    pass 
                                        except:
                                            #如果csv沒有id資訊就添加進去
                                            df.loc[df.index.max()+1] = [re_issue_id, 2]
                                            df.to_csv(new_path + '\\marker.csv', index=False)
                                            modify_code = 0
                                        if modify_code == 2:
                                            #已檢查過狀態為Transferred
                                            if pims_status.name == 'Analyze' or pims_status.name == 'Wait' or pims_status.name == 'Review' or pims_status.name == 'Verify':
                                                try:
                                                    if pqm_sub not in pims_pqm_sub:
                                                        redmine.issue.update(re_issue_id, notes='Error message : Subcomponent be Changed/Mismatch')
                                                        self.trigger_fail.emit('Error message : Subcomponent be Changed/Mismatch')
                                                    self.trigger.emit('Check completed.')
                                                except:
                                                    redmine.issue.update(re_issue_id, notes='Error message : Subcomponent be Changed/Mismatch')
                                                    self.trigger_fail.emit('Error message : Subcomponent be Changed/Mismatch')
                                            elif pims_status.name == 'Closed' or pims_status.name == 'Draft' or pims_status.name == 'cancel':
                                                #Status ID: 14, Status Name: Transfferred-Closed
                                                redmine.issue.update(re_issue_id, status_id = 14, custom_fields=[{'id': 157, 'value': 'Closed'}]) 
                                                self.trigger.emit('Due to the Jira status being "' + str( pims_status.name) + '", change the Redmine status to "Closed" successfully.')
                                                #刪除狀態
                                                df = df.drop(index_of_id_name)
                                                #檢查CSV有沒有id資訊,帶出來
                                                df.to_csv(new_path + '\\marker.csv', index=False)
                                            else:                  
                                                self.trigger_fail.emit("Error message : (Redmine id : " + str(re_issue_id) + " / Status : " + re_status + ");(" + pims + " / Status : " + pims_status.name + ")")        
                                        else:
                                            # Comments處理
                                            try:
                                                #20240306 fix
                                                comsun = 0
                                                if lc != "" and '[20xxxxxx EC xxx]' not in lc:
                                                    # 使用正規表達式找到 [EC status start] 到 [EC status end] 之間的字串並刪除
                                                    pattern = re.compile(r'\[EC status start\].*?\[EC status end\]', re.DOTALL)
                                                    # 用來加入comment的字串
                                                    lc_new = pattern.sub('', lc)
                                                    # 如果沒有評論
                                                    if not comment:
                                                        jira_client.add_comment(pims, str(lc_new))
                                                        self.trigger.emit('Update the comment as completed.')
                                                    # 如果有評論檢查最後一則評論是否一樣
                                                    elif len(comment) > 0:
                                                        for i in range(0,len(comment)):
                                                            if comment[i].body == lc_new:
                                                                comsun += 1
                                                        if comsun == 0:
                                                            jira_client.add_comment(pims, str(lc_new))
                                                            self.trigger.emit('Update the comment as completed.')
                                                        else:
                                                            self.trigger.emit("No update needed as the comment remains the same.")
                                            except:
                                                pass
                                            try:
                                                # 使用 re.search() 提取 [EC status start] 和 [EC status end] 之間的部分
                                                redmine_match = re.search(r'\[EC status start\](.*?)\[EC status end\]', lc, re.DOTALL)
                                                if redmine_match:
                                                    #用判斷裡面是不是空值
                                                    r_extracted_part_judge = redmine_match.group(1).strip()
                                                    #要加入的資料
                                                    r_extracted_part = redmine_match.group(0).strip()
                                                else:
                                                    r_extracted_part_judge = ''
                                                    r_extracted_part = ''
                                                # 如果jira executive summary不是空值
                                                if executive_summary != None:
                                                    # 使用正規表達式找到 [EC status start] 到 [EC status end] 之間的字串並刪除
                                                    pattern = re.compile(r'\[EC status start\].*?\[EC status end\]', re.DOTALL)
                                                    executive_summary_new = pattern.sub('', executive_summary)
                                                    #將jira 裡面的EC資訊抓取出來
                                                    jira_match = re.search(r'\[EC status start\](.*?)\[EC status end\]', executive_summary, re.DOTALL)
                                                    if jira_match:
                                                        j_extracted_part = jira_match.group(0).strip()
                                                    else:
                                                        j_extracted_part = ''
                                                    # 如果Redmine上不是空值
                                                    if r_extracted_part_judge != "":
                                                        # 判斷j_extracted_part是否和r_extracted_part
                                                        if j_extracted_part != r_extracted_part:
                                                            new_es = r_extracted_part + "\n" + executive_summary_new
                                                            pims_issue.update(customfield_28473 = str(new_es))
                                                            self.trigger.emit('Update the executive summary as completed.')
                                                elif r_extracted_part_judge != "":
                                                    pims_issue.update(customfield_28473 = str(r_extracted_part))
                                                    self.trigger.emit('Update the executive summary as completed.')
                                            except:
                                                pass
                                            
                                            if pims_status.name == 'Analyze' or pims_status.name == 'Wait' or pims_status.name == 'Review':
                                                try:
                                                    # 抓取當前PIMS Assignee 人員 (minse_yang)
                                                    assignee_name_now = pims_issue.fields.assignee.name
                                                    try:
                                                        # 當前PIMS不等於Redmine人員時修改為Redmine Assignee的人
                                                        if transferred_user_info != "":
                                                            if assignee_name_now != transferred_user_info:
                                                                if assignee_name_now != account_user_info:
                                                                    #先修改成自己在修改Redmine Assignee的人
                                                                    jira_client.assign_issue(pims, account_user_info)
                                                                jira_client.assign_issue(pims, transferred_user_info)
                                                                self.trigger.emit('Change the Assignee to "' + str(transferred_user_info) + '" successfully.')
                                                            else:
                                                                self.trigger.emit('The Assignee is the same person, do nothing.')
                                                            if pims_issue.fields.customfield_28471.name != transferred_user_info:
                                                                if pims_status.name == 'Analyze':
                                                                    jira_client.transition_issue(pims, 371, fields={'customfield_28471': {'name': transferred_user_info}})
                                                                elif pims_status.name == 'Wait':
                                                                    jira_client.transition_issue(pims, 601, fields={'customfield_28471': {'name': transferred_user_info}})
                                                                elif pims_status.name == 'Review':
                                                                    jira_client.transition_issue(pims, 621, fields={'customfield_28471': {'name': transferred_user_info}}) 
                                                                self.trigger.emit('Change the Analyst to "' + str(transferred_user_info) + '" successfully.')
                                                            else:
                                                                self.trigger.emit('The Analyst is the same person, do nothing.')
                                                        else:
                                                            self.trigger_fail.emit("Error message : Check if the 'Transferred Assignee For PQM' field is filled with the correct email.")
                                                    except:
                                                        self.trigger_fail.emit("Error message : Check if the 'Transferred Assignee For PQM' field is filled with the correct email.")
                                                except:
                                                    self.trigger_fail.emit("Error message : PIMS cannot locate the original assignee.")
                                                try:
                                                    sub_id_value = component_mapping.get(str(pqm_sub))
                                                    if pqm_sub not in pims_pqm_sub:
                                                        pims_issue.update(fields={'customfield_28436': [{'id':sub_id_value}]})
                                                        self.trigger.emit('Change the Component/Subcomponent to "' + str(pqm_sub) + '" successfully.')
                                                except:
                                                    self.trigger_fail.emit("Error message : Unable to retrieve content within Component/Subcomponent.")
                                            elif pims_status.name == 'Verify':
                                                self.trigger.emit("Do nothing!")
                                            else:
                                                self.trigger_fail.emit("Error message : (Redmine id : " + str(re_issue_id) + " / Status : " + re_status + ");(" + pims + " / Status : " + pims_status.name + ")")

                                    # Transferred-Closed
                                    elif re_status == 'Transferred-Closed' or re_status == 'ATS (P1/P2/P3)-Closed':
                                        #檢查CSV有沒有id資訊,帶出來
                                        #new_path = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
                                        new_path = os.getcwd()
                                        df = pd.read_csv(new_path + '\\marker.csv')
                                        try:
                                            #如果csv找尋到id就把modify_code帶出,代表有跑過一次
                                            index_of_id_name = int(df.index[df['Redmine_id'] == re_issue_id].values[0])
                                            #刪除狀態
                                            df = df.drop(index_of_id_name)
                                            df.to_csv(new_path + '\\marker.csv', index=False)
                                        except:
                                            pass
                                        if pims_status.name == 'Closed' or pims_status.name == 'Draft' or pims_status.name == 'cancel':
                                            self.trigger.emit("Do nothing!")
                                        else:
                                            self.trigger_fail.emit("Error message : (Redmine id : " + str(re_issue_id) + " / Status : " + re_status + ");(" + pims + " / Status : " + pims_status.name + ")")
                      
                                    # ATS/WAD-Can方法
                                    elif re_status == 'ATS/WAD - Can':
                                        #檢查CSV有沒有id資訊,帶出來
                                        #new_path = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
                                        new_path = os.getcwd()
                                        df = pd.read_csv(new_path + '\\marker.csv')
                                        try:
                                            #如果csv找尋到id就把modify_code帶出,代表有跑過一次
                                            index_of_id_name = int(df.index[df['Redmine_id'] == re_issue_id].values[0])
                                            modify_code = df['modify_code'][index_of_id_name]
                                            if modify_code !=3:
                                                try:
                                                    #刪除狀態
                                                    df = df.drop(index_of_id_name)
                                                    df.to_csv(new_path + '\\marker.csv', index=False)
                                                    #再添加一次進CSV
                                                    df.loc[df.index.max()+1] = [re_issue_id, 3]
                                                    df.to_csv(new_path + '\\marker.csv', index=False)
                                                    modify_code = 0
                                                except:
                                                    pass 
                                        except:
                                            #如果csv沒有id資訊就添加進去
                                            df.loc[df.index.max()+1] = [re_issue_id, 3]
                                            df.to_csv(new_path + '\\marker.csv', index=False)
                                            modify_code = 0
                                        # 檔案上傳
                                        if (pims_status.name == 'Review' or pims_status.name == 'Analyze' or pims_status.name == 'Wait') and redmine_document != "":
                                            redmine_attachment = redmine.attachment.get(redmine_document)
                                            attachment_content = redmine_attachment.download().content
                                            # 包裝成模擬文件
                                            with io.BytesIO(attachment_content) as attachment_file:
                                                # 上傳資料
                                                jira_attachments = jira_client.issue(pims).fields.attachment
                                                jira_attachment_list = [str(attachment.filename) for attachment in jira_attachments]
                                        
                                                if str(redmine_attachment.filename) not in jira_attachment_list:
                                                    jira_client.add_attachment(issue=pims, attachment=attachment_file, filename=str(redmine_attachment.filename))
                                                    self.trigger.emit('The file "' + str(redmine_attachment.filename) + '" has been successfully uploaded.')
                                                else:
                                                    self.trigger.emit('This file "' + str(redmine_attachment.filename) + '" already exists and will not be uploaded.')
                                        if modify_code == 3:
                                            if pims_status.name == 'Review':
                                                try:
                                                    id_value = option_mapping.get(str(dt_pqm))
                                                    if pims_dt_pqm != dt_pqm and id_value != None:
                                                        pims_issue.update(customfield_18716 = {'id': id_value})
                                                        self.trigger.emit('Change the Disposition type to "' + str(dt_pqm) + '" successfully.')
                                                    elif id_value == None:
                                                        self.trigger_fail.emit("Error message : Check if the Disposition Type is inconsistent. Redmine : " + str(dt_pqm) + ", Jira : " + str(pims_dt_pqm))
                                                    if pims_dd_pqm != dd_pqm:
                                                        pims_issue.update(customfield_18715 = str(dd_pqm))
                                                        self.trigger.emit('Change the Disposition Details to "' + str(dd_pqm) + '" successfully.')
                                                    if pims_trc_pqm != trc_pqm:
                                                        pims_issue.update(customfield_26700 = str(trc_pqm))
                                                        self.trigger.emit('Change the Technical root cause to "' + str(trc_pqm) + '" successfully.')
                                                    # ATS
                                                    if discretionary_field2 == '':
                                                        pims_issue.update(customfield_28468 = {'id': '120596'})
                                                        self.trigger.emit('Change the discretionary_field2 to "ATS Candidate" successfully.')
                                                    elif discretionary_field2 == 'ATS P1' or discretionary_field2 == 'ATS P2' or discretionary_field2 == 'ATS P3':
                                                        redmine.issue.update(re_issue_id, status_id = 16) #Status ID: 16, Status Name: ATS P1/P2/P3
                                                        redmine.issue.update(re_issue_id, notes='Change the status of Redmine to "ATS (P1/P2/P3)" successfully.')
                                                        self.trigger.emit('Change the status of Redmine to "ATS (P1/P2/P3)" successfully.')
                                                    elif discretionary_field2 == 'ATS P4' or discretionary_field2 == 'ATS P5' or discretionary_field2 == 'EB/WAD':
                                                        redmine.issue.update(re_issue_id, status_id = 17)  #Status ID: 17, Status Name: ATS P4/P5
                                                        redmine.issue.update(re_issue_id, notes='Change the status of Redmine to "ATS (P4/P5) / WAD" successfully.')
                                                        self.trigger.emit('Change the status of Redmine to "ATS (P4/P5) / WAD" successfully.')
                                                    elif discretionary_field2 != 'ATS P1' and discretionary_field2 != 'ATS P2' and discretionary_field2 != 'ATS P3' and discretionary_field2 != 'ATS P4' and discretionary_field2 != 'ATS P5' and discretionary_field2 != 'EB/WAD' and discretionary_field2 != 'ATS Candidate':
                                                        redmine.issue.update(re_issue_id, notes='DF2 is not expected!')
                                                        self.trigger_fail.emit("Error message : DF2 is not expected!")
                                                except:
                                                    self.trigger_fail.emit("Error message : During the status transition process, check for any issues with mandatory fields.(Disposition Type For PQM, Disposition Details For PQM, Technical Root Cause For PQM, ATS/WAD related document/mail)")

                                            else:
                                                self.trigger_fail.emit("Error message : (Redmine id : " + str(re_issue_id) + " / Status : " + re_status + ");(" + pims + " / Status : " + pims_status.name + ")")
                                        else:        
                                            # 如果PIMS為Analyze或Wait : 修改Ａｓｓｉｇｎｅｅ、　Disposition typｅ、　Disposition Details、 Technical root cause
                                            if pims_status.name == 'Analyze' or pims_status.name == 'Wait':
                                                try:
                                                    # 抓取當前PIMS Assignee 人員 (minse_yang)
                                                    assignee_name_now = pims_issue.fields.assignee.name
                                                    try:
                                                        id_value = option_mapping.get(str(dt_pqm))
                                                        # 先修改成Assign自己
                                                        jira_client.assign_issue(pims, account_user_info)
                                                        
                                                        # ATS
                                                        if discretionary_field2 == '':
                                                            pims_issue.update(customfield_28468 = {'id': '120596'})
                                                            self.trigger.emit('Change the discretionary_field2 to "ATS Candidate" successfully.')
                                                        elif discretionary_field2 == 'ATS P1' or discretionary_field2 == 'ATS P2' or discretionary_field2 == 'ATS P3':
                                                            redmine.issue.update(re_issue_id, status_id = 16) #Status ID: 16, Status Name: ATS P1/P2/P3
                                                            redmine.issue.update(re_issue_id, notes='Change the status of Redmine to "ATS (P1/P2/P3)" successfully.')
                                                            self.trigger.emit('Change the status of Redmine to "ATS (P1/P2/P3)" successfully.')
                                                        elif discretionary_field2 == 'ATS P4' or discretionary_field2 == 'ATS P5' or discretionary_field2 == 'EB/WAD':
                                                            redmine.issue.update(re_issue_id, status_id = 17) #Status ID: 17, Status Name: ATS P4/P5
                                                            redmine.issue.update(re_issue_id, notes='Change the status of Redmine to "ATS (P4/P5) / WAD" successfully.')
                                                            self.trigger.emit('Change the status of Redmine to "ATS (P4/P5) / WAD" successfully.')
                                                        elif discretionary_field2 != 'ATS P1' and discretionary_field2 != 'ATS P2' and discretionary_field2 != 'ATS P3' and discretionary_field2 != 'ATS P4' and discretionary_field2 != 'ATS P5' and discretionary_field2 != 'EB/WAD' and discretionary_field2 != 'ATS Candidate':
                                                            redmine.issue.update(re_issue_id, notes='DF2 is not expected!')
                                                            self.trigger_fail.emit("Error message : DF2 is not expected!")
                                                
                                                        if pims_issue.fields.issuetype.name == 'Change Request':
                                                            #pc_value = problem_mapping.get(str(pc_pqm)) not used
                                                            jira_client.transition_issue(pims, 61,fields={'customfield_18716': {"id":id_value},
                                                                                                          'customfield_18715': str(dd_pqm),
                                                                                                          'customfield_26700': str(trc_pqm),
                                                                                                          #'customfield_28480': {"id":pc_value}, not use
                                                                                                          })
                                                            # 轉換完後會讓assign變成Reporter的人,所以要再轉換一次
                                                            jira_client.assign_issue(pims, account_user_info)
                                                            jira_client.assign_issue(pims, assignee_name_now)
                                                            self.trigger.emit('Change the status of PIMS to "Review" successfully.')
                                                            self.trigger.emit('Change the Disposition type to "' + str(dt_pqm) + '" successfully.')
                                                            self.trigger.emit('Change the Disposition Details to "' + str(dd_pqm) + '" successfully.')
                                                            self.trigger.emit('Change the Technical root cause to "' + str(trc_pqm) + '" successfully.')
                                                            #self.trigger.emit('Change the Problem Category to "' + str(pc_pqm) + '" successfully.')   not use
                                                        else:
                                                            jira_client.transition_issue(pims, 61,fields={'customfield_18716': {"id":id_value},
                                                                                                          'customfield_18715': str(dd_pqm),
                                                                                                          'customfield_26700': str(trc_pqm),
                                                                                                          })
                                                            # 轉換完後會讓assign變成Reporter的人,所以要再轉換一次
                                                            jira_client.assign_issue(pims, account_user_info)
                                                            jira_client.assign_issue(pims, assignee_name_now)
                                                            self.trigger.emit('Change the status of PIMS to "Review" successfully.')
                                                            self.trigger.emit('Change the Disposition type to "' + str(dt_pqm) + '" successfully.')
                                                            self.trigger.emit('Change the Disposition Details to "' + str(dd_pqm) + '" successfully.')
                                                            self.trigger.emit('Change the Technical root cause to "' + str(trc_pqm) + '" successfully.')                            
                                                    except:
                                                        jira_client.assign_issue(pims, assignee_name_now)
                                                        self.trigger_fail.emit("Error message : During the status transition process, check for any issues with mandatory fields.(Disposition Type For PQM, Disposition Details For PQM, Technical Root Cause For PQM, ATS/WAD related document/mail)")
                                                except:
                                                    self.trigger_fail.emit("Error message : PIMS cannot locate the original assignee.")
                                            # 修改Disposition typｅ、　Disposition Details、 Technical root cause
                                            elif pims_status.name == 'Review':
                                                try:
                                                    id_value = option_mapping.get(str(dt_pqm))
                                                    if pims_dt_pqm != dt_pqm and id_value != None:
                                                        pims_issue.update(customfield_18716 = {'id': id_value})
                                                        self.trigger.emit('Change the Disposition type to "' + str(dt_pqm) + '" successfully.')
                                                    elif id_value == None:
                                                        self.trigger_fail.emit("Error message : Check if the Disposition Type is inconsistent. Redmine : " + str(dt_pqm) + ", Jira : " + str(pims_dt_pqm))
                                                    if pims_dd_pqm != dd_pqm:
                                                        pims_issue.update(customfield_18715 = str(dd_pqm))
                                                        self.trigger.emit('Change the Disposition Details to "' + str(dd_pqm) + '" successfully.')
                                                    if pims_trc_pqm != trc_pqm:
                                                        pims_issue.update(customfield_26700 = str(trc_pqm))
                                                        self.trigger.emit('Change the Technical root cause to "' + str(trc_pqm) + '" successfully.')
                                                    # ATS
                                                    if discretionary_field2 == '':
                                                        pims_issue.update(customfield_28468 = {'id': '120596'})
                                                        self.trigger.emit('Change the discretionary_field2 to "ATS Candidate" successfully.')
                                                    elif discretionary_field2 == 'ATS P1' or discretionary_field2 == 'ATS P2' or discretionary_field2 == 'ATS P3':
                                                        redmine.issue.update(re_issue_id, status_id = 16)
                                                        redmine.issue.update(re_issue_id, notes='Change the status of Redmine to "ATS (P1/P2/P3)" successfully.')
                                                        self.trigger.emit('Change the status of Redmine to "ATS (P1/P2/P3)" successfully.')
                                                    elif discretionary_field2 == 'ATS P4' or discretionary_field2 == 'ATS P5' or discretionary_field2 == 'EB/WAD':
                                                        redmine.issue.update(re_issue_id, status_id = 17)
                                                        redmine.issue.update(re_issue_id, notes='Change the status of Redmine to "ATS (P4/P5) / WAD" successfully.')
                                                        self.trigger.emit('Change the status of Redmine to "ATS (P4/P5) / WAD" successfully.')
                                                    elif discretionary_field2 != 'ATS P1' and discretionary_field2 != 'ATS P2' and discretionary_field2 != 'ATS P3' and discretionary_field2 != 'ATS P4' and discretionary_field2 != 'ATS P5' and discretionary_field2 != 'EB/WAD' and discretionary_field2 != 'ATS Candidate':
                                                        redmine.issue.update(re_issue_id, notes='DF2 is not expected!')
                                                        self.trigger_fail.emit("Error message : DF2 is not expected!")
                                                    self.trigger.emit('Check completed.')
                                                except:
                                                    self.trigger_fail.emit("Error message : During the status transition process, check for any issues with mandatory fields.(Disposition Type For PQM, Disposition Details For PQM, Technical Root Cause For PQM, ATS/WAD related document/mail)")
                                            # 如果PIMS不為Review, Show error 
                                            else:
                                                self.trigger_fail.emit("Error message : (Redmine id : " + str(re_issue_id) + " / Status : " + re_status + ");(" + pims + " / Status : " + pims_status.name + ")")
                                            
                                    # ATS [P4/P5] / WAD方法
                                    elif re_status == 'ATS (P4/P5) / WAD':
                                        #檢查CSV有沒有id資訊,帶出來
                                        #new_path = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
                                        new_path = os.getcwd()
                                        df = pd.read_csv(new_path + '\\marker.csv')
                                        try:
                                            #如果csv找尋到id就把modify_code帶出,代表有跑過一次
                                            index_of_id_name = int(df.index[df['Redmine_id'] == re_issue_id].values[0])
                                            #刪除狀態
                                            df = df.drop(index_of_id_name)
                                            df.to_csv(new_path + '\\marker.csv', index=False)
                                        except:
                                            pass
                                        
                                        if pims_status.name == 'Verify' or pims_status.name == 'Review' or pims_status.name == 'Wait':
                                            try:
                                                id_value = option_mapping.get(str(dt_pqm))
                                                if pims_dt_pqm != dt_pqm and id_value != None:
                                                    pims_issue.update(customfield_18716 = {'id': id_value})
                                                    self.trigger.emit('Change the Disposition type to "' + str(dt_pqm) + '" successfully.')
                                                elif id_value == None:
                                                    self.trigger_fail.emit("Error message : Check if the Disposition Type is inconsistent. Redmine : " + str(dt_pqm) + ", Jira : " + str(pims_dt_pqm))
                                                if pims_dd_pqm != dd_pqm:
                                                    pims_issue.update(customfield_18715 = str(dd_pqm))
                                                    self.trigger.emit('Change the Disposition Details to "' + str(dd_pqm) + '" successfully.')
                                                if pims_trc_pqm != trc_pqm:
                                                    pims_issue.update(customfield_26700 = str(trc_pqm))
                                                    self.trigger.emit('Change the Technical root cause to "' + str(trc_pqm) + '" successfully.')
                                                self.trigger.emit('Check completed.')
                                            except:
                                                self.trigger_fail.emit("Error message : During the status transition process, check for any issues with mandatory fields.(Disposition Type For PQM, Disposition Details For PQM, Technical Root Cause For PQM, ATS/WAD related document/mail)")

                                            # 檔案上傳
                                            if redmine_document != "":
                                                redmine_attachment = redmine.attachment.get(redmine_document)
                                                attachment_content = redmine_attachment.download().content
                                                # 包裝成模擬文件
                                                with io.BytesIO(attachment_content) as attachment_file:
                                                    # 上傳資料
                                                    jira_attachments = jira_client.issue(pims).fields.attachment
                                                    jira_attachment_list = [str(attachment.filename) for attachment in jira_attachments]
                                            
                                                    if str(redmine_attachment.filename) not in jira_attachment_list:
                                                        jira_client.add_attachment(issue=pims, attachment=attachment_file, filename=str(redmine_attachment.filename))
                                                        self.trigger.emit('The file "' + str(redmine_attachment.filename) + '" has been successfully uploaded.')
                                                    else:
                                                        self.trigger.emit('This file "' + str(redmine_attachment.filename) + '" already exists and will not be uploaded.')                         
                                        elif pims_status.name == 'cancel' or pims_status.name == 'Closed' or pims_status.name == 'Draft':
                                            try:
                                                redmine.issue.update(re_issue_id, custom_fields=[{'id': 56, 'value': 'Closed'}]) #cf_157 = Jira Status
                                                self.trigger.emit('Due to the Jira status being "' + str( pims_status.name) + '", change the Redmine status to "Closed" successfully.')
                                            except:
                                                self.trigger_fail.emit("Error message : During the status transition process, check for any issues with mandatory fields.(Disposition Type For PQM, Disposition Details For PQM, Technical Root Cause For PQM, ATS/WAD related document/mail)")
                                        else:
                                            self.trigger_fail.emit("Error message : (Redmine id : " + str(re_issue_id) + " / Status : " + re_status + ");(" + pims + " / Status : " + pims_status.name + ")")
                                   
                                    # Waiting方法
                                    elif re_status == 'Waiting':
                                        #檢查CSV有沒有id資訊,帶出來
                                        #new_path = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
                                        new_path = os.getcwd()
                                        df = pd.read_csv(new_path + '\\marker.csv')
                                        try:
                                            #如果csv找尋到id就把modify_code帶出,代表有跑過一次
                                            index_of_id_name = int(df.index[df['Redmine_id'] == re_issue_id].values[0])
                                            modify_code = df['modify_code'][index_of_id_name]
                                            if modify_code !=4:
                                                try:
                                                    #刪除狀態
                                                    df = df.drop(index_of_id_name)
                                                    df.to_csv(new_path + '\\marker.csv', index=False)
                                                    #再添加一次進CSV
                                                    df.loc[df.index.max()+1] = [re_issue_id, 4]
                                                    df.to_csv(new_path + '\\marker.csv', index=False)
                                                    modify_code = 0
                                                except:
                                                    pass 
                                        except:
                                            #如果csv沒有id資訊就添加進去
                                            df.loc[df.index.max()+1] = [re_issue_id, 4]
                                            df.to_csv(new_path + '\\marker.csv', index=False)
                                            modify_code = 0
                                        if modify_code == 4:
                                            if pims_status.name != 'Wait':
                                                self.trigger_fail.emit("Error message : (Redmine id : " + str(re_issue_id) + " / Status : " + re_status + ");(" + pims + " / Status : " + pims_status.name + ")")
                                            else:
                                                self.trigger.emit('Check completed.')
                                        else:
                                            try:
                                                # 抓取當前PIMS Assignee 人員 (minse_yang)
                                                assignee_name_now = pims_issue.fields.assignee.name
                                                # 如果PIMS狀態為Analyze就修改成wait
                                                if pims_status.name == 'Analyze':
                                                    jira_client.assign_issue(pims, account_user_info)
                                                    jira_client.transition_issue(pims, 51)
                                                # 當前PIMS不等於Redmine人員時修改為Redmine Assignee的人
                                                if re_user_info != "":
                                                    if assignee_name_now != re_user_info:
                                                        if assignee_name_now != account_user_info:
                                                            #先修改成自己在修改Redmine Assignee的人
                                                            jira_client.assign_issue(pims, account_user_info)
                                                        jira_client.assign_issue(pims, re_user_info)
                                                        self.trigger.emit('Change the Assignee to "' + str(re_user_info) + '" successfully.')
                                                    else:
                                                        self.trigger.emit('The Assignee is the same person, do nothing.')
                                                else:
                                                    self.trigger_fail.emit('Error message : Redmine id : '+ str(re_issue_id)  +' has not been assigned to anyone.')
                                                if pims_status.name != 'Analyze' and pims_status.name != 'Wait' and pims_status.name != 'Review':
                                                    self.trigger_fail.emit("Error message : (Redmine id : " + str(re_issue_id) + " / Status : " + re_status + ");(" + pims + " / Status : " + pims_status.name + ")")
                                            except:
                                               self.trigger_fail.emit("Error message : PIMS cannot locate the original assignee.")
                                    
                                    
                                    # PQM Issue Status
                                    try:
                                        pims_issue = jira_client.issue(pims)
                                        if issue_status_pqm != pims_issue.fields.status.name:
                                            redmine.issue.update(re_issue_id, custom_fields=[{'id': 157, 'value': str(pims_issue.fields.status.name)}])
                                            self.trigger.emit('Change the Redmine Jira Status to "' + str(pims_issue.fields.status.name) +  '" successfully.')
                                    except:
                                        pass
                                    break
                                    '''
                                    # Current BIOS Redmine doesn't have this custom field
                                    # Issue Severity
                                    try:
                                        if pims_issue_severity != issue_severity_pqm:
                                            redmine.issue.update(re_issue_id, custom_fields=[{'id': 5, 'value': str(pims_issue_severity)}])
                                            self.trigger.emit('Change the Redmine PQM Issue Severity to "' + str(pims_issue_severity) +  '" successfully.')
                                    except:
                                        pass
                                        
                                    break
                                    '''
                                except:
                                    if update_error == 0:
                                        self.trigger_wait.emit("Connection abnormal, waiting for the network to reconnect.")
                                        update_error+=1
                                    else:
                                        update_error+=1
                                    if update_error == 3:
                                        self.trigger_fail.emit("Error message : Redmine id : " + str(re_issue_id) + " / Status : " + re_status + ". Wait time exceeded. Please handle this issue manually.")
                                        break
                    else:
                        self.trigger_fail.emit("Error message : Redmine id : " + str(re_issue_id) + " / Status : " + re_status + " ,but there is no corresponding PIMS number.")
                except:
                    self.trigger_fail.emit("Error message : Redmine id : " + str(re_issue_id) + " / Status : " + re_status + ". Still need to confirm.")

                count_bar += count_pims
                self.countChanged.emit(int(round(count_bar,0)))
                self.trigger_pims.emit("------------------------------------------------")
        self.trigger_bt_on.emit()
        self.trigger.emit('EC part execution completed !')
        self.trigger_pims.emit("------------------------------------------------")
        self.trigger_pims.emit("Auto Run completes. The window closes automatically.")
        self.trigger_pims.emit("------------------------------------------------")
        self.trigger_pims.emit("------------------------------------------------")


if __name__ == '__main__':
    myapp = QApplication(sys.argv)
    
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    mywidget = MyWidget()
    mywidget.show()
    sys.exit(myapp.exec_())

# 查看搜索的內容
'''
for issue in search_results:
    print(f"Issue ID: {issue.id}, Subject: {issue.subject}, Status: {issue.status.name}, Priority: {issue.priority.name}")
'''
# CEP
'''
for link_id in pims_issue.fields.issuelinks:
    issue_link  = jira_client.issue_link(link_id)
    if hasattr(issue_link, 'inwardIssue'):
        inward_issue = jira_client.issue(issue_link.inwardIssue.key)
        print(f"Inward Issue: {inward_issue.key} - {inward_issue.fields.summary}")
    if hasattr(issue_link, 'outwardIssue'):
       outward_issue = jira_client.issue(issue_link.outwardIssue.key)
       print(f"Outward Issue: {outward_issue.key} - {outward_issue.fields.summary}")
issue_key1 = 'PIMS-291362'
issue_key2 = 'CEP-14875'

jira_client.create_issue_link(type='is Fixed by', inwardIssue=issue_key1, outwardIssue=issue_key2)
jira_client.issue_link_types()
'''
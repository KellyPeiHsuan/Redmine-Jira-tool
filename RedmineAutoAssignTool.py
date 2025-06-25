__author__ = "Compal SE"
__company__ = "Compal Electronics, Inc."
__copyright__ = "Copyright (c) 2025"
__version__ = "1.0"
__product__ = "Redmine Auto Assign Tool"

import requests
from requests_ntlm import HttpNtlmAuth
from redminelib import Redmine
import time
from datetime import datetime, timedelta
import configparser
import ctypes
import os
import json
import sys
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.application import MIMEApplication
import smtplib

sender_email = "a31autotool@gmail.com"
password = "owik tsop nvnc qgso"
receiver_email = "Darren_Chang@compal.com,aslan_chen@compal.com"
receiver_list = [email.strip() for email in receiver_email.split(',')]

if getattr(sys, 'frozen', False):
    base_path = os.path.dirname(sys.executable)  # PyInstaller 打包執行檔用
else:
    base_path = os.path.dirname(os.path.abspath(__file__))  # 一般 Python 執行

log_folder_path = os.path.join(base_path, "Log")

def get_latest_log_files(folder_path, n=1):
    try:
        files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
        files_sorted = sorted(files, key=os.path.getmtime, reverse=True)
        return files_sorted[:n]
    except Exception as e:
        print(f"找尋最新檔案失敗：{e}")
        return []

def send_email():
    latest_files = get_latest_log_files(log_folder_path, 1)
    if not latest_files:
        print("找不到最新的 log 檔案，停止寄信。")
        return

    message = MIMEMultipart()
    message["Subject"] = "Redmine Auto Assign tool Log"
    message["From"] = sender_email
    message["To"] = receiver_email

    body = MIMEText("此為Tool自動通知信請勿直接回覆此mail", "plain")
    message.attach(body)

    for file in latest_files:
        with open(file, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(file))
        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(file)}"'
        message.attach(part)

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_list, message.as_string())
        print("✅ 郵件寄出成功")
    except Exception as e:
        print(f"❌ 郵件寄出失敗: {e}")


# 後面函式就可以直接用這些變數


def handle_api_failure():
    msg = "API request failed. Please verify your connection or account configuration."
    print(msg)
    print(msg, file=log_file)
    show_error(msg)

    choice = input("\nDo you want to try again? (y/n): ").strip().lower()
    if choice == 'y':
        print("\n🔁 Restarting...\n")
        main()  # 重新執行 main
    else:
        print("\n❌ Exiting program.")
        sys.exit(1)
        
def show_error(message):
    ctypes.windll.user32.MessageBoxW(0, message, "Error", 0x10)

def get_category_id_by_name(project_id, category_name):
    if project_id not in category_cache:
        category_cache[project_id] = redmine.issue_category.filter(project_id=project_id)
    for c in category_cache[project_id]:
        if c.name.strip().lower() == category_name.strip().lower():
            return c.id
    return None

def process_issue(issue):
    if hasattr(issue, 'assigned_to'):
        return  # 已有 assignee

    print(f"Redmine ID: {issue.id}, Subject: {issue.subject}")
    print(f"Redmine ID: {issue.id}, Subject: {issue.subject}", file=log_file)
    
    subject = issue.subject
    if subject.startswith('['):
        subject = subject[subject.find(']') + 1:].lstrip()
    response = None
    for attempt in range(3):
        try:
            api_url = f'http://tperdvap2/Milano.api/Milano/searchCase?src=PIMS&startYear={milano_search_startyear}'
            response = requests.post(api_url, auth=HttpNtlmAuth(compal_account, compal_password), json=str(subject), timeout=20)
            if response.status_code == 200 and  response.json()['Result'] != []:
                break
        except Exception as e:
            print(f"Attempt {attempt+1} failed: {e}")
        time.sleep(1)

    if response is None or response.status_code != 200:
        handle_api_failure()  # 發生錯誤時呼叫 handle_api_failure，重新啟動程式

    ## 發現 Milano 回傳資料是空值
    data = response.json()
    if 'Result' not in data or not data['Result']:
        # 指派給預設好的null人選
        owner_null = assigned_dict['null'].strip()
        owner_id = user_dict.get(owner_null)
        issue.assigned_to_id = owner_id
        issue.notes = (
            f"API returned successfully, but no Milano data found. Assigned to default user '{owner_null}'. "
        )
        issue.save()
        msg = f"API returned successfully, but no Milano data found. Assigned to default user '{owner_null}'. "
        print(msg)
        print(msg, file=log_file)
        print("-" * 60)
        print("-" * 60, file=log_file)
        return
    
    ## 發現 Milano 找不到任何 PIMS
    pims_cases = [item['case_id'] for item in data['Result'] if 'PIMS' in item['case_id']]
    if not pims_cases:
        # 指派給預設好的null人選
        owner_null = assigned_dict['null'].strip()
        owner_id = user_dict.get(owner_null)
        issue.assigned_to_id = owner_id
        issue.notes = (
            f"No PIMS case found in Milano. Assigned to default user '{owner_null}'."
        )
        issue.save()
        msg = f"No PIMS case found in Milano. Assigned to default user '{owner_null}'."
        print(msg)
        print(msg, file=log_file)
        print("-" * 60)
        print("-" * 60, file=log_file)
        return

    category_list, pims_found = [], []
    for pims_id in pims_cases[:50]:
        search_results = redmine.issue.filter(status_id='!*', cf_1=pims_id)
        issue_search = next(iter(search_results), None)
        if issue_search and hasattr(issue_search, 'category') and issue_search.category:
            category_list.append(issue_search.category.name)
            pims_found.append(pims_id)
    
    ## 發現從 Milano 挑選的 PIMS 在 Redmine上查找不到使得 category_list 為空值
    if not category_list:
        # 指派給預設好的null人選
        owner_null = assigned_dict['null'].strip()
        owner_id = user_dict.get(owner_null)
        issue.assigned_to_id = owner_id
        issue.notes = (
            f"No PIMS case found in Redmine. Assigned to default user '{owner_null}'."
        )
        issue.save()
        msg = f"No PIMS case found in Redmine. Assigned to default user '{owner_null}'."
        print(msg)
        print(msg, file=log_file)
        print("-" * 60)
        print("-" * 60, file=log_file)
        return

    counts = {}
    for c in category_list:
        counts[c] = counts.get(c, 0) + 1
    sorted_counts = sorted(counts.items(), key=lambda x: x[1], reverse=True)

    print("Milano output:", pims_found)
    print("Milano output:", pims_found, file=log_file)
    print("Category classification results:", sorted_counts)
    print("Category classification results:", sorted_counts, file=log_file)
    
    # 如果都找不到對應的 assignee 或 category，指派給預設的 null 使用者
    assigned = False
    for category_name, _ in sorted_counts:
        key = category_name.strip().lower()
        for defined_cat in assigned_dict:
            if key == defined_cat.strip().lower():
                owner_name = assigned_dict[defined_cat].strip()
                owner_id = user_dict.get(owner_name)
                if not owner_id:
                    msg = f"Error Message: '{owner_name}' is not a valid user or does not exist in Redmine. Category is '{key}'."
                    print(msg)
                    print(msg, file=log_file)
                    break  

                category_id = get_category_id_by_name(issue.project.id, category_name)
                if not category_id:
                    break  

                issue.assigned_to_id = owner_id
                issue.category_id = category_id
                issue.notes = (
                    "Redmine present: " + str(pims_found) + "\n\n" +
                    "Category classification results: " + str(sorted_counts)
                )
                try:    
                    issue.save()
                    msg = f"Issue {issue.id} updated. Assigned to {owner_name}. Category is '{key}'"
                    print(msg)
                    print(msg, file=log_file)
                    print("-" * 60)
                    print("-" * 60, file=log_file)
                    assigned = True
                    return  
                    
                except Exception as e:
                    print(f"Error message : Failed to update Issue {issue.id}: {e}")
                    print(f"Error message : Failed to update Issue {issue.id}: {e}", file=log_file)
                    print("-" * 60)
                    print("-" * 60, file=log_file)
                    return  
            
    if not assigned:
        # 若都沒找到對應的 assignee 或 category，指派給預設的 null 使用者
        owner_null = assigned_dict['null'].strip()
        owner_id = user_dict.get(owner_null)
        issue.assigned_to_id = owner_id
        issue.notes = (
            "Redmine present: " + str(pims_found) + "\n\n" +
            "Category classification results: " + str(sorted_counts) + "\n\n" +
            f"Cannot find suitable assignee or category. Assigned to default user '{owner_null}'."
        )
        issue.save()
        msg = f"Cannot find suitable assignee or category. Assigned to default user '{owner_null}'."
        print(msg)
        print(msg, file=log_file)
        print("-" * 60)
        print("-" * 60, file=log_file)

def process_project_issues(pid):
    print(f"\n🔍 Processing project: {pid}")
    print(f"\n🔍 Processing project: {pid}", file=log_file)
    print("-" * 60)
    print("-" * 60, file=log_file)
    try:
        issues = redmine.issue.filter(project_id=pid, created_on=created_range, tracker_id=1, status_id=8)
        for issue in issues:
            process_issue(issue)
    except Exception as e:
        print(f"Error message : Failed to process {pid}: {e}")
        print(f"Error message : Failed to process {pid}: {e}", file=log_file)
        

# 主程式
def main():
    global redmine, category_cache, user_dict, assigned_dict, redmine_lookback_days
    global milano_search_startyear, compal_account, compal_password, created_range, log_file
# === 決定 base_path，適用打包與開發兩種狀態
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

        log_folder_path = os.path.join(base_path, "Log")

# === 寄信資訊（寫死，不使用 .ini）
        sender_email = "a31autotool@gmail.com"
        password = "owik tsop nvnc qgso"
        receiver_email = "aslan_chen@compal.com, Darren_Chang@compal.com,"
        receiver_list = [email.strip() for email in receiver_email.split(',')]

    # 載入 INI 設定
    config = configparser.ConfigParser()
    config.read(os.path.join(os.getcwd(), "Parameter.ini"))

    # 取得 project IDs，過濾空值
    project_ids = [
        config.get("Filter", "Project"),
        config.get("Filter", "Project_1"),
        config.get("Filter", "Project_2"),
    ]
    project_ids = [pid for pid in project_ids if pid.strip()]

    redmine_account = config.get("Redmine", "Account")
    redmine_password = config.get("Redmine", "Password")
    redmine_web = config.get("Web", "Redmine_url")

    milano_search_startyear = config.get("Milano", "Milano_search_startyear")
    compal_account = config.get("Milano", "Compal_account")
    compal_password = config.get("Milano", "Compal_password")
    assigned_dict = json.loads(config.get("Milano", "redmine_catagory_assigned_id"))
    redmine_lookback_days = config.get("Milano", "Redmine_Lookback_Days")

    # 建立 logs 資料夾與 log 檔案
    log_time = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
    log_dir = os.path.join(os.getcwd(), "Log")
    os.makedirs(log_dir, exist_ok=True)
    log_filename = os.path.join(log_dir, f"Redmine_assign_{log_time}.log")
    log_file = open(log_filename, 'w', encoding='utf-8')

    # 嘗試連接 Redmine
    try:
        redmine = Redmine(redmine_web, username=redmine_account, password=redmine_password)
        len(redmine.project.all()) # 判斷redmine是否有連接成功
    except:
        show_error("Error message : Connection to Redmine failed. Please check your network or server status.")
        return

    # 建立 user_dict
    unique_user_ids = set()
    for proj in redmine.project.all():
        try:
            for m in proj.memberships:
                if hasattr(m, 'user'):
                    unique_user_ids.add((m.user.name.strip(), m.user.id))
        except:
            continue
    user_dict = {name: uid for name, uid in unique_user_ids}

    # 建立日期範圍
    today = datetime.today().date()
    seven_days_ago = today - timedelta(days=int(redmine_lookback_days))
    created_range = f"><{seven_days_ago}|{today}"

    # 建立 category 快取
    category_cache = {}

    # 處理所有 project
    for pid in project_ids:
        process_project_issues(pid)

    log_file.close()
    send_email()

if __name__ == '__main__':
    main()
    input("\nProgram finished. Press Enter to exit...")


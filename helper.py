import calendar

import tkinter as tk
import tkinter.font as tkFont
import tkinter.ttk as ttk

import time
import datetime

import pytesseract
from datetime import date, timedelta, datetime
from PIL import Image

from tkcalendar import DateEntry, Calendar
from tkinter import messagebox
from tkinter import *

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

'''
回傳值列表:
    選擇的系統
    select_system: "NTU" or "NTNU"
    
    台大帳號密碼字典: NTU_account_dict
        身分別: ID
        帳號: account
        密碼: password
    台大借球場字典: NTU_court_dict
        Email:   email
        申請場地: court
        預約日期: date
        開始時間: start_time
        結束時間: end_time
        場地數量: court_num
        付款方式: payment_method
        收據編號: receipt
'''

# 師大帳號密碼、借場資訊字典
ntnu_dict = {"帳號" : "", "密碼" : "", "社團活動" : "", "場地類別名稱" : "", "場地名稱" : "", "預約日期" : "", "開始時間" : "", "結束時間" : "", "check" : False }
# 獲取現在年月日，方便日期選擇
now = datetime.now()
current_year = now.year
current_month = now.month
current_day = now.day

# 選擇台大或是師大系統頁面
class select_system_page(tk.Frame):
    
    # 建構子
    def __init__(self):
        tk.Frame.__init__(self)
        self.grid()
        self.master.title("借場小幫手")
        self.create_widget()

    # 建立介面元件
    def create_widget(self):
        f2 = tkFont.Font(size = 30, family = "Georgia")
        btn = tk.Button(self, text='台大預約球場系統', command=self.clickBtnNTU, font=f2)
        btn.grid(row=0, column=0)
        self.ntuBtn = btn
        btn = tk.Button(self, text='師大社團場地系統', command=self.clickBtnNTNU, font=f2)
        btn.grid(row=1, column=0)
        self.ntnuBtn = btn

    # 點選台大系統
    def clickBtnNTU(self):
        # 記憶選擇的系統，回傳後端
        global select_system
        self.destroy()
        select_system = "NTU"
        NTU_system_page()

    # 點選師大系統
    def clickBtnNTNU(self):
        global select_system
        self.destroy()
        select_system = "NTNU"
        NTNU_system_page()

# 模擬登入台大球場預約系統的動作
def ntu_login_id():
    if NTU_account_dict['ID'] == '眷屬/校友/校外':
        driver.find_element(By.XPATH, "//*[@id=\"__tab_ctl00_ContentPlaceHolder1_tcTab_TabPanel3\"]").click()
        account = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_tcTab_TabPanel3_txtMemberID"]')
        account.clear()
        account.send_keys(NTU_account_dict["account"])
        password = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_tcTab_TabPanel3_txtPassword"]') #網頁的密碼點
        password.clear()
        password.send_keys(NTU_account_dict["password"])
        driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_tcTab_TabPanel3_btnLogin"]').click()
    else:
        driver.find_element(By.XPATH, "//*[@id=\"ctl00_ContentPlaceHolder1_tcTab_tpValidator_HypLinkStu\"]").click()
        account = driver.find_element(By.XPATH, "//*[@id=\"myTable\"]/td/input")
        account.clear()
        account.send_keys(NTU_account_dict["account"])
        password = driver.find_element(By.XPATH, '//*[@id="myTable2"]/td/input') #網頁的密碼點
        password.clear()
        password.send_keys(NTU_account_dict["password"])
        driver.find_element(By.XPATH, "//*[@id=\"content\"]/form/table/tbody/tr[3]/td[2]/input").click()

# 檢查帳號密碼是否能夠登入
def ntu_check_login():
    driver.get('https://ntupesc.ntu.edu.tw/facilities/')
    ntu_login_id()
    temp_url = driver.current_url
    if str(temp_url) == "https://ntupesc.ntu.edu.tw/facilities/PlaceQuery.aspx":
        driver.close()
        return True
    else:
        return False

# 選擇身分別頁面
class NTU_system_page(tk.Frame):
    # 台大帳號密碼資料儲存字典:
    global NTU_account_dict
    NTU_account_dict = dict()

    def __init__(self):
        tk.Frame.__init__(self)
        self.grid()
        self.master.title("借場小幫手")
        self.choose_Identity()

    def choose_Identity(self):
        f1 = tkFont.Font(size = 20, family = "Georgia")
        f2 = tkFont.Font(size = 15, family = "Georgia")

        # 在 (0, 0) 的位置建立一個會顯示 '身分選擇' 的文字區塊
        self.ID_Label = tk.Label(self, text='身分選擇：', font=f1)
        self.ID_Label.grid(row=0, column=0, sticky='w')
        # 在 (0, 1) 的位置建立一個選擇 身分別 的下拉式選單
        IDLst = ['學生', '教職員', '眷屬/校友/校外']
        self.ID_entry = ttk.Combobox(self, values=IDLst, font=f2)
        self.ID_entry.grid(row=0, column=1, sticky = tk.NE + tk.SW)
        self.ID_entry.current(0)
        
        # 在 (1, 0) 的位置建立一個會顯示 '會員帳號' 的文字區塊
        self.account_Label = tk.Label(self, text='會員帳號：', font=f1)
        self.account_Label.grid(row=1, column=0, sticky='w')
        
        # 在 (1, 1) 的位置建立一個輸入 會員帳號 的欄位
        self.account_entry = tk.Entry(self, width=15, font=f2)
        self.account_entry.grid(row=1, column=1, sticky = tk.NE + tk.SW)

        # 在 (2, 0) 的位置建立一個會顯示 '會員密碼' 的文字區塊
        self.password_Label = tk.Label(self, text='會員密碼：', font=f1)
        self.password_Label.grid(row=2, column=0, sticky='w')
        
        # 在 (2, 1) 的位置建立一個輸入 會員密碼 的欄位
        self.password_entry = tk.Entry(self, width=15, font=f2, show='*')
        self.password_entry.grid(row=2, column=1, sticky = tk.NE + tk.SW)

        # 建一個框架 frame 放在 grid 的格子中
        button_frame = tk.Frame(self)
        button_frame.grid(row=3, column=1)

        # 在 frame 裡放「回上一頁」按鈕，靠左
        self.back = tk.Button(button_frame, text="回上一頁", font=f2, bg="silver", command=self.clickBtnBack)
        self.back.pack(side=tk.LEFT, padx=5)

        # 在 frame 裡放「確認」按鈕，靠右
        self.verify = tk.Button(button_frame, text="確認", font=f2, bg="silver", command=self.login)
        self.verify.pack(side=tk.LEFT, padx=5)

    # 回上一頁
    # 跳回select_system_page 並消除此視窗
    def clickBtnBack(self):
        self.master.destroy()
        select_system_page()
        

    # 嘗試登入，並連結到預約球場頁面
    # 事先檢查是否欄位空白，若有錯誤則顯示彈跳視窗
    def login(self):

        # 檢查ID選擇是否空白
        try:
            self.ID_entry.get()[0]
            NTU_account_dict['ID'] = self.ID_entry.get()
        except:
            self.error_page("請選擇身分")
        else:
            # 檢查帳號輸入是否空白
            try:
                self.account_entry.get()[0]
                NTU_account_dict['account'] = self.account_entry.get()
            except:
                self.error_page("請輸入帳號")
            else:
                # 檢查密碼輸入是否空白
                try:
                    self.password_entry.get()[0]
                    NTU_account_dict['password'] = self.password_entry.get()
                except:
                    self.error_page("請輸入密碼")
                else:
                    # 嘗試登入系統
                    # 如果登入失敗，跳出錯誤視窗，並請使用者重新輸入
                    self.login_result = ntu_check_login()
                    if self.login_result == False:
                        self.error_page("帳號密碼錯誤，請重新輸入")

                    else:
                        self.master.destroy()
                        court_entry_page()

    # 出現錯誤的彈跳視窗，title 在 mac 不會出現
    def error_page(self, error_message):
        self.error = tk.messagebox.showerror(title="錯誤訊息", message=error_message)



# 台大系統借球場頁面
'''
         col=0       col=1       col=2       col=3
row = 0  Email：
row = 1  申請場地：　　　　　　（下拉式選單）
row = 2  預約日期：  選擇日期
row = 3  起訖時間：  開始時間   　　　~   　　結束時間   
row = 4  場地數量：
row = 5  付款方式：（下拉式選單：現金、時數券）
row = 6  收據編號：　　　　　　　　　　　　　（付款方式若為時數券才需填寫）　　
row = 7            回上一頁　    　確認
'''
class court_entry_page(tk.Frame):
    # 台大借球場資料儲存字典
    global NTU_court_dict
    NTU_court_dict = dict()
    
    def __init__(self):
        tk.Frame.__init__(self)
        self.grid()
        self.master.title("借場小幫手")
        self.courtEntry()

    def courtEntry(self):
        f1 = tkFont.Font(size = 20, family = "Georgia")
        f2 = tkFont.Font(size = 15, family = "Georgia")
        f3 = tkFont.Font(size = 12, family = "Georgia")

        # 在 (1, 0) 的位置建立一個會顯示 'Email：' 的文字區塊
        self.email_Label = tk.Label(self, text='Email：', font=f1)
        self.email_Label.grid(row=1, column=0, sticky='w')
        # 在 (1, 1) 的位置建立一個輸入 Email 的欄位
        self.email_entry = tk.Entry(self, width=15, font=f2)
        self.email_entry.grid(row=1, column=1, columnspan = 3, sticky = tk.NE + tk.SW)


        # 在 (2, 0) 的位置建立一個會顯示 '申請場地' 的文字區塊
        self.court_Label = tk.Label(self, text='申請場地：', font=f1)
        self.court_Label.grid(row=2, column=0,sticky='w')
        # 在 (2, 1) 的位置建立一個選擇 球場 的下拉式選單
        courtLst = ['3F羽球場', '1F羽球場', 'B1桌球室', 'B1壁球室', 'B109教室(桌球)']        
        self.court_entry = ttk.Combobox(self, values=courtLst, font=f2)
        self.court_entry.grid(row=2, column=1, columnspan = 3, sticky = tk.NE + tk.SW)
        self.court_entry.current(2)

        # 在 (3, 0) 的位置建立一個會顯示 '預約日期' 的文字區塊
        self.date_Label = tk.Label(self, text='預約日期：', font=f1)
        self.date_Label.grid(row=3, column=0, sticky='w')
        # 在 (3, 1) 的位置建立一個日期選擇器按鈕
        self.BtnDate = tk.Button(self, text="選擇日期", font=f2, bg="silver", command = self.clickBtnDate)
        self.BtnDate.grid(row=3, column=1, columnspan = 3, sticky = tk.NE + tk.SW)

        
        # 在 (4, 0) 的位置建立一個會顯示 '起訖時間' 的文字區塊
        self.Label4 = tk.Label(self, text='起訖時間：', font=f1)
        self.Label4.grid(row=4, column=0, sticky='w')
        # 在 (4, 1) 的位置建立一個選擇 時間 的下拉式選單
        timeLst = ['8:00', '9:00', '10:00', '11:00', '12:00', '13:00', '14:00', '15:00',
                   '16:00', '17:00', '18:00', '19:00', '20:00', '21:00', '22:00']
        # 開始時間的下拉式選單
        self.startTime = ttk.Combobox(self, values=timeLst, font=f2)
        self.startTime.grid(row=4, column=1, sticky = tk.NE + tk.SW)
        # 時間幾點到幾點的符號
        self.spanLabel = tk.Label(self, text='～', font=f2)
        self.spanLabel.grid(row=4, column=2)
        # 結束時間的下拉式選單
        self.endTime = ttk.Combobox(self, values=timeLst,font=f2)
        self.endTime.grid(row=4, column=3, sticky = tk.NE + tk.SW)

        # 在 (5, 0) 的位置建立一個會顯示 '場地數量' 的文字區塊
        self.courtnum_Label = tk.Label(self, text='場地數量：', font=f1)
        self.courtnum_Label.grid(row=5, column=0, sticky='w')
        # 在 (5, 1) 的位置建立一個選擇 場地數量 下拉式選單
        courtnumLst = list()
        for i in range(15):
            courtnumLst.append(i+1)
        self.courtnum_entry = ttk.Combobox(self, values=courtnumLst, font=f2)
        self.courtnum_entry.grid(row=5, column=1, columnspan = 3, sticky = tk.NE + tk.SW)

        # 在 (6, 0) 的位置建立一個會顯示 '付款方式' 的文字區塊
        self.payment_Label = tk.Label(self, text='付款方式：', font=f1)
        self.payment_Label.grid(row=6, column=0, sticky='w')
        # 在 (6, 1) 的位置建立一個選擇 時間 的下拉式選單
        paymentLst = ["現金", "時數券"]
        # 付款方式的下拉式選單
        self.payment = ttk.Combobox(self, values=paymentLst, font=f2)
        self.payment.grid(row=6, column=1, columnspan=3, sticky = tk.NE + tk.SW)
        self.payment.current(0)

        # 在 (7, 0) 的位置建立一個會顯示 '收據編號' 的文字區塊
        self.receipt_Label = tk.Label(self, text='收據編號：', font=f1)
        self.receipt_Label.grid(row=7, column=0, sticky='w')
        # 在 (7, 1) 的位置建立一個輸入 收據編號 的欄位
        self.receipt_entry = tk.Entry(self, width=15, font=f2)
        self.receipt_entry.grid(row=7, column=1, columnspan = 2, sticky = tk.NE + tk.SW)
        # 在 (7, 1) 的位置建立一個備註的文字區塊
        self.remind_Label = tk.Label(self, text="(付款方式若為時數券才需填寫)", font=f3, fg="red")
        self.remind_Label.grid(row=7, column=3, columnspan = 1, sticky = tk.NE + tk.SW)

        # 在 (8, 0) 的位置建立一個'回上一頁'的按鈕
        # 按鈕按下後會跳回 NTU_system_page
        self.backBtn = tk.Button(self, text="回上一頁", font=f2, bg="silver", command = self.clickBtnBack)
        self.backBtn.grid(row=8, column=1)

        # 在 (8, 1) 的位置建立一個確認的按鈕
        # 按鈕按下後會執行 login() 函式
        self.verify = tk.Button(self, text="確認", font=f2, bg="silver", command = self.login)
        self.verify.grid(row=8, column=3, sticky='W')

    # 回上一頁
    def clickBtnBack(self):
        self.master.destroy()
        NTU_system_page()

    # 按下選擇日期的按鈕
    def clickBtnDate(self):
        f1 = tkFont.Font(size = 20, family = "Georgia")
        # 去掉選擇日期的按鈕
        self.BtnDate.destroy()
        # 出現日期選擇器
        self.calendar_window()

    # 日期選擇器
    def calendar_window(self):
        global cal
        cal = Calendar(self, selectmode = 'day', year = current_year, month = current_month,day = current_day)
        cal.grid(row=3, column=1)
        self.verifyBtn = tk.Button(self, text = "確認", command = self.grab_date)
        self.verifyBtn.grid(row=3, column=2)


    def grab_date(self):
        f1 = tkFont.Font(size = 20, family = "Georgia")
        global select_date
        select_date = cal.get_date()
        cal.destroy()
        self.verifyBtn.destroy()
        
        # 把選擇的日期變成label
        select_date = select_date.split('/')
        self.select_year = "20" + select_date[2]
        if len(select_date[0]) == 1:
            self.select_month = '0' + select_date[0]
        else:
            self.select_month = select_date[0]
        
        if len(select_date[1]) == 1:
            self.select_day = '0' + select_date[1]
        else:
            self.select_day = select_date[1]
        
        datestr = self.select_year + '/' + self.select_month + '/' + self.select_day
        self.dateLabel = tk.Label(self, text=datestr, font=f1)
        self.dateLabel.grid(row=3, column=1, columnspan = 3, sticky = "w")

        # 在(3,2)新增'重新選擇'的按鈕
        self.select_again = tk.Button(self, text="重新選擇", command = self.click_select_again)
        self.select_again.grid(row=3, column=2)

    def click_select_again(self):
        self.select_again.destroy()
        self.dateLabel.destroy()
        self.clickBtnDate()


    # 按下確認按鈕
    def login(self):
        try:
            self.email_entry.get()[0]
            NTU_court_dict['email'] = self.email_entry.get()
        except:
            self.error_page("請填寫Email")
        else:
            # 檢查申請場地
            try:
                self.court_entry.get()[0]
                NTU_court_dict['court'] = self.court_entry.get()
            except:
                self.error_page("請選擇場地")
            else:
                # 檢查日期是否空白
                try:
                    self.dateLabel.cget("text")[0]
                    NTU_court_dict['date'] = self.dateLabel.cget("text")
                    # 檢查日期是否小於可預約日期(現在日期減兩天)
                    self.current_date = date(current_year, current_month, current_day)
                    self.select_date = date(int(self.select_year), int(self.select_month), int(self.select_day))
                    if self.current_date >= self.select_date + timedelta(days=-1):
                        raise ValueError
                        
                except AttributeError:
                    self.error_page("請選擇日期")
                except ValueError:
                    self.error_page("此日期無法預約，請重新選擇日期")
                else:
                    # 檢查開始時間是否空白
                    try:
                        self.startTime.get()[0]
                        NTU_court_dict['start_time'] = self.startTime.get()
                    except:
                        self.error_page("請選擇開始時間")
                    else:
                        # 檢查結束時間是否空白
                        try:
                            self.endTime.get()[0]
                            NTU_court_dict['end_time'] = self.endTime.get()
                        except:
                            self.error_page("請選擇結束時間")
                        else:
                            # 檢查開始時間是否小於等於結束時間
                            try:
                                if time.strptime(NTU_court_dict['start_time'], "%H:%M") >= time.strptime(NTU_court_dict['end_time'], "%H:%M"):
                                    raise ValueError
                            except ValueError:
                                self.error_page("時段不正確，請重新選擇時段")
                            else:
                                # 檢查場地數量輸入是否空白
                                try:
                                    self.courtnum_entry.get()[0]
                                    NTU_court_dict['court_num'] = self.courtnum_entry.get()
                                except:
                                    self.error_page("請選擇場地數量")
                                else:
                                    # 檢查付款方式是否空白
                                    try:
                                        self.payment.get()[0]
                                        NTU_court_dict['payment_method'] = self.payment.get()
                                    except:
                                        self.error_page("請選擇付款方式")
                                    else:
                                        if NTU_court_dict['payment_method'] != "時數券":
                                            self.master.destroy()
                                            self.check_reservation_result()
                                        # 如果付款方式選擇時數券，檢查收據是否空白
                                        if NTU_court_dict['payment_method'] == "時數券":
                                            try:
                                                self.receipt_entry.get()[0]
                                                NTU_court_dict['receipt'] = self.receipt_entry.get()
                                                self.master.destroy()
                                            except:
                                                self.error_page("請輸入收據編號")
                                            else:
                                                self.check_reservation_result()

    # 出現錯誤的彈跳視窗
    def error_page(self, error_message):
        self.error = tk.messagebox.showerror(title="錯誤訊息", message=error_message)

        # 判斷借場是否成功
    def check_reservation_result(self):
        # 如果True，回傳成功借到幾場
        # 如果False，回傳失敗
        self.reservation_result = ntu_schedule_book()
        if self.reservation_result[0] == True:
            show_result_string = '您已成功預約' + str(NTU_court_dict['start_time']) + '至' + NTU_court_dict['end_time'] + ' ' + str(NTU_court_dict['court']) + ' ' + '共' + str(self.reservation_result[1]) + '場'
            self.result = tk.messagebox.showinfo(title='預約場地結果', message=show_result_string)
        else:
            show_result_string = '很抱歉，您沒有成功預約到場地'
            self.result = tk.messagebox.showinfo(title='預約場地結果', message=show_result_string)       

###########################################################################################################

# 師大系統頁面: 僅能借社團場地
class NTNU_system_page(tk.Frame):
    def __init__(self):
        tk.Frame.__init__(self)
        self.grid()
        self.master.title("借場小幫手")
        self.NTNU_system()
    
    def NTNU_system(self):
        f1 = tkFont.Font(size = 20, family = "Georgia")
        f2 = tkFont.Font(size = 15, family = "Georgia")
        
        # 在 (0, 0) 的位置建立一個會顯示 '帳號' 的文字區塊
        self.label0 = tk.Label(self, text='帳號：', font=f1).grid(row=0, column=0)
        # 在 (0, 1) 的位置建立一個輸入 帳號 的欄位
        self.entry0 = tk.Entry(self, width=15, font=f2)
        self.entry0.grid(row=0, column=1)
        
        # 在 (1, 0) 的位置建立一個會顯示 '密碼' 的文字區塊
        self.Label1 = tk.Label(self, text='密碼：', font=f1).grid(row=1, column=0)
        # 在 (1, 1) 的位置建立一個輸入 密碼 的欄位
        self.entry1 = tk.Entry(self, width=15, font=f2, show='*')
        self.entry1.grid(row=1, column=1)
        
        button_frame2 = tk.Frame(self)
        button_frame2.grid(row=3, column=1)

        # 在 (3, 1) 的位置建立一個'回上一頁'的按鈕，
        # 按鈕按下後會跳回 select_system_page
        self.back2 = tk.Button(button_frame2, text="回上一頁", font=f2, bg="silver", command=self.clickBtnBack)
        self.back2.pack(side=tk.LEFT, padx=5)

        # 在 (3, 1) 的位置建立一個確認的按鈕，
        # 按鈕按下後會執行 login() 函式
        self.verify2 = tk.Button(button_frame2, text="確認", font=f2, bg="silver", command=self.login)
        self.verify2.pack(side=tk.LEFT, padx=5)
    
    # 回上一頁
    # 跳回select_system_page 並消除此視窗
    def clickBtnBack(self):
        self.master.destroy()
        select_system_page()
    
    def login(self):
        # 檢查 帳號 選擇是否空白
        if len(self.entry0.get()) == 0:
            self.error_page("請輸入帳號")
        # 檢查 密碼 是否空白
        elif len(self.entry1.get()) == 0:
            self.error_page("請輸入密碼")
        else:
            # 嘗試登入
            ntnu_dict["帳號"] = self.entry0.get()
            ntnu_dict["密碼"] = self.entry1.get()
            self.login_result = ntnu_check_login(ntnu_dict)
            if self.login_result == False:
                self.error_page("帳號密碼錯誤，請重新輸入")
            else:
                self.master.destroy()
                ntnu_info_page()

    # 出現錯誤的彈跳視窗
    def error_page(self, error_message):
        self.error = tk.messagebox.showerror(title="錯誤訊息", message=error_message)
    
def ntnu_check_login(ntnu_dict):
    # 連到網站
    driver.get('http://iportal.ntnu.edu.tw/ntnu/')
    # 填入帳號
    account = driver.find_element_by_xpath('//*[@id="muid"]')
    account.clear()
    account.send_keys(ntnu_dict["帳號"])
    # 填入密碼
    password = driver.find_element_by_xpath('//*[@id="mpassword"]')
    password.clear()
    password.send_keys(ntnu_dict["密碼"])
    # 按下確認
    driver.find_element_by_xpath('/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/table/tbody/tr[2]/td/table[1]/tbody/tr[2]/td/div/input').click()
    if str(driver.current_url) == 'https://iportal.ntnu.edu.tw/login.do':
        return False
    else:
        return True
    driver.close()

# 師大借場資訊
'''
        col=0  col=1  col=2  col=3
row = 0 社團活動
row = 1 場地類別名稱
row = 2 場地名稱
row = 3 預約日期
row = 4 開始時間
row = 5 結束時間
'''

class ntnu_info_page(tk.Frame):
    def __init__(self):
        tk.Frame.__init__(self)
        self.grid()
        self.master.title("借場小幫手")
        self.ntnu_info()
    # 師大系統--社團
    def ntnu_info(self):
        f1 = tkFont.Font(size = 20, family = "Georgia")
        f2 = tkFont.Font(size = 15, family = "Georgia")
        f3 = tkFont.Font(size = 12, family = "Georgia")
        # 取得使用者的輸入
        def func(selected_value):
            room.set_menu(*roomLst.get(selected_value))
            startTime.set_menu(*startTimeLst.get(selected_value))
            endTime.set_menu(*endTimeLst.get(selected_value))
        # 儲存使用者的輸入
        def store_building(*args):
            ntnu_dict["場地類別名稱"] = var1.get()
        def store_room(*args):
            ntnu_dict["場地名稱"] = var2.get()
        def store_start(*args):
            ntnu_dict["開始時間"] = var3.get()
        def store_end(*args):
            ntnu_dict["結束時間"] = var4.get()
        
        # 在 (0, 0) 的位置建立一個會顯示 社團活動 的文字區塊
        self.Label0 = tk.Label(self, text='社團活動：', font=f1)
        self.Label0.grid(row=0, column=0, sticky='w')
        # 在 (0, 1) 的位置建立一個輸入 社團活動 的欄位
        self.entry0 = tk.Entry(self, width=15, font=f2)
        self.entry0.grid(row=0, column=1, columnspan = 3, sticky = tk.NE + tk.SW)
        
        # 在 (1, 0) 的位置建立一個會顯示 場地類別名稱 的文字區塊
        self.Label1 = tk.Label(self, text='場地類別名稱：', font=f1)
        self.Label1.grid(row=1, column=0, sticky='w')
        # 在 (1, 1) 的位置建立一個選擇 場地類別 的下拉式選單
        buildingLst = ["-",'綜合大樓3、6樓活動室', '誠正大樓地下室', '平日晚間樸大樓']        
        var1 = StringVar()
        myStyle=ttk.Style()
        myStyle.configure('my.TMenubutton',font=('Georgia',15))
        building = ttk.OptionMenu(self, var1, *buildingLst, command=func, style='my.TMenubutton')
        var1.trace_add('write', store_building)
        building['menu'].configure(font=('Georgia',15))
        building.grid(row=1, column=1, columnspan = 3, sticky = tk.NE + tk.SW)
        
        # 在 (2, 0) 的位置建立一個會顯示 場地名稱 的文字區塊
        self.Label1 = tk.Label(self, text='場地名稱：', font=f1)
        self.Label1.grid(row=2, column=0, sticky='w')
        # 在 (2, 1) 的位置建立一個選擇 場地名稱 的下拉式選單
        roomLst = {"綜合大樓3、6樓活動室" : ['綜302', '綜303', '綜304', '綜307', '綜6A舞蹈教室', '綜6B'], "誠正大樓地下室" : ['誠正大樓地下室A區', '誠正大樓地下室B區', '誠正大樓地下室C區', '誠正大樓地下室D區', '誠正大樓地下室E區'], "平日晚間樸大樓" : ['樸105', '樸106', '樸202', '樸203', '樸204', '樸205', '樸206', '樸301', '樸302', '樸303', '樸304', '樸305', '樸306', '樸401', '樸402', '樸403', '樸404', '樸405', '樸406', '樸407']}
        var2 = StringVar()
        myStyle=ttk.Style()
        myStyle.configure('my.TMenubutton',font=('Georgia',15))
        room = ttk.OptionMenu(self, var2, "-", style='my.TMenubutton')
        var2.trace_add('write', store_room)
        room['menu'].configure(font=('Georgia',15))
        room.grid(row=2, column=1, columnspan = 3, sticky = tk.NE + tk.SW)
        
        # 在 (3, 0) 的位置建立一個會顯示 '預約日期' 的文字區塊
        self.Label3 = tk.Label(self, text='預約日期：', font=f1)
        self.Label3.grid(row=3, column=0, sticky='w')
        # 在 (3, 1) 的位置建立一個日期選擇器按鈕
        self.BtnDate = tk.Button(self, text="選擇日期", font=f2, bg="silver", command = self.clickBtnDate)
        self.BtnDate.grid(row=3, column=1, columnspan = 3, sticky = tk.NE + tk.SW)
        
        # 在 (4, 0) 的位置建立一個會顯示 '開始時間' 的文字區塊
        self.Label4 = tk.Label(self, text='開始時間：', font=f1)
        self.Label4.grid(row=4, column=0, sticky='w')
        # 在 (4, 1) 的位置建立一個選擇 開始時間 的下拉式選單
        startTimeLst = {"綜合大樓3、6樓活動室" : ['請選擇', '08:00', '12:00', '14:00','18:00'], "誠正大樓地下室" : ['請選擇', '08:00', '12:00', '14:00','18:00'], "平日晚間樸大樓" : ['請選擇', '18:30']}
        var3 = StringVar()
        myStyle=ttk.Style()
        myStyle.configure('my.TMenubutton',font=('Georgia',15))
        startTime = ttk.OptionMenu(self, var3, "-", style='my.TMenubutton')
        var3.trace_add('write', store_start)
        startTime['menu'].configure(font=('Georgia',15))
        startTime.grid(row=4, column=1, columnspan = 3, sticky = tk.NE + tk.SW)
        
        # 在 (5, 0) 的位置建立一個會顯示 '結束時間' 的文字區塊
        self.Label4 = tk.Label(self, text='結束時間：', font=f1)
        self.Label4.grid(row=5, column=0, sticky='w')
        # 在 (5, 1) 的位置建立一個選擇 結束時間 的下拉式選單
        endTimeLst = {"綜合大樓3、6樓活動室" : ['請選擇', '12:00', '14:00', '18:00','22:00'], "誠正大樓地下室" : ['請選擇', '12:00', '14:00', '18:00','22:00'], "平日晚間樸大樓" : ['請選擇', '22:15']}
        var4 = StringVar()
        myStyle=ttk.Style()
        myStyle.configure('my.TMenubutton',font=('Georgia',15))
        endTime = ttk.OptionMenu(self, var4, "-", style='my.TMenubutton')
        var4.trace_add('write', store_end)
        endTime['menu'].configure(font=('Georgia',15))
        endTime.grid(row=5, column=1, columnspan = 3, sticky = tk.NE + tk.SW)
        
        # 在 (6, 0) 的位置建立一個'回上一頁'的按鈕
        # 按鈕按下後會跳回 NTU_system_page
        self.backBtn = tk.Button(self, text="回上一頁", font=f2, bg="silver", command = self.clickBtnBack)
        self.backBtn.grid(row=6, column=0)
        # 在 (6, 1) 的位置建立一個確認的按鈕，
        # 按鈕按下後會執行 login() 函式
        self.verify = tk.Button(self, text="確認", font=f2, bg="silver", command = self.login)
        self.verify.grid(row=6, column=1, columnspan=2)

        
    # 回上一頁
    # 跳回select_system_page 並消除此視窗
    def clickBtnBack(self):
        self.master.destroy()
        NTNU_system_page()
    
    # 將輸入的資料傳入網站相對應的位置
    # 檢查輸入是否為空或合理
    def login(self):
        check_start = ntnu_dict["開始時間"]
        check_start = check_start.replace(":", "")
        check_end = ntnu_dict["結束時間"]        
        check_end = check_end.replace(":", "")
        if ntnu_dict["預約日期"] != "":           
            self.current_date = date(current_year, current_month, current_day)
            self.select_date = date(int(self.select_year), int(self.select_month), int(self.select_day))
        # 檢查 社團活動名稱 是否為空
        if self.entry0.get() == "":
            self.error_page("請填入社團活動名稱")
        # 檢查 場地類別名稱 是否為空
        elif ntnu_dict["場地類別名稱"] == "":
            self.error_page("請選擇場地類別名稱")
        # 檢查 預約日期 是否為空            
        elif ntnu_dict["預約日期"] == "":
            self.error_page("請選擇日期")
        # 檢查日期是否小於可預約日期(現在
        elif self.current_date >= self.select_date:
            self.error_page("此日期錯誤，請重新選擇日期")  
        # 檢查 時段 是否合理
        elif int(check_start) > int(check_end):
            self.error_page("此時段錯誤，請重新選擇時段")            
        else:
            ntnu_dict["社團活動"] = self.entry0.get()
            ntnu_dict["check"] = True 
            self.master.destroy()
            ntnu_book(ntnu_dict)   # 執行師大預約
    
    # 出現錯誤的彈跳視窗
    def error_page(self, error_message):
        self.error = tk.messagebox.showerror(title="錯誤訊息", message=error_message)
    
    # 按下選擇日期的按鈕
    def clickBtnDate(self):
        f1 = tkFont.Font(size = 20, family = "Georgia")
        # 去掉選擇日期的按鈕
        self.BtnDate.destroy()
        # 出現日期選擇器
        self.calendar_window()
    
    # 日期選擇器
    def calendar_window(self):
        global cal
        cal = Calendar(self, selectmode = 'day', year = current_year, month = current_month,day = current_day)
        cal.grid(row=3, column=1)
        self.verifyBtn = tk.Button(self, text = "確認", command = self.grab_date)
        self.verifyBtn.grid(row=3, column=2)


    def grab_date(self):
        f1 = tkFont.Font(size = 20, family = "Georgia")
        global select_date
        select_date = cal.get_date()
        cal.destroy()
        self.verifyBtn.destroy()
        
        # 把選擇的日期變成label
        select_date = select_date.split('/')
        self.select_year = "20" + select_date[2]
        self.select_month = select_date[0]
        self.select_day = select_date[1]
        datestr = self.select_year + '-' + self.select_month + '-' + self.select_day
        ntnu_dict["預約日期"] = datestr
        self.dateLabel = tk.Label(self, text=datestr, font=f1)
        self.dateLabel.grid(row=3, column=1, columnspan = 3, sticky = "w")

        # 在(3,4)新增'重新選擇'的按鈕
        self.select_again = tk.Button(self, text="重新選擇", command = self.click_select_again)
        self.select_again.grid(row=3, column=4)

    def click_select_again(self):
        self.select_again.destroy()
        self.clickBtnDate()

# 台大球場預約後端
def ntu_book():
    driver = webdriver.Chrome(service=s)
    driver.get('https://ntupesc.ntu.edu.tw/facilities/')
    def check_remain_number():
        window_before = driver.current_window_handle
        # 檢驗跨到的每個時段是否有足夠數量
        try:
            for j in NTU_court_dict['timeindex']:
                check_path = f"//*[@id=\"ctl00_ContentPlaceHolder1_tab1\"]/tbody/tr[{str(j)}]/td[{NTU_court_dict['weekday']}]/a"
                check_now = driver.find_element(By.XPATH, check_path).click()
                handles_list = driver.window_handles
                for i in handles_list:
                    if i != window_before:
                        window_after = i
                driver.switch_to.window(window_after)
                remain_path = "//*[@id=\"lblOrderNum3\"]"
                remain_now = driver.find_element(By.XPATH, remain_path).text

                if int(remain_now) < int(NTU_court_dict['court_num']):
                     NTU_court_dict['court_num'] = int(remain_now)
                driver.close() # 關閉跳出視窗
                driver.switch_to.window(window_before)
        except:
            return False

    def send_information():
        # 填寫電子郵件
        email_input = driver.find_element(By.XPATH, "//*[@id=\"ctl00_ContentPlaceHolder1_txtEmail\"]")
        email_input.clear()
        email_input.send_keys(NTU_court_dict['email'])

        # 選擇付款方式
        selectpayment = Select(driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$DropLstPayMethod"))
        selectpayment.select_by_visible_text(NTU_court_dict['payment_method'])
        if NTU_court_dict['payment_method'] == "時數券": # 時數卷時輸入收據編號
            receipt_number = driver.find_element(By.XPATH, "//*[@id=\"ctl00_ContentPlaceHolder1_txtpayHourNum\"]")
            receipt_number.clear()
            receipt_number.send_keys(NTU_court_dict['receipt'])
        
        select_endtime = Select(driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$DropLstTimeEnd"))
        select_endtime.select_by_visible_text(NTU_court_dict['end_time'])

        court_number = driver.find_element(By.XPATH, "//*[@id=\"ctl00_ContentPlaceHolder1_txtPlaceNum\"]")
        court_number.clear()
        court_number.send_keys(NTU_court_dict['court_num'])

    # 驗證碼輸入，它會等待用戶手動輸入驗證碼，然後將用戶輸入的驗證碼填入網頁中的對應欄位
    def send_vaildatecode():
        testinput = input()
        vaildatecode = driver.find_element(By.XPATH, "//*[@id=\"ctl00_ContentPlaceHolder1_txtValidateCode\"]")
        vaildatecode.clear()
        vaildatecode.send_keys(testinput)  #暫時用input處理，到時候丟驗證碼變數

    # 開始執行
    # 登入
    if NTU_account_dict['ID'] == '眷屬/校友/校外':
        driver.find_element(By.XPATH, "//*[@id=\"__tab_ctl00_ContentPlaceHolder1_tcTab_TabPanel3\"]").click()
        account = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_tcTab_TabPanel3_txtMemberID"]')
        account.clear()
        account.send_keys(NTU_account_dict["account"])
        password = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_tcTab_TabPanel3_txtPassword"]') #網頁的密碼點
        password.clear()
        password.send_keys(NTU_account_dict["password"])
        driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_tcTab_TabPanel3_btnLogin"]').click()
    else:
        driver.find_element(By.XPATH, "//*[@id=\"ctl00_ContentPlaceHolder1_tcTab_tpValidator_HypLinkStu\"]/p").click()
        account = driver.find_element(By.XPATH, "//*[@id=\"myTable\"]/td/input")
        account.clear()
        account.send_keys(NTU_account_dict["account"])
        password = driver.find_element(By.XPATH, '//*[@id="myTable2"]/td/input') #網頁的密碼點
        password.clear()
        password.send_keys(NTU_account_dict["password"])
        driver.find_element(By.XPATH, "//*[@id=\"content\"]/form/table/tbody/tr[3]/td[2]/input").click()
    
    # 處理星期時間(影響後續定位)
    weekday = calendar.weekday(int(NTU_court_dict['date'][0:4]), int(NTU_court_dict['date'][5:7]), int(NTU_court_dict['date'][8:]))
    NTU_court_dict['weekday'] = str(weekday + 2)

    NTU_court_dict['timeindex'] = []
    time2index_list = ['8:00', '9:00', '10:00', '11:00', '12:00', '13:00', '14:00', '15:00',
                       '16:00', '17:00', '18:00', '19:00', '20:00', '21:00', '22:00']

    start_index = time2index_list.index(NTU_court_dict['start_time'])
    end_index = time2index_list.index(NTU_court_dict['end_time'])
    crossperiod = end_index - start_index
    if crossperiod != 1:
        for i in range(crossperiod):
            NTU_court_dict['timeindex'].append((start_index + i) + 2)
    else:
        NTU_court_dict['timeindex'].append(start_index + 2)

    while True:
        # 選擇場地日期
        select = Select(driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$tcTab$tpValidator$DropLstPlace"))
        select.select_by_visible_text(NTU_court_dict['court']) # 輸入場地
        search_data = driver.find_element(By.XPATH, "//*[@id=\"ctl00_ContentPlaceHolder1_tcTab_tpValidator_DateTextBox\"]")
        search_data.clear()
        search_data.send_keys(NTU_court_dict['date']) # 輸入日期
        driver.find_element(By.XPATH, "//*[@id=\"ctl00_ContentPlaceHolder1_tcTab_tpValidator_Button1\"]").click()

        # 選擇要預約時間(索引值先時段再星期)
        time.sleep(1)
        driver.find_element(By.XPATH, "/html/body/div/div[1]/button").click()
        # 先檢查剩餘場地數量
        check_result = check_remain_number()
        if int(NTU_court_dict['court_num']) == 0:
            return [False]
        elif check_result == False:
            return [False]

        temp_path = f"/html/body/form/table/tbody/tr[3]/td/div/table/tbody/tr/td[2]/div/div[2]/table/tbody/tr[2]/td/table[2]/tbody/tr[{str(NTU_court_dict['timeindex'][0])}]/td[{NTU_court_dict['weekday']}]//*[@id=\"btnOrder\"]"
        driver.find_element(By.XPATH, temp_path).click()

        # 輸入資料
        send_information()
        # 輸入判別的驗證碼
        send_vaildatecode()
        # 送出
        driver.find_element(By.XPATH, "//*[@id=\"ctl00_ContentPlaceHolder1_btnOrder\"]").click()

        # 送出後的提示訊息並關閉視窗
        time.sleep(2)
        alert = driver.switch_to.alert
        alert_text = alert.text

        # 驗證碼錯誤則迴圈嘗試至正確
        while alert_text == "圖像驗證碼錯誤！\n請重新輸入！":
            alert.accept()
            send_vaildatecode()
            driver.find_element(By.XPATH, "//*[@id=\"ctl00_ContentPlaceHolder1_btnOrder\"]").click()
            time.sleep(2)
            alert = driver.switch_to.alert
            alert_text = alert.text
        # 場地數量不足則重新判斷再輸入
        if alert_text == f"你預約的場地數，已超過時段:{NTU_court_dict['start_time']}至{NTU_court_dict['end_time']} 開放的場地數，請重新輸入預約的場地數":
            alert.accept()
            driver.back()
            driver.back()
            continue  # 重新回到選擇場地日期

        if alert_text == "恭喜您!場地已預約成功，本系統另寄通知信告知": # 訂成功時
            time.sleep(2)
            alert.accept()
            driver.close()
            return [True, NTU_court_dict['court_num']]
        elif alert_text == "預約資料成功!":
            time.sleep(2)
            alert.accept()
            driver.close()
            return [True, NTU_court_dict['court_num']]
        else:
            driver.close()
            return [False]


# 排程預約時間
def ntu_schedule_book():
    # 要預約的日期
    book = NTU_court_dict['date'].split('/')
    book_date = date(int(book[0]), int(book[1]), int(book[2]))

    # 開始可以預約場地的時間
    startDate = book_date + timedelta(days = -7)
    startTime = datetime(startDate.year, startDate.month, startDate.day,
                                  8, 0, 0)

    # 如果現在時間小於可以開始的時間，就等待
    while datetime.now() < startTime:
            time.sleep(1)
    # 時間到就執行程式
    return ntu_book()


# 師大球場預約後端
def ntnu_book(ntnu_dict):
    if ntnu_dict["check"] == True:
        # 連到網站
        driver.get('http://iportal.ntnu.edu.tw/ntnu/')
        # 填入帳號
        account = driver.find_element_by_xpath('//*[@id="muid"]')
        account.clear()
        account.send_keys(ntnu_dict["帳號"])
        # 填入密碼
        password = driver.find_element_by_xpath('//*[@id="mpassword"]')
        password.clear()
        password.send_keys(ntnu_dict["密碼"])
        # 按下確認
        driver.find_element_by_xpath('/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/table/tbody/tr[2]/td/table[1]/tbody/tr[2]/td/div/input').click()
        # 點選 學務相關系統
        driver.find_element_by_xpath('//*[@id="divStandaptree"]/ul/li[3]/a').click()
        # 點選 學生場地借用系統
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ap-affair"]/ul/li[11]/a'))).click()
        # 換到新分頁
        driver.switch_to.window(driver.window_handles[1])
        # 點選社團幹部登入
        driver.find_element_by_xpath('//*[@id="sidebar1"]/li[4]/a').click()
        # 點擊 登入
        driver.find_element_by_xpath('//*[@id="ClubSpace_CadreLoginList01"]/table/tbody/tr/td[6]/button').click()
        WebDriverWait(driver, 10).until(EC.alert_is_present())
        driver.switch_to.alert.accept()

        # 點擊 叉叉
        driver.find_element_by_xpath('/html/body/div/div/div/a').click()      
        # 點擊 社團活動場地登記
        driver.find_element_by_xpath('//*[@id="sidebar1"]/li[7]/a/span').click()
        # 點擊 新增
        driver.find_element_by_xpath('//*[@id="button_1I2A"]/a').click()
        
        # 點擊 社團活動        
        driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/form/div[1]/div[2]/select').click()
        # 選擇 社團活動        
        driver.find_element(By.XPATH, f'//*[text()="{ntnu_dict["社團活動"]}"]').click()

        # 點擊 場地類別名稱
        driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/form/div[2]/div[2]/select').click()
        # 選擇 場地類別名稱
        driver.find_element(By.XPATH, f'//*[text()="{ntnu_dict["場地類別名稱"]}"]').click()
        
        # 點擊 場地名稱
        driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/form/div[3]/div[2]/select').click()
        # 選擇 場地名稱                
        driver.find_element(By.XPATH, f'//*[text()="{ntnu_dict["場地名稱"]}"]').click()
        
        # 填入 預約日期
        ntnu_date = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/form/div[4]/div[2]/input')
        ntnu_date.clear()
        ntnu_date.send_keys(ntnu_dict["預約日期"])
        
        # 選擇 租借開始時間
        select1 = Select(driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/form/div[5]/div[2]/select'))
        select1.select_by_visible_text(ntnu_dict["開始時間"])


        # 選擇 租借結束時間       
        select2 = Select(driver.find_element_by_xpath('//*[@id="el_ClubSpace_SpaceBookingCadreInsert01_EndTimePeriod"]'))
        select2.select_by_visible_text(ntnu_dict["結束時間"])
        
        # 儲存
        #driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/span[3]/table/tbody/tr/td[2]/a').click()

# 建立借場小幫手主程式
rent = select_system_page()
rent.master.title("借場小幫手")  # 標題設定
rent.mainloop()

#!/usr/bin/env python
# coding: utf-8

# In[71]:


import pandas as pd
import time
from selenium.webdriver.common.keys import Keys
import requests
import time
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import NoSuchElementException
import os,sys
import tkinter
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import datetime
import pandas as pd 
import numpy as np
import threading

#レジのログイン画面のURL
url = 'http://10.0.1.11/Login/login'


# In[72]:


#ファイル参照
def open_file():
    fTyp = [("","*")]
    iDir = os.path.abspath(os.path.dirname('__file__'))
    filepath = filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)
    path1.set(filepath)

#フォルダ参照
def open_folder():
    iDir = os.path.abspath(os.path.dirname('__file__'))
    folderpath = filedialog.askdirectory(initialdir = iDir)
    path2.set(folderpath)


# In[69]:


#自動入力
def auto_work(user_id,password,file,folder):
    #エクセル読み込み、データフレーム化
    df = pd.read_excel(file)
    if '保証番号' and '請求額' not in list(df.columns):
        messagebox.showerror('フォーマットエラー', 'ファイルのフォーマットが違います。')
        return -1
    
    df = df.loc[:, ~df.columns.str.contains('Unnamed')]
    df1 = df[['保証番号','請求額']]
    
    #Chrome起動（chromeドライバーの読み込み）※アプリファイルとドライバーファイルは同じ場所に格納
    #chrome driverとchromeのバージョンはそろえる
    driver = webdriver.Chrome(resource_path('./driver/chromedriver.exe'))
    driver.get(url)

    #レジログイン
    #id,passwordはキーボードから入力
    login_id = driver.find_element_by_id('userid')
    login_pass = driver.find_element_by_id('password')

    login_id.send_keys(user_id)
    login_pass.send_keys(password)

    click_login_button = driver.find_element_by_id('id_login').click()

    #督促検索画面に移行
    click_demand_search = driver.find_element_by_xpath('/html/body/div[4]/div/div/div[5]/div[2]/form/button').click()
    driver.switch_to.window(driver.window_handles[1])

    #処理されなかったデータのリスト
    skip_no = []
    #データの数分ループ
    for item in df1.iterrows():
        """
        approval_no:承認番号(保証番号)
        claim_amount:請求額
        """
        approval_no = item[1][0].item()
        claim_amount = item[1][1].item()

        #承認番号検索、詳細画面へ移行
        input_approval_no = driver.find_element_by_id('approvalno')
        input_approval_no.send_keys(approval_no)
        time.sleep(0.1)
        input_approval_no.send_keys(Keys.ENTER)
        
        #承認番号検索に1件も引っかからなかった場合の例外処理
        try:
            click_demand_detail = driver.find_element_by_xpath('/html/body/div[4]/div/div/div[1]/table/tbody/tr[2]/td[3]/form/button').click()
        except NoSuchElementException:
            skip_no.append([approval_no, '検索該当なし'])
            #検索欄をクリア
            approval_no_clear = driver.find_element_by_id('approvalno').clear()
            continue

        driver.switch_to.window(driver.window_handles[2])


        #請求情報詳細
        click_clame_detail_but = driver.find_element_by_id('id_claim_detail_Demand').click()
        driver.switch_to.window(driver.window_handles[3])

        #既に変動費が入力済みの場合はスキップ
        skip = 0
        items = driver.find_elements_by_class_name('tr-default')
        for item in items:
            if 'その他変動費' in item.text:
                skip = 1
        if skip == 1:
            skip_no.append([approval_no,'変動費入力済み'])
        if skip == 0:        
            #明細行追加
            add_detail = driver.find_element_by_id('id_addRow_Tenant').click()

            #テーブルの1番下のドロップリストから"その他変動費"を選択
            table_items = driver.find_elements_by_class_name('tr-default')
            item_num = str(len(table_items)-1)
            select_item = driver.find_element_by_id('itemno'+item_num)
            select = Select(select_item)
            select.select_by_visible_text('その他変動費')

            #請求金額を入力
            input_claim_amount = driver.find_element_by_id('claimamountin'+item_num+'_text')
            input_claim_amount.send_keys(claim_amount)

            #更新
            update_claim_info = driver.find_element_by_id('id_update_TenantCharge')
            update_claim_info.click()
            Alert(driver).accept()

        #開いたタブを閉じる
        close_claim = driver.find_element_by_id('id_close_TenantCharge').click()
        Alert(driver).accept()
        driver.switch_to.window(driver.window_handles[2])
        close_demand = driver.find_element_by_id('id_close_Demand').click()
        Alert(driver).accept()

        #承認番号検索欄をクリア
        driver.switch_to.window(driver.window_handles[1])
        approval_no_clear = driver.find_element_by_id('approvalno').clear()
    
    #未処理データがある場合はそのデータをエクセルファイルとして出力、メッセージの表示
    if len(skip_no) > 0:
        skip_data = pd.DataFrame(skip_no, columns=['保証番号','備考'])
        skip_data.to_excel(folder+'/'+'error_data.xls',index=False)
        messagebox.showinfo('WARNING', '処理されていないデータがあります')
    driver.quit()


# In[70]:


root = tkinter.Tk()
# ウィンドウのタイトルを指定する
root.title("Automation App")
# ウィンドウサイズを指定する。横×縦
root.geometry("960x300")


font = "Arial"

#user id 入力欄
label1 =tkinter.Label(
    root,
    font = font,
    text = "USER ID",
)
label1.grid(
    row = 0,
    column = 0,
)
input_userId = tkinter.Entry(width=90)
input_userId.grid(row=0, column=1)

#password 入力欄
label2 =tkinter.Label(
    root,
    font = font,
    text = "PASSWORD"
)
label2.grid(
    row = 1,
    column = 0,
)
input_password = tkinter.Entry(width=90)
input_password.grid(row=1, column=1)

#Excel file 読み込み
label3 =tkinter.Label(
    root,
    font = font,
    text = "Excel File"
)
label3.grid(
    row = 2,
    column = 0,
)
path1 = StringVar()
file_path = ttk.Entry(textvariable=path1, width=90)
file_path.grid(row=2, column=1)

#未処理データのエクセルファイルを保存するフォルダの欄
label4 = tkinter.Label(
        root,
    font = font,
    text = "Folder"
)
label4.grid(
    row = 3,
    column = 0,
)

path2 = StringVar()
folder_path = ttk.Entry(textvariable=path2, width=90)
folder_path.grid(row=3, column=1)

label5 = tkinter.Label(
    root,
    font = font,
    text = "実行中..."
)
progressbar = ttk.Progressbar(
    root,
    length=400,
    mode="indeterminate",
)

#実行ボタンを押したときに実行される関数    
def click_func():
    user_id = input_userId.get()
    password = input_password.get()
    file = file_path.get()
    folder = folder_path.get()
    
    #プログレスバー表示
    label5.grid(row=6, column=2)
    progressbar.grid(row=6, column=1)
    progressbar.start()
    
    #作業実行
    if auto_work(user_id,password,file,folder) == -1:
        progressbar.stop()
        label5.grid_forget()
        progressbar.grid_forget()
        messagebox.showinfo('FileReference Tool', '中止しました')
        return
    
    progressbar.stop()
    label5.grid_forget()
    progressbar.grid_forget()
    messagebox.showinfo('FileReference Tool', '完了')    

#実行す処理を別スレッドで処理
def start_thread1():
    thread1 = threading.Thread(target=click_func)
    thread1.start()
    
    
    
#各ボタンの関数の割り当て    
button1 = tkinter.Button(
    text="参照",
    width=10,
    command = open_file)
button1.grid(row=2, column=2)

button2 = tkinter.Button(
    text="参照",
    width=10,
    command = open_folder)
button2.grid(row=3, column=2)


button2_2 = tkinter.Button(
    text="実行",
    width=10,
    command = start_thread1)
button2_2.grid(row=3, column=3)



root.mainloop()


# In[ ]:





# In[ ]:





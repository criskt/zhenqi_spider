#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2018/3/7 21:50
# @Author : ShawSha # @Site :

import tkinter
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

from time import sleep
import time
import pandas as pd
import random


class spider_zhenqi(object):

    def __init__(self):
        # 创建主窗口,用于容纳其它组件
        self.root = tkinter.Tk()
        # 给主窗口设置标题内容
        self.root.title("同比变化率爬虫")
        # 创建 TWO 个输入框,并设置尺寸
        self.label_null = tkinter.Label(self.root, text='@Spider: 真气网同比变化率数据>>>')
        self.label_st = tkinter.Label(self.root, text='Start date:')
        self.label_et = tkinter.Label(self.root, text='End date:')

        base_date = tkinter.StringVar()
        base_date.set('2018-01-01')
        today = tkinter.StringVar()
        today.set(time.strftime('%Y-%m-%d', time.localtime()))
        self.st_input = tkinter.Entry(self.root, width=20, textvariable=base_date)
        self.et_input = tkinter.Entry(self.root, width=20, textvariable=today)

        self.var = tkinter.StringVar()
        self.var.set('sf')
        self.sf = tkinter.Radiobutton(self.root, variable=self.var, text='省份', value='sf')
        self.cs = tkinter.Radiobutton(self.root, variable=self.var, text='74城市', value='cs')

        # 创建一个回显列表
        self.display_info = tkinter.Listbox(self.root, width=50)

        # 创建一个查询结果的按钮
        self.result_button = tkinter.Button(self.root, command=self.data_c, text="去爬", width=20)

        options = Options()
        options.add_argument('--headless')
        #options.add_argument('--disable-gpu')
        #options.add_argument('lang=zh_CN.UTF-8')
        #options.add_argument(
         #   'user-agent="Mozilla/5.0 (iPod; U; CPU iPhone OS 2_1 like Mac OS X; ja-jp) AppleWebKit/525.18.1 (KHTML, like Gecko) Version/3.1.1 Mobile/5F137 Safari/525.20"')
        self.browser = webdriver.Chrome()
        #self.browser = webdriver.Chrome(chrome_options=options)

        self.wait = WebDriverWait(self.browser, 10)
        self.data_f = pd.DataFrame()

    # 完成布局
    def gui_arrang(self):
        self.label_null.grid(columnspan=2, row=0, pady=12, stick=tkinter.W)
        self.label_et.grid(column=0, row=2, stick=tkinter.E)
        self.label_st.grid(column=0, row=1, stick=tkinter.E)
        self.st_input.grid(column=1, row=1, stick=tkinter.W)
        self.et_input.grid(column=1, row=2, stick=tkinter.W)
        self.sf.grid(column=0, row=3, stick=tkinter.E, pady=5)
        self.cs.grid(column=1, row=3, pady=5)
        self.display_info.grid(row=4, columnspan=2, stick=tkinter.W+tkinter.E, padx=10, pady=10)
        self.result_button.grid(row=5, columnspan=2, pady=5)
        self.display_info.insert(tkinter.END, '已准备就绪... ')

    # 选择时间
    def select_date(self):

        self.startTime = self.st_input.get()
        self.endTime = self.et_input.get()
        try:
            date1 = self.browser.find_element_by_xpath('//*[@id="startTime1date"]')
            date1.clear()
            date1.send_keys(self.startTime)

            date2 = self.browser.find_element_by_xpath('//*[@id="endTime1date"]')
            date2.clear()
            date2.send_keys(self.endTime)

            self.browser.find_element_by_xpath('/html/body/main/div[2]/button').click()

        except TimeoutException:
            return self.select_date()

    def select_arr(self):
        area = 19
        if self.var.get() == 'cs':
            area = 13
        try:
            submitsf = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'body > main > label:nth-child(' + str(area) + ')'))
            )
            submitsf.click()

        except TimeoutException:
            return self.select_arr()

    def get_columns(self):
        namet1 = []
        namet2 = []
        for x in range(1, 10):
            cc = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                      'body > main > div.bootstrap-table > div.fixed-table-container > div.fixed-table-header > table > thead > tr:nth-child(1) > th:nth-child(' + str(x) + ')')))
            namet1.append(cc.text)
        for y in range(1, 15):
            xj = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                     'body > main > div.bootstrap-table > div.fixed-table-container > div.fixed-table-header > table > thead > tr:nth-child(2) > th:nth-child(' + str(y) + ')')))
            namet2.append(xj.text)
        columns = [namet1[0], namet1[1], namet1[2] + namet2[0], namet1[2] + namet2[1], namet1[3] + namet2[2],
                   namet1[3] + namet2[3],
                   namet1[4] + namet2[4], namet1[4] + namet2[5], namet1[5] + namet2[6], namet1[5] + namet2[7],
                   namet1[6] + namet2[8],
                   namet1[6] + namet2[9], namet1[7] + namet2[10], namet1[7] + namet2[11], namet1[8] + namet2[12],
                   namet1[8] + namet2[13]]
        #print('Two : Gain Columns Successful !!')
        return columns

    def get_data(self):
        line = 32
        if self.var.get() == 'cs':
            line = 74

        data_zip = []
        down_url = 'http://palm.zq12369.com/new2/resource/img/arrow-down.png'
        for i in range(1, line):
            temp = []
            for j in range(1, 17):
                data_src = self.browser.find_element_by_xpath(
                    '//*[@id="pointdata"]/tbody/tr[' + str(i) + ']/td[' + str(j) + ']')
                data = data_src.text
                for img in range(4, 17, 2):
                    if j == img:
                        data = float(data.strip('%')) / 100
                        data = float('%.4f' % data)
                        imgl = self.browser.find_element_by_xpath(
                            '//*[@id="pointdata"]/tbody/tr[' + str(i) + ']/td[' + str(j) + ']/img')
                        imgurl = imgl.get_attribute('src')
                        if imgurl == down_url:
                            data = 0 - data
                temp.append(data)
            data_zip.append(temp)
        return data_zip

    def data_c(self):
        self.login()
        self.display_info.insert(tkinter.END, 'Login successful ...')
        sleep(2)
        self.select_date()
        self.display_info.insert(tkinter.END, 'Date successful ...')
        sleep(3)
        self.select_arr()
        self.display_info.insert(tkinter.END, 'Area successful ...')
        sleep(4)
        columns = self.get_columns()
        print(columns)
        self.display_info.insert(tkinter.END, 'Columns successful ...')
        sleep(2)
        data = self.get_data()
        print(data)
        self.display_info.insert(tkinter.END, 'Gain data successful ...')
        data_f = pd.DataFrame(data, columns=columns)

        self.ran = random.sample(
            ['z', 'y', 'x', 'w', 'v', 'u', 't', 's', 'r', 'q', 'p', 'o', 'n', 'm', 'l', 'k', 'j', 'i', 'h', 'g', 'f',
             'e', 'd', 'c', 'b', 'a'], 3)
        file_name = 'file_'+''.join(self.ran) + '.xlsx'
        writer = pd.ExcelWriter(file_name)
        data_f.to_excel(writer, 'sheet1', index=False)
        self.display_info.insert(tkinter.END, 'Save data to excel successful ...')
        print(self.data_f)


    def login(self):
        try:
            self.browser.get('http://palm.zq12369.com/new2/login.php')
            inputun = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#username'))
            )
            inputpd = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#password'))
            )
            submit = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '#login-button'))
            )
            inputun.send_keys('*****')
            inputpd.send_keys('********')
            submit.click()

            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#sidebar > ul > li:nth-child(5) > a > span'))
            )

            self.browser.get('http://palm.zq12369.com/new2/module/ratio.php')
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'body > main > button.btn.btn-info.btn-zdy'))
            )

            self.browser.find_element_by_xpath('/html/body/main/button[1]').click()

            print(' Login Successful !!')
        except TimeoutException:
            return self.login()


def main():

    spider = spider_zhenqi()
    spider.gui_arrang()
    tkinter.mainloop()



if __name__ == "__main__":
    main()

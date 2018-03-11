#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2018/3/7 21:50
# @Author : ShawSha # @Site :

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from time import sleep
import pandas as pd

options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
options.add_argument('--windows-size=1920,1080')
browser = webdriver.Chrome(chrome_options=options)

wait = WebDriverWait(browser, 10)


def login():
    try:
        browser.get('http://palm.zq12369.com/new2/login.php')
        inputun = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '#username'))
        )
        inputpd = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '#password'))
        )
        submit = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '#login-button'))
        )
        inputun.send_keys('zzzjxz')
        inputpd.send_keys('zhenqi123')
        submit.click()

        wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '#sidebar > ul > li:nth-child(5) > a > span'))
        )

        browser.get('http://palm.zq12369.com/new2/module/ratio.php')
        wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'body > main > button.btn.btn-info.btn-zdy'))
        )

        browser.find_element_by_xpath('/html/body/main/button[1]').click()

        print('One: Login Successful !!')
    except TimeoutException:
        return login()


def select_date(startTime, endTime):
    try:
        date1 = browser.find_element_by_xpath('//*[@id="startTime1date"]')
        date1.clear()
        date1.send_keys(startTime)

        date2 = browser.find_element_by_xpath('//*[@id="endTime1date"]')
        date2.clear()
        date2.send_keys(endTime)

        browser.find_element_by_xpath('/html/body/main/div[2]/button').click()

        print('Two: Select date successful !!')
    except TimeoutException:
        return select_date(startTime, endTime)


def select_arr(area):
    try:
        submitsf = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'body > main > label:nth-child(' + str(area) + ')'))
            # body > main > label:nth-child(15) 338城市
            # body > main > label:nth-child(17) 2+26城市
            # body > main > label:nth-child(13) 74城市
            # body > main > label:nth-child(19) 省份
        )
        submitsf.click()

        print('Three: Select area successful !!')
    except TimeoutException:
        return select_arr(area)


def get_columns():
    namet1 = []
    namet2 = []
    for x in range(1, 10):
        cc = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                                                        'body > main > div.bootstrap-table > div.fixed-table-container > div.fixed-table-header > table > thead > tr:nth-child(1) > th:nth-child(' + str(
                                                            x) + ')')))
        namet1.append(cc.text)
    for y in range(1, 15):
        xj = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                                                        'body > main > div.bootstrap-table > div.fixed-table-container > div.fixed-table-header > table > thead > tr:nth-child(2) > th:nth-child(' + str(
                                                            y) + ')')))
        namet2.append(xj.text)
    columns = [namet1[0], namet1[1], namet1[2] + namet2[0], namet1[2] + namet2[1], namet1[3] + namet2[2],
               namet1[3] + namet2[3], namet1[4] + namet2[4], namet1[4] + namet2[5], namet1[5] + namet2[6],
               namet1[5] + namet2[7],
               namet1[6] + namet2[8], namet1[6] + namet2[9], namet1[7] + namet2[10], namet1[7] + namet2[11],
               namet1[8] + namet2[12],
               namet1[8] + namet2[13]]
    print('Four: Gain Columns Successful !!')
    return columns


def get_data(line):
    data_zip = []
    down_url = 'http://palm.zq12369.com/new2/resource/img/arrow-down.png'
    for i in range(1, line):
        temp = []
        for j in range(1, 17):
            data_src = browser.find_element_by_xpath('//*[@id="pointdata"]/tbody/tr[' + str(i) + ']/td[' + str(j) + ']')
            data = data_src.text
            for img in range(4, 17, 2):
                if j == img:
                    data = float(data.strip('%')) / 100
                    data = float('%.4f' % data)
                    imgl = browser.find_element_by_xpath(
                        '//*[@id="pointdata"]/tbody/tr[' + str(i) + ']/td[' + str(j) + ']/img')
                    imgurl = imgl.get_attribute('src')
                    if imgurl == down_url:
                        data = 0 - data
            temp.append(data)
        data_zip.append(temp)
    print('Five: Gain data successful !!')
    return data_zip


def main():
    # body > main > label:nth-child(15) 338城市 area= 15 line= 335
    # body > main > label:nth-child(17) 2+26城市 area= 17 line= 29
    # body > main > label:nth-child(13) 74城市 area= 13 line= 75
    # body > main > label:nth-child(19) 省份 area= 19 line= 32

    startTime = '2018-01-01'
    endTime = '2018-03-08'

    area = 17
    line = 29

    np = 0
    if area == 17:
        np = '2+26城市'
    if area == 15:
        np = '338城市'
    if area == 13:
        np = '74城市'
    if area == 19:
        np = '省份'

    login()
    sleep(2)
    select_date(startTime, endTime)
    select_arr(area)
    columns = get_columns()
    data = get_data(line)
    filename = 'file_' + np + '_' + startTime + '_' + endTime + '.xlsx'
    sheetname = startTime + '_' + endTime
    writer = pd.ExcelWriter(filename)
    data_p = pd.DataFrame(data, columns=columns)
    data_p.to_excel(writer, sheet_name=sheetname, index=False)
    writer.save()

    browser.close()


if __name__ == '__main__':
    main()

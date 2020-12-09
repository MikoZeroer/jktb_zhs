from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
import os
import xlrd


def dk_main(username, password, qsh):
    driver = webdriver.Chrome()
    while 1:
        try:
            driver.get("https://ehall.jlu.edu.cn/infoplus/form/BKSMRDK/start")
            driver.find_element_by_name("username").send_keys(username)
            time.sleep(3)
            break
        except:
            time.sleep(3)
    driver.find_element_by_name("password").send_keys(password)
    driver.find_element_by_name("login_submit").click()
    time.sleep(3)
    while 1:
        try:
            driver.find_element_by_xpath("//span[@class='select2-selection select2-selection--single']")
            break
        except:
            driver.refresh()
            time.sleep(3)
    if driver.find_element_by_xpath("//span[@id='select2-V1_CTRL4-container']").get_attribute("title") != "2018":
        driver.find_element_by_xpath("//span[@class='select2-selection select2-selection--single']").click()
        time.sleep(1)
        driver.find_element_by_xpath("//span[text()='2018']/..").click()
        time.sleep(1)
    if driver.find_element_by_xpath("//span[@id='select2-V1_CTRL5-container']").get_attribute("title") != "191831":
        driver.find_element_by_xpath("//span[@aria-labelledby='select2-V1_CTRL5-container']").click()
        time.sleep(1)
        driver.find_element_by_xpath("//span[text()='191831']/..").click()
        time.sleep(1)
    driver.find_element_by_xpath("//select[@name='fieldSQxq']").click()
    time.sleep(1)
    driver.find_element_by_xpath("//option[text()='中心校区']").click()
    time.sleep(1)
    driver.find_element_by_xpath("//select[@name='fieldSQgyl']").click()
    time.sleep(1)
    driver.find_element_by_xpath("//option[text()='北苑1公寓']").click()
    time.sleep(1)
    qsh_box = driver.find_element_by_xpath("//input[@placeholder='例如101、102 请填写数字']")
    for i in range(0, 7):
        qsh_box.send_keys(Keys.BACK_SPACE)
    qsh_box.send_keys(qsh)
    driver.find_element_by_xpath("//input[@name='fieldZtw' and @value='1']").click()
    driver.find_element_by_xpath("//li[@class='command_li color0']/a").click()
    time.sleep(2)
    driver.find_element_by_xpath("//button[@class='dialog_button default fr']").click()
    time.sleep(2)
    if driver.find_element_by_xpath("//div[text()='办理成功!']"):
        s = "已打卡"
    else:
        s = "打卡失败"
    time.sleep(1)
    driver.find_element_by_xpath("//button[@class='dialog_button default fr']").click()
    time.sleep(1)
    driver.close()
    print(s, username)


def fuc_dk(path):
    readbook = xlrd.open_workbook(path)
    sheet = readbook.sheet_by_name('Sheet1')
    i = sheet.nrows - 1
    while i > 0:
        dk_main(sheet.cell_value(i, 0), sheet.cell_value(i, 1), str(int(sheet.cell_value(i, 2))))
        i -= 1

from selenium import webdriver
import time
import re
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
import datetime
import os
import xlrd


def zhs_login(driver, username, password):
    while 1:
        try:
            driver.get(
                "https://studyh5.zhihuishu.com/videoStudy.html#/studyVideo?recruitAndCourseId"
                "=495d585b42524159454a5f585c")
            time.sleep(1)
            driver.find_element_by_xpath("//a[text()='学号']").click()
            break
        except:
            time.sleep(5)
    driver.find_element_by_xpath(
        "//input[@onclick='userindex.selectSchoolByName();' and @onkeyup='userindex.selectSchoolByName();']").send_keys(
        "吉林大学")
    xx = driver.find_element_by_xpath("//input[@id='clSchoolId']")
    driver.execute_script("arguments[0].value = '252';", xx)
    driver.find_element_by_xpath("//input[@placeholder='大学学号']").send_keys(username)
    driver.find_element_by_xpath("//input[@placeholder='密码']").send_keys(password)
    driver.find_element_by_xpath("//span[@class='wall-sub-btn']").click()
    time.sleep(1)


def zhs_closetest(driver):
    try:
        driver.find_element_by_xpath("//span[text()='A.']").click()
        time.sleep(1)
        driver.find_element_by_xpath(
            "//span[text()='" + re.findall(r'正确答案：(.)', driver.find_element_by_xpath("//p[@class='answer']").text)[
                0] + ".']").click()
        driver.find_element_by_xpath("//div[text()='关闭']").click()
    except:
        time.sleep(1)


def zhs_over(driver):
    try:
        driver.find_element_by_xpath(
            "//div[text()='特殊说明：此次提醒仅告知您可能会超长学习，并不代表您已完成了当日的学习时间，请自觉根据习"
            "惯分规则，计算自己的有效观看时间，或第二天去成绩分析中查看学习是否有效。']")
        return 1
    except:
        return 0


def zhs_next(driver):
    try:
        driver.find_element_by_xpath("//span[@class='progress-num']")
    except:
        zhs_closetest(driver)
        try:
            driver.find_element_by_xpath(
                "//li[@class='clearfix video current_play']/../following-sibling::div[1]").click()
        except:
            try:
                driver.find_element_by_xpath(
                "//li[@class='clearfix video current_play']/../../following-sibling::ul[1]/div[1]").click()
            except:
                time.sleep(1)


def zhs_goon(driver):
    try:
        ActionChains(driver).move_to_element(driver.find_element_by_xpath("//div[@class='videoArea']")).perform()
        time.sleep(1)
        driver.find_element_by_xpath("//div[@class='playButton']").click()
    except:
        time.sleep(1)


def zhs_main(path):
    readbook = xlrd.open_workbook(path)
    sheet = readbook.sheet_by_name('Sheet1')
    row_i = sheet.nrows - 1
    while row_i > 0:
        username = sheet.cell_value(row_i, 0)
        password = sheet.cell_value(row_i, 1)
        print("zhs_" + str(username) + "_start:", datetime.datetime.now())
        time_start = time.time()
        chrome_options = Options()
        chrome_options.add_argument("--mute-audio")
        driver = webdriver.Chrome(options=chrome_options)
        zhs_login(driver, username, password)
        time.sleep(1)
        driver.find_element_by_xpath("//a[text()='暂不修改']").click()
        i = 0
        while i < 3:
            try:
                driver.find_element_by_xpath("//div[@aria-label='智慧树警告']/div[3]").click()
            except:
                time.sleep(1)
            try:
                driver.find_element_by_xpath("//i[@class='iconfont iconguanbi']").click()
            except:
                time.sleep(1)
            i += 1
            time.sleep(2)
        zhs_closetest(driver)
        zhs_next(driver)
        zhs_goon(driver)
        try:
            driver.find_element_by_xpath("//div[@aria-label='智慧树警告']/div[3]").click()
        except:
            time.sleep(1)
        try:
            driver.find_element_by_xpath("//i[@class='iconfont iconguanbi']").click()
        except:
            time.sleep(1)
        while 1:
            zhs_closetest(driver)
            zhs_next(driver)
            zhs_goon(driver)
            if zhs_over(driver):
                time_end = time.time()
                dec = time_end - time_start
                minute = int(dec / 60)
                second = dec % 60
                print(str(username) + "总用时：%02d分%02d秒" % (minute, second))
                driver.close()
                break
            time.sleep(40)
        row_i += 1

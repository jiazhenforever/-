# _*_ coding: UTF-8 _*_

# Python sysnew 签到

import requests
import time
import webbrowser
from selenium import webdriver
# 键盘事件
from selenium.webdriver.support.select import Select

from bs4 import BeautifulSoup
from urllib import request
import xlwt
import xlrd
from xlutils.copy import copy
import os
import collections
import json

from soup import BeautifulSoup

class Sysnew:
    # sysnew 脚本

    _normalHeaders = {"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
                      "Accept-Encoding": "gzip, deflate", "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
                      "Content-Type": "application/x-www-form-urlencoded; charset=utf-8",
                      "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36",
                      "X-Requested-With": "XMLHttpRequest"}

    def __init__(self, username, pwd, number, url, data_path, sheetname):
        # 用户名
        self.username = username
        # 密码
        self.pwd = pwd
        # 需求编号
        self.number = number
        # 网址
        self.url = url
        # excel表格路径，需传入绝对路径
        self.data_path = data_path
        # excel表格内sheet名
        self.sheetname = sheetname

        self.driver = webdriver.Chrome()

        self.rs = requests.session()
        pass

    def getData(self):
       # 开始打开地址
       print("开始打开地址：", self.url)
       self.driver = webdriver.Chrome()
       print("打开 管理平台 页面：http://172.17.249.10/NewSys/Login.aspx")
       self.driver.get("http://172.17.249.10/NewSys/Login.aspx")
       # print(driver.page_source)
       print("输入用户名和密码")
       userName = self.driver.find_element_by_name("TextBoxUserName")
       userName.send_keys(self.username)
       pwd = self.driver.find_element_by_name("TextBoxPassword")
       pwd.send_keys(self.pwd)
       imageBtn = self.driver.find_element_by_name("ImageButtonLogin")
       imageBtn.click()
       print("")
       time.sleep(3)
       
       # 判断表格路径是否存在
       path = os.path.join(self.data_path)
       isExist = os.path.exists(path)
       if isExist:
            # 参数说明: formatting_info=True 保留原excel格式
            self.workbook = xlrd.open_workbook(self.data_path,formatting_info=False)
            self.table = self.workbook.sheet_by_name(self.sheetname)
            self.keys = self.table.row_values(0)
            self.rowNumber = self.table.nrows
            self.colNumber = self.table.ncols


            newworkbook = copy(self.workbook)
            newsheet = newworkbook.get_sheet(0)

            data = self.readExcel()
            print("读取数据data:", self.rowNumber)

            self.driver.get(self.url)
            if self.url.find('QuickId') != -1:
                # 快速版本
                print('存在快速版本')
                logBtn = self.driver.find_element_by_id("ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolderRight_lbReleaseName")
#                print("名称：", logBtn.get_attribute('textContent'))
                logBtn.click()
            
                array = []
                for i in range(1,self.rowNumber):
                    temp = self.table.row_values(i)[13]
                    # 一行值取完之后（一个字典），追加到L列表中
                    array.append(temp)
                print("临时数据：", array)

                index_index = ""
                if int(float(self.number)) in array:
                    for i in range(0, len(data)):
                        if int(float(self.number)) == data[i]["需求编号"]:
                            # print("名称：", logBtn.get_attribute('textContent'))
                            # print("版本名称：", data[i]["版本名称"])
                            # print("序号：", data[i]["序号"])
                            # newsheet.write(data[i]["序号"], 2, logBtn.get_attribute('textContent'))
                            index_index = data[i]["序号"]
                            newsheet.write(data[i]["序号"], 13, int(self.number))
                            newworkbook.save(self.data_path)
                else:
                    # newsheet.write(self.rowNumber, 2, logBtn.get_attribute('textContent'))
                    newsheet.write(self.rowNumber, 0, self.rowNumber)
                    index_index = self.rowNumber
                    newsheet.write(self.rowNumber, 13, int(self.number))
                    newworkbook.save(self.data_path)
                        
                # 关联对象
                log1Btn = self.driver.find_element_by_id("__tab_ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolderRight_TabContainer_TabPanel1")
                log1Btn.click()

                objectArray = self.driver.find_elements_by_class_name("dxtlNode")

                # print("objectArray:",objectArray)

                # for i in range(0, len(objectArray)):
                #     # print("获取的td:",val.find_elements_by_class_name("dxtl dxtl_B0"))
                #
                #     print("objectArray:", objectArray[0].find_elements_by_tag_name("td").get_attribute("innerHTML"))

                links_temp = []
                for index, val in enumerate(objectArray):
                    links_temp.append(val.find_element_by_tag_name("a"))
                    print(val.find_element_by_tag_name("a").get_attribute("innerHTML"))

                length = len(links_temp)
                print("--------：", links_temp)
                for i in range(0, length):
                    links = links_temp
                    link = links[i]
                    # print("=======：", link.get_attribute("innerHTML"))
                    # if not ("_blank" in link.get_attribute("target") or "http" in link.get_attribute("href")):

                    link.click()

                    handleb = self.driver.current_window_handle  # 获取当前页，当前页面B赋给handleb
                    sreach_windows = self.driver.window_handles  # 将当前所有handle放入列表中
                    # logBtn.get_attribute('textContent')

                    for newhandle in sreach_windows:
                        if (newhandle != handleb):
                            self.driver.switch_to.window(newhandle)

                            self.driver.refresh()


                            # h_title = driver.find_element_by_id("__tab_TabContainer_TabPanel6")
                            # print("h_title:================================",h_title.get_attribute('textContent'))

                            # s = self.driver.find_elements_by_css_selector(css_selector="ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolderRight_hlUserLog")


                            if self.is_element_exist(self.driver, "__tab_TabContainer_TabPanel6") is True:
                                if self.driver.find_element_by_id("__tab_TabContainer_TabPanel6").get_attribute('textContent') == "测试缺陷":
                                    log1Btn = self.driver.find_element_by_id("__tab_TabContainer_TabPanel6")
                                    log1Btn.click()

                                    if self.is_element_exist(self.driver, "__tab_TabContainer_TabPanel6") is True:
                                        if self.is_element_exist(self.driver, "TabContainer_TabPanel6_TestBugList_gvBugs") is True:
                                            log2Btn = self.driver.find_element_by_id("TabContainer_TabPanel6_TestBugList_gvBugs")
                                            # print("缺陷数量：", len(log2Btn.get_attribute("innerHTML")))
                                            print("缺陷：", log2Btn.get_attribute("innerHTML"))

                                            trs = log2Btn.find_element_by_tag_name('tbody').find_elements_by_tag_name(
                                               'tr')
                                            print("ui数量：", trs)

                                            res = []
                                            for i in range(1, len(trs)):
                                                # print(val.find_element_by_tag_name("a").get_attribute("innerHTML"))
                                                tr = trs[i]
                                                print("==========", tr.get_attribute("innerHTML"))
                                                res.append(tr)

                                            # print("==========::::::::", res)

                                            result = []
                                            for i in range(0, len(res)):
                                                keys = {"编号": '', "摘要": '', "负责人": '', "提出人": '', "创建时间": '', "轮次": '',
                                                       "状态": '', "严重程度": ''}
                                                print("==========::::::::", res[i].get_attribute("innerHTML"))
                                                allcols = res[i].find_elements_by_tag_name('td')
                                                for j in range(0, len(allcols)):
                                                    col = allcols[j]
                                                    print("++++++++++++++", col.text)
                                                    # 定义一个字典用来存放对应数据
                                                    dict = {}
                                                    # j对应列值
                                                    keys["编号"] = allcols[0].text
                                                    keys["摘要"] = allcols[1].text
                                                    keys["负责人"] = allcols[2].text
                                                    keys["提出人"] = allcols[3].text
                                                    keys["创建时间"] = allcols[4].text
                                                    keys["轮次"] = allcols[5].text
                                                    keys["状态"] = allcols[6].text
                                                    keys["严重程度"] = allcols[7].text
                                                    # 一行值取完之后（一个字典），追加到res列表中
                                                result.append(keys)

                                            numberStr = ""
                                            for i in range(0, len(result)):
                                                str = result[i]["负责人"] + "  1  " + result[i]["严重程度"] + "\n"
                                                # print("===================", str)
                                                # print("index_index===================", index_index)
                                                numberStr = numberStr + str

                                            newsheet.write(index_index, 19, numberStr)
                                            newworkbook.save(self.data_path)

                                            print("numberStr", numberStr)



                                        print("测数据快速版本")


                            self.driver.close()
                            # 重新定位到页面A
                            self.driver.switch_to.window(sreach_windows[0])
                    time.sleep(5)

                    # 获取计划工时和实际工时
                self.getData_temp()
                    
                        
                        
            elif self.url.find('BugzillaId') != -1:
                # 版本Bugzilla
                print('存在版本Bugzilla')
                logBtn = self.driver.find_element_by_id("ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolderRight_lblSummary")
                #                print("名称：", logBtn.get_attribute('textContent'))

                string = logBtn.get_attribute('textContent')

                array = []
                for i in range(1,self.rowNumber):
                    temp = self.table.row_values(i)[13]
                    # 一行值取完之后（一个字典），追加到L列表中
                    array.append(temp)
                print("临时数据：", array)

                index_index = ""
                if int(float(self.number)) in array:
                    for i in range(0, len(data)):
                        if int(float(self.number)) == data[i]["需求编号"]:
                            # print("名称：", logBtn.get_attribute('textContent'))
                            # print("版本名称：", data[i]["版本名称"])
                            # print("序号：", data[i]["序号"])
                            # newsheet.write(data[i]["序号"], 2, logBtn.get_attribute('textContent'))
                            index_index = data[i]["序号"]
                            newsheet.write(data[i]["序号"], 13, int(self.number))
                            newworkbook.save(self.data_path)
                else:
                    # newsheet.write(self.rowNumber, 2, logBtn.get_attribute('textContent'))
                    newsheet.write(self.rowNumber, 0, self.rowNumber)
                    index_index = self.rowNumber
                    newsheet.write(self.rowNumber, 13, int(self.number))
                    newworkbook.save(self.data_path)

                # 关联对象
                log1Btn = self.driver.find_element_by_id("__tab_ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolderRight_TabContainer_TabPanel6")
                log1Btn.click()

                objectArray = self.driver.find_elements_by_class_name("dxtlNode")

                links_temp = []
                for index, val in enumerate(objectArray):
                    links_temp.append(val.find_element_by_tag_name("a"))
                    print(val.find_element_by_tag_name("a").get_attribute("innerHTML"))

                length = len(links_temp)
                for i in range(0, length):
                    links = links_temp
                    link = links[i]
                    # print("=======：", link.get_attribute("innerHTML"))
                    # if not ("_blank" in link.get_attribute("target") or "http" in link.get_attribute("href")):

                    link.click()

                    handleb = self.driver.current_window_handle  # 获取当前页，当前页面B赋给handleb
                    sreach_windows = self.driver.window_handles  # 将当前所有handle放入列表中
                    # logBtn.get_attribute('textContent')

                    for newhandle in sreach_windows:
                        if (newhandle != handleb):
                            self.driver.switch_to.window(newhandle)

                            if self.is_element_exist(self.driver, "__tab_TabContainer_TabPanel6") is True:
                                if self.driver.find_element_by_id("__tab_TabContainer_TabPanel6").get_attribute('textContent') == "测试缺陷":
                                    log1Btn = self.driver.find_element_by_id("__tab_TabContainer_TabPanel6")
                                    log1Btn.click()

                                    if self.is_element_exist(self.driver, "TabContainer_TabPanel6_TestBugList_gvBugs") is True:
                                        log2Btn = self.driver.find_element_by_id("TabContainer_TabPanel6_TestBugList_gvBugs")
                                        # print("缺陷数量：", len(log2Btn.get_attribute("innerHTML")))
                                        print("缺陷：", log2Btn.get_attribute("innerHTML"))

                                        trs = log2Btn.find_element_by_tag_name('tbody').find_elements_by_tag_name('tr')
                                        print("ui数量：", trs)

                                        res = []
                                        for i in range(1, len(trs)):
                                            # print(val.find_element_by_tag_name("a").get_attribute("innerHTML"))
                                            tr = trs[i]
                                            print("==========", tr.get_attribute("innerHTML"))
                                            res.append(tr)

                                        # print("==========::::::::", res)

                                        result = []
                                        for i in range(0, len(res)):
                                            keys = {"编号": '', "摘要": '', "负责人": '', "提出人": '', "创建时间": '', "轮次": '',
                                                    "状态": '', "严重程度": ''}
                                            print("==========::::::::", res[i].get_attribute("innerHTML"))
                                            allcols = res[i].find_elements_by_tag_name('td')
                                            for j in range(0, len(allcols)):
                                                col = allcols[j]
                                                print("++++++++++++++", col.text)
                                                # 定义一个字典用来存放对应数据
                                                dict = {}
                                                # j对应列值
                                                keys["编号"] = allcols[0].text
                                                keys["摘要"] = allcols[1].text
                                                keys["负责人"] = allcols[2].text
                                                keys["提出人"] = allcols[3].text
                                                keys["创建时间"] = allcols[4].text
                                                keys["轮次"] = allcols[5].text
                                                keys["状态"] = allcols[6].text
                                                keys["严重程度"] = allcols[7].text
                                                # 一行值取完之后（一个字典），追加到res列表中
                                            result.append(keys)

                                        numberStr = ""
                                        for i in range(0, len(result)):
                                            str = result[i]["负责人"] + " " + result[i]["严重程度"] + "\n"
                                            # print("===================", str)
                                            # print("index_index===================", index_index)
                                            numberStr = numberStr + str

                                        newsheet.write(index_index, 19, numberStr)
                                        newworkbook.save(self.data_path)

                                        print("numberStr", numberStr)

                                    print("测数据版本Bugzilla")

                            self.driver.close()
                            # 重新定位到页面A
                            self.driver.switch_to.window(sreach_windows[0])

                    time.sleep(5)

                    # 获取计划工时和实际工时
                self.getData_temp()
        
        
            elif self.url.find('ProjectId') != -1:
                # 快速发布申请详情
                print('存在快速发布申请详情')
                logBtn = self.driver.find_element_by_id("ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolderRight_lbReleaseName")
                #                print("名称：", logBtn.get_attribute('textContent'))
                
                
                array = []
                for i in range(1,self.rowNumber):
                    temp = self.table.row_values(i)[13]
                    # 一行值取完之后（一个字典），追加到L列表中
                    array.append(temp)
                print("临时数据：", array)

                index_index = ""
                
                if int(float(self.number)) in array:
                    for i in range(0, len(data)):
                        if int(float(self.number)) == data[i]["需求编号"]:
                            # print("名称：", logBtn.get_attribute('textContent'))
                            # print("版本名称：", data[i]["版本名称"])
                            # print("序号：", data[i]["序号"])
                            # newsheet.write(data[i]["序号"], 2, logBtn.get_attribute('textContent'))
                            index_index = data[i]["序号"]
                            newsheet.write(data[i]["序号"], 13, int(self.number))
                            newworkbook.save(self.data_path)
                else:
                    # newsheet.write(self.rowNumber, 2, logBtn.get_attribute('textContent'))
                    newsheet.write(self.rowNumber, 0, self.rowNumber)
                    index_index = self.rowNumber
                    newsheet.write(self.rowNumber, 13, int(self.number))
                    newworkbook.save(self.data_path)

                # 关联对象
                log1Btn = self.driver.find_element_by_id("__tab_ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolderRight_TabContainer_TabPanel1")
                log1Btn.click()

                objectArray = self.driver.find_elements_by_class_name("dxtlNode")

                links_temp = []
                for index, val in enumerate(objectArray):
                    links_temp.append(val.find_element_by_tag_name("a"))
                    print(val.find_element_by_tag_name("a").get_attribute("innerHTML"))

                length = len(links_temp)
                for i in range(0, length):
                    links = links_temp
                    link = links[i]
                    # print("=======：", link.get_attribute("innerHTML"))
                    # if not ("_blank" in link.get_attribute("target") or "http" in link.get_attribute("href")):

                    link.click()

                    handleb = self.driver.current_window_handle  # 获取当前页，当前页面B赋给handleb
                    sreach_windows = self.driver.window_handles  # 将当前所有handle放入列表中
                    # logBtn.get_attribute('textContent')

                    for newhandle in sreach_windows:
                        if (newhandle != handleb):
                            self.driver.switch_to.window(newhandle)

                            self.driver.refresh()

                            if self.is_element_exist(self.driver, "__tab_TabContainer_TabPanel6") is True:
                                if self.driver.find_element_by_id("__tab_TabContainer_TabPanel6").get_attribute('textContent') == "测试缺陷":
                                    if self.is_element_exist(self.driver, "__tab_TabContainer_TabPanel6") is True:
                                        log1Btn = self.driver.find_element_by_id("__tab_TabContainer_TabPanel6")
                                        log1Btn.click()
                                        if self.is_element_exist(self.driver, "TabContainer_TabPanel6_TestBugList_gvBugs") is True:
                                            log2Btn = self.driver.find_element_by_id("TabContainer_TabPanel6_TestBugList_gvBugs")
                                            # print("缺陷数量：", len(log2Btn.get_attribute("innerHTML")))
                                            print("缺陷：", log2Btn.get_attribute("innerHTML"))

                                            trs = log2Btn.find_element_by_tag_name('tbody').find_elements_by_tag_name('tr')
                                            print("ui数量：", trs)

                                            res = []
                                            for i in range(1, len(trs)):
                                                # print(val.find_element_by_tag_name("a").get_attribute("innerHTML"))
                                                tr = trs[i]
                                                print("==========", tr.get_attribute("innerHTML"))
                                                res.append(tr)

                                            # print("==========::::::::", res)

                                            result = []
                                            for i in range(0, len(res)):
                                                keys = {"编号": '', "摘要": '', "负责人": '', "提出人": '', "创建时间": '', "轮次": '',
                                                    "状态": '', "严重程度": ''}
                                                print("==========::::::::", res[i].get_attribute("innerHTML"))
                                                allcols = res[i].find_elements_by_tag_name('td')
                                                for j in range(0, len(allcols)):

                                                    col = allcols[j]
                                                    print("++++++++++++++", col.text)
                                                    # 定义一个字典用来存放对应数据
                                                    dict = {}
                                                    # j对应列值
                                                    keys["编号"] = allcols[0].text
                                                    keys["摘要"] = allcols[1].text
                                                    keys["负责人"] = allcols[2].text
                                                    keys["提出人"] = allcols[3].text
                                                    keys["创建时间"] = allcols[4].text
                                                    keys["轮次"] = allcols[5].text
                                                    keys["状态"] = allcols[6].text
                                                    keys["严重程度"] = allcols[7].text
                                                    # 一行值取完之后（一个字典），追加到res列表中
                                                result.append(keys)

                                            numberStr = ""
                                            for i in range(0, len(result)):
                                                str = result[i]["负责人"] + " " + result[i]["严重程度"] + "\n"
                                                # print("===================", str)
                                                # print("index_index===================", index_index)
                                                numberStr = numberStr + str

                                            newsheet.write(index_index, 19, numberStr)
                                            newworkbook.save(self.data_path)

                                            print("numberStr", numberStr)


                                        print("测数据快速发布申请详情")

                            self.driver.close()
                            # 重新定位到页面A
                            self.driver.switch_to.window(sreach_windows[0])
                    time.sleep(5)

                    # 获取计划工时和实际工时
                self.getData_temp()
                        
            else:
                print('不存在,链接需要重新规划')
       else:
           # 存入excel表格
           book = xlwt.Workbook()
           sheet1 = book.add_sheet(self.sheetname, cell_overwrite_ok=True)

           heads = ['序号', '名称', '版本名称', '已发轮次', '轮次说明', '版本负责人', '开发方案负责人',
                    '版本状态', "", '发布模块', '状态更新', '发版时间', '提测时间', '需求编号', 'OA编号',
                    '产品经理', '需求接口人', '需求名称', '开发负责人', 'bug数量', '计划工时(人月)', '实际工时(人月)',
                    '功能点', '是否涉及客户端', '涉及开发组', '最新优先级', '是否涉及外部联调', '备注']

           ii = 0
           for head in heads:
             sheet1.write(0, ii, head)
             ii += 1
                                        
           self.driver.get(self.url)

           releaseName = self.driver.find_element_by_id("ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolderRight_lbReleaseName")
           sheet1.write(1, 2, releaseName.get_attribute('textContent'))
           sheet1.write(1, 0, 1)


           # 文件保存
           book.save(self.data_path)
           print("写入完成！")

           self.driver.close()
           pass

    def readExcel(self):
        if self.rowNumber < 2:
            print("excle内数据行数小于2")
        else:
            # 列表L存放取出的数据
            L = []
            # 从第二行（数据行）开始取数据
            for i in range(1,self.rowNumber):
                # 定义一个字典用来存放对应数据
                sheet_data = {}
                # j对应列值
                for j in range(self.colNumber):
                    # 把第i行第j列的值取出赋给第j列的键值，构成字典
                    sheet_data[self.keys[j]] = self.table.row_values(i)[j]
                    # 一行值取完之后（一个字典），追加到L列表中
                L.append(sheet_data)
        return L

    # 获取计划工时和实际工时
    def getData_temp(self):
        # 开始打开地址
        print("开始打开地址：", self.url)
        self.driver = webdriver.Chrome()
        print("打开 管理平台 页面：http://172.17.249.10/NewSys/Login.aspx")
        self.driver.get("http://172.17.249.10/NewSys/Login.aspx")
        # print(driver.page_source)
        print("输入用户名和密码")
        userName = self.driver.find_element_by_name("TextBoxUserName")
        userName.send_keys(self.username)
        pwd = self.driver.find_element_by_name("TextBoxPassword")
        pwd.send_keys(self.pwd)
        imageBtn = self.driver.find_element_by_name("ImageButtonLogin")
        imageBtn.click()
        print("")
        time.sleep(3)

        # 判断表格路径是否存在
        path = os.path.join(self.data_path)
        isExist = os.path.exists(path)
        if isExist:
            # 参数说明: formatting_info=True 保留原excel格式
            self.workbook = xlrd.open_workbook(self.data_path, formatting_info=False)
            self.table = self.workbook.sheet_by_name(self.sheetname)
            self.keys = self.table.row_values(0)
            self.rowNumber = self.table.nrows
            self.colNumber = self.table.ncols

            newworkbook = copy(self.workbook)
            newsheet = newworkbook.get_sheet(0)

            data = self.readExcel()
            print("读取数据data:", data)
            URL = "http://172.17.249.10/NewSys/Requirement/External/ReqDetail.aspx?Id=" + self.number
            self.driver.get(URL)
            if URL.find(self.number) != -1:

                log1Btn = self.driver.find_element_by_id("__tab_ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolderRight_TabContainer_TabPanel1")
                log1Btn.click()

                # 存放需求编号数据列表
                array = []
                for i in range(1, self.rowNumber):
                    temp = self.table.row_values(i)[13]
                    # 一行值取完之后（一个字典），追加到array列表中
                    array.append(temp)
                print("临时数据：", array)

                # if self.number in array:
                #     print("读取:", self.number)
                #     for i in range(0, len(data)):
                #         if self.number == data[i]["需求编号"]:
                #             print("需求编号：", data[i]["需求编号"])
                #             newsheet.write(data[i]["序号"], 13, int(self.number))
                #             newworkbook.save(self.data_path)
                # else:
                #     newsheet.write(self.rowNumber, 13, int(self.number))
                #     newworkbook.save(self.data_path)

                # 关联系统
                log2Btn = self.driver.find_element_by_id("ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolderRight_TabContainer_TabPanel1")
                log2Btn.click()

                table = log2Btn.find_element_by_class_name("table_border_sub")

                trs = table.find_element_by_tag_name('tbody').find_elements_by_tag_name('tr')
                # print("ui数量：", trs)

                result = []
                for i in range(0, len(trs)):
                    # print("val---------", trs[i].get_attribute("innerHTML"))
                    allcols = trs[i].find_elements_by_tag_name('td')

                    numberStr = ""
                    for j in range(0, len(allcols)):
                        col = allcols[j]
                        str = col.text
                        numberStr = numberStr + str
                    result.append(numberStr)

                # print("result=======",result)

                info = ""
                for index, val in enumerate(result):
                    # print("value",val)
                    if val.find("韩永纲") != -1:
                        info = val

                index_index = ""
                if int(float(self.number)) in array:
                    for i in range(0, len(data)):
                        if int(float(self.number)) == data[i]["需求编号"]:
                            newsheet.write(data[i]["序号"], 20, info[5:15])
                            newsheet.write(data[i]["序号"], 21, info[-10:])
                            index_index = data[i]["序号"]
                            newsheet.write(data[i]["序号"], 13, int(self.number))
                            newworkbook.save(self.data_path)
                else:
                    newsheet.write(self.rowNumber, 20, info[5:15])
                    newsheet.write(self.rowNumber, 21, info[-10:])
                    index_index = self.rowNumber
                    newsheet.write(self.rowNumber, 13, int(self.number))
                    newworkbook.save(self.data_path)

                self.driver.close()

                # # 关联对象
                # log3Btn = self.driver.find_element_by_id("__tab_ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolderRight_TabContainer_TabPanel6")
                # log3Btn.click()
                #
                #
                # log4Btn = self.driver.find_element_by_id("ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolderRight_TabContainer_TabPanel6_ReqRelation_gvRelation")
                #
                # objectArray = log4Btn.find_elements_by_tag_name("a")
                #
                # for i in range(0, len(objectArray)):
                #     link = objectArray[i]
                #     print("=======：", link.get_attribute("innerHTML"))
                #     # if not ("_blank" in link.get_attribute("target") or "http" in link.get_attribute("href")):
                #
                #     link.click()
                #
                #
                #     handleb = self.driver.current_window_handle  # 获取当前页，当前页面B赋给handleb
                #     sreach_windows = self.driver.window_handles  # 将当前所有handle放入列表中
                #     for newhandle in sreach_windows:
                #         if (newhandle != handleb):
                #             self.driver.switch_to.window(newhandle)
                #
                #             self.driver.refresh()
                #
                #             # 关联对象
                #             log5Btn = self.driver.find_element_by_id("__tab_ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolderRight_TabContainer_TabPanel1")
                #             log5Btn.click()
                #
                #             objectArr = self.driver.find_elements_by_class_name("dxtlNode")
                #
                #             links_temp_temp = []
                #             for index, val in enumerate(objectArr):
                #                 links_temp_temp.append(val.find_element_by_tag_name("a"))
                #                 print(val.find_element_by_tag_name("a").get_attribute("innerHTML"))
                #
                #             length = len(links_temp_temp)
                #             for i in range(0, length):
                #                 links = links_temp_temp
                #                 link = links[i]
                #                 print("=======：", link.get_attribute("innerHTML"))
                #                 str = link.get_attribute("innerHTML")
                #                 print("------------------------",link.get_attribute("href"))
                #                 # if not ("_blank" in link.get_attribute("target") or "http" in link.get_attribute("href")):
                #             #     if str.find("系统功能测试") != -1:
                #             #         link.click()
                #             #         print("------------------------")
                #             #
                #             # time.sleep(10)
                #             # self.driver.switch_to.window(sreach_windows[1])
                #
                #
                #
                #             self.driver.close()
                #             # 重新定位到页面A
                #             self.driver.switch_to.window(sreach_windows[0])
                #     time.sleep(10)

            else:
                print('不存在,链接需要重新规划')
        else:
             # 存入excel表格
            book = xlwt.Workbook()
            sheet1 = book.add_sheet(self.sheetname, cell_overwrite_ok=True)

            heads = ['序号', '名称', '版本名称', '已发轮次', '轮次说明', '版本负责人', '开发方案负责人',
                           '版本状态', "", '发布模块', '状态更新', '发版时间', '提测时间', '需求编号', 'OA编号',
                           '产品经理', '需求接口人', '需求名称', '开发负责人', 'bug数量', '计划工时(人月)', '实际工时(人月)',
                           '功能点', '是否涉及客户端', '涉及开发组', '最新优先级', '是否涉及外部联调', '备注']

            ii = 0
            for head in heads:
                sheet1.write(0, ii, head)
                ii += 1
            #              print(head)

            self.driver.get(self.url)
            #           logBtn = driver.find_element_by_id("ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolderRight_lblPlanTestTime")
            #           sheet1.write(1, 7, logBtn.get_attribute('textContent'))

            # releaseName = self.driver.find_element_by_id("ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolderRight_lbReleaseName")
            # sheet1.write(1, 1, releaseName.get_attribute('textContent'))
            # sheet1.write(1, 0, 1)

            # 文件保存
            book.save(self.data_path)
            print("写入完成！")

            self.driver.close()
            pass

    def is_element_exist(self, driver, idname):

        try:
            s = driver.find_element_by_id(idname)
            print('2222222')
            return True
        except BaseException:
            print('1111111')
            return  False


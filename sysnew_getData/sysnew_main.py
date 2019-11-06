# _*_ coding: UTF-8 _*_

# Python sysnew 签到


from sysnew_network import Sysnew

import schedule
import time

def data():
    # 打开
    print("打开")

    # time.sleep(10)

    # 表格路径必须保存为.xls格式，不可以.xlsx格式，表格要提前处理成一行，否则写入数据的时候会错乱

    # 文件的绝对路径
    data_path = "/Users/jiazhen/Desktop/sysnew2/账户组需求版本.xls"
    sheetname = "RN9月"
    # 参数依次为 账号 密码 需求编号 版本链接 表格路径 表格sheet
    sys = Sysnew("hlbai", "A@#938457", "13188", "http://172.17.249.10/NewSys/project/QuickRelese/QuickRelApplyDetail.aspx?QuickId=2096", data_path, sheetname)
    sys.getData()

    time.sleep(5)

    data_path1 = "/Users/jiazhen/Desktop/sysnew2/账户组需求版本.xls"
    sheetname1 = "RN9月"
    # 参数依次为 账号 密码 需求编号 版本链接 表格路径 表格sheet
    sys1 = Sysnew("hlbai", "A@#938457", "16032", "http://172.17.249.10/NewSys/IPOC/Bugzilla/BugzillaDetail.aspx?BugzillaId=32626", data_path1, sheetname1)
    sys1.getData()

    time.sleep(5)

    data_path2 = "/Users/jiazhen/Desktop/sysnew2/账户组需求版本.xls"
    sheetname2 = "RN9月"
    # 参数依次为 账号 密码 需求编号 版本链接 表格路径 表格sheet
    sys2 = Sysnew("hlbai", "A@#938457", "15961", "http://172.17.249.10/NewSys/project/QuickRelese/QuickRelApplyDetail.aspx?ProjectId=7217", data_path2, sheetname2)
    sys2.getData()

    time.sleep(5)


    # data_path2 = "/Users/jiazhen/Desktop/sysnew2/账户组需求版本.xls"
    # sheetname2 = "RN9月"
    # number = "15535"
    # url = ""
    # sys2 = Sysnew("hlbai", "A@#938457", number, url, data_path2, sheetname2)
    # sys2.getData_temp()

    # time.sleep(10)

    print("脚本运行完成")
    pass

data()



# while True:
#     schedule.run_pending()
#     time.sleep(20)

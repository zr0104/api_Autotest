# coding=utf-8
from common.operate_excel import *
import unittest
from parameterized import parameterized
from common.send_request import RunMethod
import json
from common.logger import MyLogging
import jsonpath
from common.is_instance import IsInstance
from HTMLTestRunner import HTMLTestRunner
import os
import time
from xlutils.copy import copy

lib_path1 = os.path.abspath(os.path.join(os.path.dirname(__file__), "../data"))
# lib_path2 = os.path.join(os.getcwd() + "../date")
# 输入path1和path2路径：path2多了当前文件夹
# D:\Sen\code\python\AutoTest\data
# D:\Sen\code\python\AutoTest\testcase../date
# print(lib_path1,lib_path2)
file_path = lib_path1 + "/" + "接口自动化测试.xls"  # excel的地址
sheet_name = "测试用例"
log = MyLogging().logger

#加载测试excel测试用例
def getExcelData():
    list = ExcelData(file_path, sheet_name).readExcel()
    print(list)
    return list



# UnitTest类必须继承Case类
class TestCase(unittest.TestCase):

    @parameterized.expand(getExcelData())
    def test_api(self, rowNumber, caseRowNumber, testCaseName, priority, apiName, url, method, parmsType, data,
                 checkPoint, isRun, result):
        #检查是否执行测试用例，并调RunMethod方法发送请求
        if isRun == "Y" or isRun == "y":
            log.info("【开始执行测试用例：{}】".format(testCaseName))
            headers = {"Content-Type": "application/json"}
            data = json.loads(data)  # 字典对象转换为json字符串
            c = checkPoint.split(",")
            log.info("用例设置检查点：%s" % c)
            print("用例设置检查点：%s" % c)
            log.info("请求url：%s" % url)
            log.info("请求参数：%s" % data)
            r = RunMethod()
            res = r.run_method(method, url, data, headers)
            log.info("返回结果：%s" % res)
            print(res)

            flag = None
            for i in range(0, len(c)):
                checkPoint_dict = {}
                checkPoint_dict[c[i].split('=')[0]] = c[i].split('=')[1]
                print("checkPoint_dict[c[i].split('=')[0]] = c[i].split('=')[1]")
                # jsonpath方式获取检查点对应的返回数据
                list = jsonpath.jsonpath(res, c[i].split('=')[0])
                value = list[0]
                check = checkPoint_dict[c[i].split('=')[0]]
                log.info("检查点数据{}：{},返回数据：{}".format(i + 1, check, value))
                print("检查点数据{}：{},返回数据：{}".format(i + 1, check, value))
                # 判断检查点数据是否与返回的数据一致
                flag = IsInstance().get_instance(value, check)

            if flag:
                log.info("【测试结果：通过】")
                # file2 = open("接口自动化测试.xls", "w", encoding="utf8")
                # file2.write(rowNumber + 1, 12, "Pass")
                #ExcelData(file_path, sheet_name).write(rowNumber + 1, 12, "Pass")
                read_value = xlrd.open_workbook(file_path)
                write_data = copy(read_value)
                write_data.get_sheet(sheet_name).write(rowNumber + 1, 12, "Pass")
            else:
                log.info("【测试结果：失败】")
                # file2 = open("接口自动化测试.xls", "w", encoding="utf8")
                # file2.write(rowNumber + 1, 12, "fail")

                read_value = xlrd.open_workbook(file_path)
                write_data = copy(read_value)
                write_data.get_sheet(sheet_name).write(rowNumber + 1, 12, "Fail")

            # 断言
            self.assertTrue(flag, msg="检查点数据与实际返回数据不一致")
        else:
            unittest.skip("不执行")


if __name__ == '__main__':
    # unittest.main()
    # Alt+Shift+f10 执行生成报告

    # 报告样式1
    suite = unittest.TestSuite()
    suite.addTests(unittest.TestLoader().loadTestsFromTestCase(TestCase))
    now = time.strftime('%Y-%m-%d %H_%M_%S')
    report_path = os.path.join(os.getcwd() + "/report/" + now + "report.html")
    print(report_path)
    with open(report_path, "wb") as f:
        runner = HTMLTestRunner(stream=f, title="Sen01万邦接口测试报告", description="01测试用例执行情况", verbosity=2)
        runner.run(suite)

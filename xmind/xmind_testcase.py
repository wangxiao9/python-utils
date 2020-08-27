__author__ = 'wangxiao'

import json

from xmindparser import xmind_to_dict

from utils.execl_config import ExcelConfig


class ExportXmindExcel:
    def __init__(self):
        self.name = input("请输入绝对路径:")
        self.data = xmind_to_dict(self.name)
        # 获取测试用例dict格式
        self.dict_data = self.data[0]
        # 实例化execlConfig
        self.execl = ExcelConfig()


    # 提取多少条case
    def get_cases_nums(self):
        nums = self.dict_data['topic']['topics'][0]['topics']
        return len(nums)

    # 提取cases list
    def get_cases(self):
        return self.dict_data['topic']['topics'][0]['topics']

    # 提取Test Repository Path
    def get_case_path(self, i):
        self.casesList = self.get_cases()
        index = self.casesList[i]['title']
        path = self.casesList[i]['topics'][0]['topics'][0]['title']
        tag = self.casesList[i]['topics'][1]['topics'][0]['title']
        priority = self.casesList[i]['topics'][2]['topics'][0]['title']
        summary = self.casesList[i]['topics'][3]['topics'][0]['title']
        des = self.casesList[i]['topics'][5]['topics'][0]['title']
        return index, path, tag, priority, summary, des

    # 提取action = （步骤+对应的步骤结果）
    def get_actions(self, num):
        return self.casesList[num]['topics'][4]['topics']

    # 处理action,分别提取步骤 = 结果
    def get_case_detail(self, num, list):
        step = list[num]['title']
        res = list[num]['topics'][0]['title']
        return step, res

    # 把每次循环的数据，放到字典中
    def converse_dict(self, num):
        index, path, tag, priority, summary, des = self.get_case_path(num)
        case_dict = {}
        case_dict['index'] = index
        case_dict['path'] = path
        case_dict['tag'] = tag
        case_dict['priority'] = priority
        case_dict['summary'] = summary
        case_dict['des'] = des
        case_dict['action'] = self.get_actions(num)
        return case_dict

    # 执行函数
    def run(self):
        nums = self.get_cases_nums()
        # 声明全局变量
        global number
        # 写入单元格从第二行开始
        number = 2
        for i in range(0, nums):
            cases = self.converse_dict(i)
            actions = cases['action']
            nums2 = len(actions)
            for j in range(0, nums2):
                args = self.get_case_detail(j, actions)
                self.execl.run(number, args, index=cases['index'], path=cases['path'], tag=cases['tag'], priority=cases['priority'], summary=cases['summary'], des=cases['des'])
                number +=1
        print("--------------execl已经生成---------------")

main = ExportXmindExcel()
main.run()
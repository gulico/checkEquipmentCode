import os

import pandas as pd
from PyQt5 import QtCore
from PyQt5.QtCore import pyqtSignal

class checkCode(QtCore.QThread):
    XTM_tree_FD = []
    XTM_tree_GF = []
    SBM_tree = []
    DC_type = ''
    df = pd.DataFrame()
    singlevalue = 0
    pvalue = 0
    uncheckedFilePath = ''
    # 通过类成员对象定义信号对象
    _signal_toTextEdit = pyqtSignal(str)
    _signal_toProgressBar = pyqtSignal(int)

    def __init__(self, uncheckedFilePath):
        super(checkCode, self).__init__()
        self.uncheckedFilePath = uncheckedFilePath

    def __del__(self):
        self.wait()

    def run(self):
        self.readFile(self.uncheckedFilePath)
        self.mainlogic()
        self.outputTXT()
        self._signal_toProgressBar.emit(100)


    def readFile(self, uncheckedFilePath):
        try:
            # 读取系统码树状图
            self._signal_toTextEdit.emit('读取系统码树状图')
            filePath = "系统码树状图.xlsx"
            XTM_tree = pd.read_excel(filePath, sheet_name=None)
            self.XTM_tree_FD = XTM_tree['风电树状图']
            self.XTM_tree_GF = XTM_tree['光伏树状图']

            # 读取设备编码
            self._signal_toTextEdit.emit('读取读取设备编码列表')
            filePath = "设备编码.xlsx"
            self.SBM_tree = pd.read_excel(filePath).dropna(axis=0, thresh=2)

            # 读取待检测设备编码表
            self._signal_toTextEdit.emit('读取读取设备编码列表')
            # filePath = "山东垣曲华昌光伏电站设备编码表-0506.xlsx"
            self.df = pd.read_excel(uncheckedFilePath, skiprows=3, header=None, sheet_name=None)
        except Exception as e:
            self._signal_toTextEdit.emit(str(e.args))
            self._signal_toTextEdit.emit('读取文件失败，请将excel文件放在相同文件夹下。')

    def reProgressBarvalue(self):
        self.pvalue = self.singlevalue + self.pvalue
        self._signal_toProgressBar.emit(int(self.pvalue))

    def mainlogic(self):  # 主逻辑
        self.singlevalue = 99/(8 * len(self.df))
        self.pvalue = 0
        for sheet_name, sheet in self.df.items():
            try:
                self._signal_toTextEdit.emit('**********************************************')
                self._signal_toTextEdit.emit('开始检测sheet：' + str(sheet_name))
                sheet =sheet.iloc[:, 0:15]
                sheet.columns = ['序号', '分公司', '场站名称', '场站简称', '场站类型', '设备名称', '工厂码U1', '全厂码F0', '系统码F1',
                                 '设备码F2', '产品码P1', '产品码P2', '组合', '设备层级', '上级设备编码']
                if sheet['工厂码U1'][0][4] == 'W':
                    self.DC_type = '风电'
                if sheet['工厂码U1'][0][4] == 'P':
                    self.DC_type = '光伏'
            except Exception as e:
                self._signal_toTextEdit.emit(str(e.args))
                self._signal_toTextEdit.emit('请确认待检测设备码excel文件中是否有空sheet，且每个sheet的字段为15个，没有缺少序号等必要字段，顺序必须与模板统一')
                continue
            self.duplicateCheck(sheet)
            self.reProgressBarvalue()
            self.OCheck(sheet)
            self.reProgressBarvalue()
            self.factoryCodeCheck(sheet)
            self.reProgressBarvalue()
            self.QCMCodeCheck(sheet)
            self.reProgressBarvalue()
            self.XTMCodeCheck(sheet)
            self.reProgressBarvalue()
            self.SBMCodeCheck(sheet)
            self.reProgressBarvalue()
            self.CPMCodeCheck(sheet)
            self.reProgressBarvalue()
            self.SBCJCheck(sheet)
            self.reProgressBarvalue()

    def duplicateCheck(self, sheet):  # 重码检查
        try:
            self._signal_toTextEdit.emit('################开始重码检查################')
            mark = sheet.duplicated(subset=['组合'], keep=False)
            repeatCode = sheet[mark]
            self._signal_toTextEdit.emit('重码条数：' + str(len(repeatCode)))
            for index, row in repeatCode.iterrows():
                self._signal_toTextEdit.emit('序号：' + str(row['序号']) + ' 组合：' + str(row['组合']))
            self._signal_toTextEdit.emit('################完成重码检查################')
        except Exception as e:
            self._signal_toTextEdit.emit(str(e.args))

    def OCheck(self, sheet):  # 重码检查
        try:
            self._signal_toTextEdit.emit('################开始字母欧（O）和数字0检查################')
            tmpsheet = sheet.fillna('-')
            containO = tmpsheet[tmpsheet['组合'].str.contains('O')]
            for index, row in containO.iterrows():
                self._signal_toTextEdit.emit('包含字母欧（O）的记录：')
                self._signal_toTextEdit.emit('序号：' + str(row['序号']) + ' 组合：' + str(row['组合']))
            self._signal_toTextEdit.emit('################完成字母欧（O）和数字0检查################')
        except Exception as e:
            self._signal_toTextEdit.emit(str(e.args))

    def factoryCodeCheck(self, sheet):  # 重码检查
        try:
            self._signal_toTextEdit.emit('################开始工厂码检查################')
            groupby_GCM = sheet.groupby('工厂码U1')
            if groupby_GCM.ngroups > 1:
                self._signal_toTextEdit.emit('工厂码错误，应该有且只有一个工厂码')
            self._signal_toTextEdit.emit('################完成工厂码检查################')
        except Exception as e:
            self._signal_toTextEdit.emit(str(e.args))

    def QCMCodeCheck(self, sheet):  # 重码检查
        try:
            self._signal_toTextEdit.emit('################开始全厂码检查################')
            groupby_QCM = sheet.groupby('全厂码F0')
            for name, group in groupby_QCM:
                if not (name[0] == 'T' or name[0] == 'G'):
                    self._signal_toTextEdit.emit('错误全厂码F0：' + name+' 全厂码F0首字母不为G或T，请检查全厂码F0首字母')
                if not (name[1:].isdigit()):
                    self._signal_toTextEdit.emit('错误全厂码F0：' + name+' 全厂码F0除第一位外不为数字，请调整为数字')
                if not len(name) == 3:
                    self._signal_toTextEdit.emit('错误全厂码F0：' + name+' 全厂码F0长度不为3，请调整全厂码F0长度为3')

            self._signal_toTextEdit.emit(str(groupby_QCM.size()))
            self._signal_toTextEdit.emit('全厂码个数：' + str(groupby_QCM.ngroups))
            if not (groupby_QCM.ngroups) == 1:  # 分组个数只为1
                self._signal_toTextEdit.emit('每个sheet全厂码F0应只有1种，不同全场码F0请分成不同sheet')
            self._signal_toTextEdit.emit('################完成全场码检查################')
        except Exception as e:
            self._signal_toTextEdit.emit(str(e.args))

    def XTMCodeCheck(self, sheet):  # 重码检查
        try:
            self._signal_toTextEdit.emit('################开始系统码检查################')
            groupby_XTM = sheet.groupby('系统码F1')
            for name, group in groupby_XTM:
                if not (name[0:3].isalpha()):
                    self._signal_toTextEdit.emit('错误系统码F1：' + name + ' 系统码F1前3位应全为字母')
                if not (name[3:5].isdigit()):
                    self._signal_toTextEdit.emit('错误系统码F1：' + name + ' 系统码F1后2位应全为数字')
                if not len(name) == 5:
                    self._signal_toTextEdit.emit('错误系统码F1：' + name + ' 系统码F1长度不为5，请调整系统码F1长度为5')

            # 选择系统码树状图模板
            tree_tmp = []
            if self.DC_type == '风电':
                tree_tmp = (self.XTM_tree_FD['系统码']).to_list()
            elif self.DC_type == '光伏':
                tree_tmp = (self.XTM_tree_GF['系统码']).to_list()

            # 检查系统码是否在树状图中
            for index, row in sheet.iterrows():
                if not (isinstance(row['系统码F1'], str)):  # 跳过NAN空值
                    continue
                if not (row['系统码F1'] in tree_tmp):
                    if row['系统码F1'][0:3] == 'MQA' and row['系统码F1'][3:5].isdigit():
                        continue
                    self._signal_toTextEdit.emit('系统码F1在树状图模板中不存在！ 序号：' + str(row['序号']) + ' 系统码F1：' + str(row['系统码F1']))
            self._signal_toTextEdit.emit('################完成系统码检查################')
        except Exception as e:
            self._signal_toTextEdit.emit(str(e.args))

    def SBMCodeCheck(self, sheet):  # 重码检查
        try:
            self._signal_toTextEdit.emit('################开始设备码检查################')
            # 风电设备码全部为2个字母+3位数字
            groupby_SBM = sheet.groupby('设备码F2')
            if self.DC_type == '风电':
                for name, group in groupby_SBM:
                    if not (name[0:2].isalpha()):
                        self._signal_toTextEdit.emit('错误设备码F2：' + name + ' 设备码F2前2位应全为字母')
                    if not (name[2:5].isdigit()):
                        self._signal_toTextEdit.emit('错误设备码F2：' + name + ' 设备码F2后3位应全为数字')
                    if not len(name) == 5:
                        self._signal_toTextEdit.emit('错误设备码F2：' + name + ' 设备码F2长度不为5，请调整设备码F2长度为5')
            # 所用设备码均在设备码模板中有
            tree_tmp = self.SBM_tree['设备/产品分类码'].to_list()

            # 检查设备码是否在树状图中
            for index, row in sheet.iterrows():
                if not (isinstance(row['设备码F2'], str)):  # 跳过NAN空值
                    continue
                if not (row['设备码F2'][0:2] in tree_tmp):
                    self._signal_toTextEdit.emit('设备码F2在设备编码模板中不存在！序号：' + str(row['序号']) + ' 设备码F2：' + str(row['设备码F2']))

            # 光伏逆变器、汇流箱、组串和支架等有特殊规则，按其规则仔细检查
            SBM_list = []
            for name, group in groupby_SBM:
                SBM_list.append(name)

            if self.DC_type == '光伏':
                GC_bit4 = False
                GC_bit6 = False
                for index, row in sheet.iterrows():
                    if not (isinstance(row['设备码F2'], str)):  # 跳过NAN空值
                        continue

                    name = row['设备码F2']
                    numb = row['序号']
                    ####################### 光伏组件 #######################
                    if name[0:2] == 'GC':  # 光伏组件
                        if len(name[2:]) == 4:  # 光伏组件4位设备码
                            GC_bit4 = True
                            if not ('TB' + str(name[2:4]) in SBM_list):
                                self._signal_toTextEdit.emit('序号：'+str(numb)+ '光伏组件设备码'+ str(name)+ '缺少对应逆变器'+ 'TB' + str(name[2:4]))
                        elif len(name[2:]) == 6:  # 光伏组件6位设备码
                            GC_bit6 = True
                            TB4 = 'TB' + str(name[2:6]) in SBM_list
                            TB2 = 'TB' + str(name[2:4]) in SBM_list
                            UC4 = 'UC' + str(name[2:6]) in SBM_list
                            UC2 = 'UC' + str(name[2:4]) in SBM_list

                            if TB4:  # 有逆变器4位
                                if not UC2:  # 有逆变器4位，无汇流箱2位
                                    self._signal_toTextEdit.emit('序号：' + str(numb) + '光伏组件设备码'+str(name) + '缺少对应汇流箱' + 'UC' + str(name[2:4]))
                            else:  # 没有4位逆变器
                                if UC4:  # 汇流箱4位
                                    if not TB2:  # 有汇流箱4位，无逆变器2位
                                        self._signal_toTextEdit.emit('序号：'+str(numb) + '光伏组件设备码'+str(name) + '缺少对应逆变器' + 'TB' + str(name[2:4]))
                                else:  # 没有4位逆变器，没有4位汇流箱
                                    if not (UC2 or TB2):  # 没有2位逆变器，没有2位汇流箱
                                        self._signal_toTextEdit.emit('序号：' + str(numb) + '光伏组件设备码'+str(name) + '缺少对应逆变器和汇流箱')
                                    elif UC2:
                                        self._signal_toTextEdit.emit('序号：'+str(numb) + '光伏组件设备码'+str(name) + '有汇流箱' + 'UC' + str(name[2:4]) + '缺少对应4位逆变器')
                                    elif TB2:
                                        self._signal_toTextEdit.emit('序号：'+str(numb) + '光伏组件设备码'+str(name) + '有逆变器' + 'TB' + str(name[2:4]) + '缺少对应4位汇流箱')
                        else:  # 光伏组件长度错误
                            self._signal_toTextEdit.emit('错误设备码F2：' + str(name)+'序号：' + str(numb) + '光伏组件设备码F2数字部分应该为4位或6位')
                    ####################### 逆变器 #######################
                    elif name[0:2] == 'TB' and '逆变器' in row['设备名称']:  # 逆变器
                        if len(name[2:]) == 4:
                            UC2 = 'UC' + str(name[2:4]) in SBM_list
                            if not UC2:  # 有逆变器4位，无汇流箱2位
                                self._signal_toTextEdit.emit('序号：' + str(numb) + '光伏组件设备码' + str(name) + '缺少对应汇流箱' + 'UC' + str(name[2:4]))
                        elif not (len(name[2:]) == 2):  # 逆变器长度错误
                            self._signal_toTextEdit.emit('序号：' + str(numb) + '错误设备码F2：' + str(name) + '逆变器设备码F2数字部分应该为2位或4位')
                    ####################### 汇流箱 #######################
                    elif name[0:2] == 'UC' and '汇流箱' in row['设备名称']:  # 汇流箱
                        if len(name[2:]) == 4:
                            TB2 = 'TB' + str(name[2:4]) in SBM_list
                            if not TB2:  # 有汇流箱4位，无逆变器2位
                                self._signal_toTextEdit.emit('序号：' + str(numb) + '光伏组件设备码' + str(name) + '缺少对应汇流箱' + 'TB' + str(name[2:4]))
                        elif not (len(name[2:]) == 2):  # 汇流箱长度错误
                            self._signal_toTextEdit.emit('序号：' + str(numb) + '错误设备码F2：' + str(name) + '汇流箱设备码F2数字部分应该为2位或4位')
                    ####################### 其他 #######################
                    elif not name[0:2] == 'UR':
                        if not (name[0:2].isalpha()):
                            self._signal_toTextEdit.emit('序号：' + str(numb) + '错误设备码F2：' + str(name) + '设备码F2前2位应全为字母')
                        if not (name[2:5].isdigit()):
                            self._signal_toTextEdit.emit('序号：' + str(numb) + '错误设备码F2：' + str(name) + '设备码F2后3位应全为数字')
                        if not len(name) == 5:
                            self._signal_toTextEdit.emit('序号：' + str(numb) + '错误设备码F2：' + str(name) + '设备码F2长度不为5，请调整设备码F2长度为5')
                ####################### 支架 #######################
                for index, row in sheet.iterrows():
                    if name[0:2] == 'UR':  # 支架
                        if not ((GC_bit4 and len(name[2:]) == 4) or (GC_bit6 and len(name[2:]) == 6)):
                            self._signal_toTextEdit.emit('序号：' + str(numb) + '错误设备码F2：' + str(name) + '支架编码长度必须与光伏组件统一')
            self._signal_toTextEdit.emit('################完成设备码检查################')
        except Exception as e:
            self._signal_toTextEdit.emit(str(e.args))

    def CPMCodeCheck(self, sheet):  # 重码检查
        try:
            self._signal_toTextEdit.emit('################开始产品码检查################')
            # 前两位
            tree_tmp = self.SBM_tree['设备/产品分类码'].to_list()

            # 检查设备码是否在编码模板中
            for index, row in sheet.iterrows():
                if not (isinstance(row['产品码P1'], str)):  # 跳过NAN空值
                    continue
                if not (row['产品码P1'][0:2] in tree_tmp):
                    self._signal_toTextEdit.emit('产品码P1在设备编码模板中不存在！ 序号：' + str(row['序号'])+ ' 产品码P1：'+ str(row['产品码P1']))

                if not (isinstance(row['产品码P2'], str)):  # 跳过NAN空值
                    continue
                if not (row['产品码P2'][0:2] in tree_tmp):
                    self._signal_toTextEdit.emit('产品码P2在设备编码模板中不存在！ 序号：' + str(row['序号'])+ '产品码P1：'+ str(row['产品码P2']))
            # 后2-3位
            tmpsheet = sheet.fillna('-')
            for index, row in tmpsheet.iterrows():
                numb = row['序号']
                ####################### 产品码P1 #######################
                if row['产品码P1'] == '-':  # 跳过NAN空值
                    continue
                if len(row['产品码P1'][2:]) == 2:
                    continue
                elif len(row['产品码P1'][2:]) == 3:
                    cpm2 = row['产品码P1'][0:2]  # 前两位
                    contain_cpm2_len = len(tmpsheet[tmpsheet['产品码P1'].str.contains(cpm2)])
                    if contain_cpm2_len < 100:
                        self._signal_toTextEdit.emit('序号：' + str(numb) + '错误产品码P1：' + str(row['产品码P1']) +'5级设备超过100个时，产品码序列号可以有3位数字,否则应为2位数字')
                else:
                    self._signal_toTextEdit.emit('序号：' + str(numb) + '错误产品码P1：' + str(row['产品码P1']) + '产品码P1应为2个字母+2位数字或2个字母+3位数字')
                ####################### 产品码P2 #######################
                if row['产品码P2'] == '-':  # 过NAN空值
                    continue
                if len(row['产品码P2'][2:]) == 2:
                    continue
                elif len(row['产品码P2'][2:]) == 3:
                    cpm2 = row['产品码P2'][0:2]  # 前两位
                    contain_cpm2_len = len(tmpsheet[tmpsheet['产品码P2'].str.contains(cpm2)])
                    if contain_cpm2_len < 100:
                        self._signal_toTextEdit.emit('序号：' + str(numb) + '错误产品码P2：' + str(row['产品码P2']) + '5级设备超过100个时，产品码序列号可以有3位数字,否则应为2位数字')
                else:
                    self._signal_toTextEdit.emit('序号：' + str(numb) + '错误产品码P2：' + str(row['产品码P2']) + '产品码P2应为2个字母+2位数字或2个字母+3位数字')
            self._signal_toTextEdit.emit('################完成产品码检查################')
        except Exception as e:
            self._signal_toTextEdit.emit(str(e.args))

    def SBCJCheck(self, sheet):  # 重码检查
        try:
            self._signal_toTextEdit.emit('################开始设备层级检查################')
            # 筛选查看“设备层级”列，下拉选项中有且只有1~5级
            groupby_SBCJ = sheet.groupby('设备层级')
            Sum_CJ = 0
            for name, group in groupby_SBCJ:
                Sum_CJ = Sum_CJ + name
                if not (name in [1, 2, 3, 4, 5]):
                    self._signal_toTextEdit.emit('错误的层级'+ str(name) + '设备层级有且只有1~5级')
            if not (Sum_CJ == 15):
                self._signal_toTextEdit.emit('缺少层级，设备层级有且只有1~5级')

            # 检查筛选出的编码组成部分与所勾选的设备层级是否匹配
            CJM_name_list = ['工厂码U1', '全厂码F0', '系统码F1', '设备码F2', '产品码P1']
            for index, row in sheet.iterrows():
                sbcj = row['设备层级']  # 设备层级1,2,3,4,5
                numb = row['序号']
                for cjEle in CJM_name_list[0:sbcj]:  # 遍历应该有内容的层级
                    if (row[cjEle + ''] != row[cjEle]):  # 为nan
                        self._signal_toTextEdit.emit('序号：' + str(numb) + '缺少设备层级：' + str(cjEle) + '设备码组合：' + str(row['组合']))

                for cjEle in CJM_name_list[sbcj:]:  # 遍历不应该有内容的层级
                    if not (row[cjEle] != row[cjEle]):  # 不为nan
                        self._signal_toTextEdit.emit('序号：' + str(numb) + '多余的设备层级：' + str(cjEle) + '设备码组合：' + str(row['组合']))
            self._signal_toTextEdit.emit('################完成设备层级检查################')
        except Exception as e:
            self._signal_toTextEdit.emit(str(e.args))

    def outputTXT(self):
        #output_path = '检查结果.txt'
        #with open(output_path, 'w', encoding='utf-8') as file1:
        #    print(self.textEdit.toPlainText(), file=file1)
        desktop_path = os.path.join(os.path.expanduser('~'), "Desktop/")
        full_path = desktop_path + self.uncheckedFilePath.split('/')[-1].split('.')[0] + '检查结果.txt'  # 也可以创建一个.doc的word文档
        with open(full_path, 'w', encoding='utf-8') as file1:
            print('111', file=file1)
        self._signal_toTextEdit.emit('*已生成检查结果到桌面*')

    @property
    def signal_toTextEdit(self):
        return self._signal_toTextEdit

    @property
    def signal_toProgressBar(self):
        return self._signal_toProgressBar



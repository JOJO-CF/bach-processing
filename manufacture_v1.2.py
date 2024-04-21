# coding=gbk
from inspect import Parameter
import xml.etree.ElementTree as ET
import pandas as pd
import numpy as np
import pyautogui
import pyperclip
import openpyxl
import win32api
import win32gui
import win32con
import shutil
import time
import math
import os
import re

N = 5  # 批量产生随机数时随机数的个数,建立的文件夹个数,替换生成的文件个数,操纵软件计算的次数,获得的安全系数的个数
#程序的计算流程控制，列表中第一个字典中的“是”和“否”代表开启和关闭相关功能
process_control = [
    {'是否生成所需参数保存至Excel表格':'否'},
    {'是否添加或者替换数据':'否'},
    {'是否操作软件进行批量计算':'否'},
    {'是否提取计算结果至Excel表格':'是'}
] 
#列表中第一个字典中的“是”和“否”代表开启和关闭相关功能,它们不能同时为“是”，后面字典中的数字代表需要替换或添加的参数数量，添加区域坐标功能未完成
area_coordinates_options = [
    {'替换区域坐标':'否', '添加区域坐标':'否'},
    {'横坐标X':6,'正态分布均值1': 25, '正态分布方差1': 56.25,'正态分布均值2': 5, '正态分布方差2': 2.25,'正态分布均值3': 5, '正态分布方差3': 2.25,'正态分布均值4': 5, '正态分布方差4': 2.25,'正态分布均值5': 5, '正态分布方差5': 2.25,'正态分布均值6': 5, '正态分布方差6': 2.25, 'LeftSideLeftPt X': 2, 'LeftSideRightPt X': 3, 'RightSideLeftPt X': 4, 'RightSideRightPt X': 5},
    {'纵坐标Y':6,'正态分布均值1': 5, '正态分布方差1': 2.25,'正态分布均值2': 5, '正态分布方差2': 2.25,'正态分布均值3': 5, '正态分布方差3': 2.25,'正态分布均值4': 5, '正态分布方差4': 2.25,'正态分布均值5': 5, '正态分布方差5': 2.25,'正态分布均值6': 5, '正态分布方差6': 2.25, 'LeftSideLeftPt Y': 2, 'LeftSideRightPt Y': 3, 'RightSideLeftPt Y': 4, 'RightSideRightPt Y': 5},
    {'其它坐标X':4, 'LeftSideLeftPt X': 2, 'LeftSideRightPt X': 3, 'RightSideLeftPt X': 4, 'RightSideRightPt X': 5},
    {'其它坐标Y':4, 'LeftSideLeftPt Y': 2, 'LeftSideRightPt Y': 3, 'RightSideLeftPt Y': 4, 'RightSideRightPt Y': 5}
]   
#列表中第一个字典中的“是”和“否”代表开启和关闭相关功能,它们不能同时为“是”，后面字典中参数名后数字代表需要替换或添加的参数数量，添加土体参数功能未完成
soil_parameter_options = [
    {'替换土体参数':'是', '添加土体参数':'否'},
    {'内摩擦角':1,'正态分布均值1': 25, '正态分布方差1': 56.25},
    {'粘聚力':1,'正态分布均值1': 5, '正态分布方差1': 2.25},
    {'重度':1,'正态分布均值1': 5, '正态分布方差1': 2.25}
]   
#列表中第一个字典中的“是”和“否”代表开启和关闭相关功能,它们不能同时为“是”，变化浸润线利用了3δ原则
saturation_line_options = [
    {'替换浸润线':'否', '添加浸润线':'否'},
    {'浸润线横坐标':['6', '12', '18', '24'],'浸润线纵坐标下限':[6, 7, 9, 13],'浸润线纵坐标上限':[6, 11, 15, 17]},
    {'浸润线Ywn':1, '正态分布均值1':0.5, '正态分布方差1':0.5 / 3, '恒定分布0':0.5, '无分布0':0}
]
#批处理操作软件相关参数
批处理操作路径 = 'C:\\Program Files (x86)\\GEO-SLOPE\\GeoStudio 9\\Bin\\GeoStudio.exe'  #路径字符串中必须采用双反斜杠
命令10匹配存在的窗口标题 = []                                                                   #盛放命令10匹配的窗口标题
hwnd_title = {}                                                                               #盛放获取的窗口标题
                                                                                              # "batch_instruction_set1"指令位于批处理代码执行处(函数check_blank_calculate中)
batch_instruction_set2 = [
    {'指令1':3,'内容':'alt,w,down,down,down,down,enter,space','循环次数':1},
    {'指令2':6,'内容':8,'循环次数':1}
]                                                                                             #软件窗口处的操作
                                                                                              # "batch_instruction_set3"指令位于批处理代码执行处(函数check_blank_calculate中)
batch_instruction_set4 = [
    {'指令1':8,'内容':'taskkill /F /IM GeoStudio.exe','循环次数':1}
]                                                                                             #关闭批处理软件

#通常情况下批处理不需要修改的部分
key_map = {
    "0": 96, "1": 97, "2": 98, "3": 99, "4": 100, "5": 101, "6": 102, "7": 103, "8": 104, "9": 105, 
    '*': 106, '+': 107, '-': 109, '.': 110, '/': 111,
    'F1': 112, 'F2': 113, 'F3': 114, 'F4': 115, 'F5': 116, 'F6': 117, 'F7': 118, 'F8': 119,
    'F9': 120, 'F10': 121, 'F11': 122, 'F12': 123, 'F13': 124, 'F14': 125, 'F15': 126, 'F16': 127,
    "A": 65, "B": 66, "C": 67, "D": 68, "E": 69, "F": 70, "G": 71, "H": 72, "I": 73, "J": 74,
    "K": 75, "L": 76, "M": 77, "N": 78, "O": 79, "P": 80, "Q": 81, "R": 82, "S": 83, "T": 84,
    "U": 85, "V": 86, "W": 87, "X": 88, "Y": 89, "Z": 90,
    'BACKSPACE': 8, 'TAB': 9, 'TABLE': 9, 'CLEAR': 12, 'ENTER': 13, 'SHIFT': 16, 'CTRL': 17,
    'CONTROL': 17, 'ALT': 18, 'ALTER': 18, 'PAUSE': 19, 'BREAK': 19, 'CAPSLK': 20, 'CAPSLOCK': 20, 
    'ESC': 27,'SPACE': 32, 'SPACEBAR': 32, 'PGUP': 33, 'PAGEUP': 33, 'PGDN': 34,'PAGEDOWN': 34, 
    'END': 35, 'HOME': 36,'LEFT': 37, 'UP': 38, 'RIGHT': 39, 'DOWN': 40, 'SELECT': 41,'PRTSC': 42, 
    'PRINTSCREEN': 42, 'SYSRQ': 42,'SYSTEMREQUEST': 42, 'EXECUTE': 43, 'SNAPSHOT': 44,'INSERT': 45, 
    'DELETE': 46, 'HELP': 47, 'WIN': 91,'WINDOWS': 91, 'NMLK': 144,'NUMLK': 144, 'NUMLOCK': 144, 
    'SCRLK': 145, '[': 219, ']': 221, 
    '音量加':175, '音量减':174, '停止':179, '静音':173, '浏览器':172, '邮件':180, '搜索':170
    }                                                                                          #win32api按键值对应键盘按键，依次包括数字键盘按键、功能键按键、字母键按键、控制键按键、多媒体键按键
work_dir = os.path.dirname(os.path.abspath(__file__))                                          # 获取工作文件夹路径
original_file_name = '模型'                                                                    # 批处理替换数据原始文件名称
original_file_name_type = '模型.xml'                                                           # 批处理替换数据带后缀原始文件名称
doc_name1 = '模型参数：'
parameter_options = []
main_folder_name1 = '批处理目标文件：'
for dic in area_coordinates_options + soil_parameter_options + saturation_line_options:
    for k,v in dic.items():
        if v == '是':
            parameter_options.append(k)
            doc_name1 += ',' + k
            main_folder_name1 += ',' + k
doc_name1 += '.xlsx'

doc_name1 = doc_name1.partition(",")[0] + doc_name1.partition(",")[2]                            # 存储各类型批处理数据的文件名
main_folder_name1 = main_folder_name1.partition(",")[0] + main_folder_name1.partition(",")[2]    # 建立各类型批处理主的文件夹名

def main():
    if process_control[0]['是否生成所需参数保存至Excel表格'] == '是':
        generate_random_numbers(doc_name1)                                                # 产生随机数，若需读取准备好的数据可将此行注释
    
    if process_control[1]['是否添加或者替换数据'] == '是':
        rd_data = []                                                                          # 定义盛放Excel文件中数据的数组
        io = work_dir + '\\' + doc_name1                                          
                      # 为Excel文件所在位置路径附上Excel文件名及格式后缀
        rd_data.append(pd.read_excel(io))                                                     # 读取从外部获得的储存于Excel中的数据
        doc_col_name = list(rd_data[0])                                                       # 读取Excel文件的列名
        parameters_name = doc_col_name[1:]                                                    # 读取Excel文件的各参数名
        path = work_dir + '\\' + original_file_name_type                                      # 为原始xml文件所在位置路径附上原始xml文件名及格式后缀

        if area_coordinates_options[0]['添加区域坐标'] == '是' or soil_parameter_options[0]['添加土体参数'] == '是' or saturation_line_options[0]['添加浸润线'] == '是':
            replace_add_data(main_folder_name1,rd_data[0],parameters_name)                   # 批量替换土体参数，添加浸润线
        else:
            replace_data(path, main_folder_name1, rd_data[0], 0, parameters_name)         # 批量替换区域坐标
            add_data(path, main_folder_name1,rd_data[0],0)                                   # 批量添加水压线
        move_file(main_folder_name1)                                                      # 批量移动文件至相应的文件夹
    
    if process_control[2]['是否操作软件进行批量计算'] == '是':
        check_blank_calculate(main_folder_name1)                                          # 操纵软件进行批量运算
    
    if process_control[3]['是否提取计算结果至Excel表格'] == '是':
        get_data(doc_name1, main_folder_name1)                                        # 获取安全系数以及滑动体积

 # 根据各种分布生成数据
def generate_random_numbers(d_name):
    soil_parameters = {}
    
    if soil_parameter_options[0]['替换土体参数'] == '是' or soil_parameter_options[0]['添加土体参数'] == '是':
        #生成土体相关参数
        for i in soil_parameter_options:
            if '内摩擦角' in i:
                parameter_name = '内摩擦角'
            elif '粘聚力' in i:
                parameter_name = '粘聚力'
            elif '重度' in i:
                parameter_name = '重度'
            else:
                continue
            for j in range(i[parameter_name]):
                for k in i:
                    str_num = str(j + 1)
                    if '正态分布均值' + str_num == k:
                        R = np.random.normal(i['正态分布均值' + str_num], i['正态分布方差' + str_num], N)
                        soil_parameters[parameter_name + str_num] = R
                    elif '对数正态分布均值' + str_num == k:
                        xm = math.log((i['正态分布均值' + str_num] ** 2) / math.sqrt(i['正态分布方差' + str_num] + (i['正态分布均值' + str_num] ** 2)))
                        xd = math.sqrt(math.log(i['正态分布方差' + str_num] / (i['正态分布均值' + str_num] ** 2) + 1))
                        R = np.random.lognormal(xm, xd, N)
                        soil_parameters[parameter_name + str_num] = R
    
    if area_coordinates_options[0]['替换区域坐标'] == '是' or area_coordinates_options[0]['添加区域坐标'] == '是':
        #生成区域坐标
        for i in area_coordinates_options:
            if '横坐标X' in i:
                parameter_name = '横坐标X'
            elif '纵坐标Y' in i:
                parameter_name = '纵坐标Y'
            else:
                continue
            for j in range(i[parameter_name]):
                for k in i:
                    str_num = str(j + 1)
                    if '正态分布均值' + str_num == k:
                        R = np.random.normal(i['正态分布均值' + str_num], i['正态分布方差' + str_num], N)
                        soil_parameters[parameter_name + str_num] = R

                    elif '对数正态分布均值' + str_num == k:
                        xm = math.log((i['正态分布均值' + str_num] ** 2) / math.sqrt(i['正态分布方差' + str_num] + (i['正态分布均值' + str_num] ** 2)))
                        xd = math.sqrt(math.log(i['正态分布方差' + str_num] / (i['正态分布均值' + str_num] ** 2) + 1))
                        R = np.random.lognormal(xm, xd, N)
                        soil_parameters[parameter_name + str_num] = R

    if saturation_line_options[0]['替换浸润线'] == '是' or saturation_line_options[0]['添加浸润线'] == '是':
        #生成变化浸润线相关参数
        for i in saturation_line_options:
            if '浸润线Ywn' in i:
                parameter_name = '浸润线Ywn'
            else:
                continue
            for j in range(i[parameter_name]):
                for k in i:
                    str_num = str(j + 1)
                    if '正态分布均值' + str_num == k:
                        R = np.random.normal(i['正态分布均值' + str_num], i['正态分布方差' + str_num], N)
                        soil_parameters[parameter_name + str_num] = R
                    elif '恒定分布' + str_num == k:
                        R = [i['恒定分布' + str_num] for m in R]
                        soil_parameters[parameter_name + str_num] = R
                    elif '无分布' + str_num == k:
                        R = [i['无分布' + str_num] for m in R]
                        soil_parameters[parameter_name + str_num] = R

    # 输出参数
    sp = pd.DataFrame(soil_parameters)
    sp.index = sp.index + 1
    if os.path.exists(work_dir + '\\' + d_name):
        print('已存在该文件')
    else:
        sp.to_excel(work_dir + '\\' + d_name, sheet_name='原始数据', index=1, index_label='模型')

# 批量创建文件夹
def make_dir(m_folder_name):
    path = work_dir + '\\' + m_folder_name
    if not os.path.exists(path):
        os.mkdir(path)
    for i in range(N):
        path = work_dir + '\\' + m_folder_name
        path = path + "\\" + str(i + 1) + '_Runs'
        if not os.path.exists(path):
            os.mkdir(path)


# 将替换过的文件移动到新建的文件夹中
def move_file(m_folder_name):
    cur_file_dir = work_dir + '\\' + m_folder_name
    for i in range(N):
        des_file_dir = work_dir + '\\' + m_folder_name + '\\' + str(i + 1) + '_Runs'
        file_dir = cur_file_dir + '\\' + original_file_name + str(i + 1) + '.xml'
        shutil.move(file_dir, des_file_dir)


# 美化标签
def prettyXml(element, indent, newline, level=0):  # elemnt为传进来的Elment类，参数indent用于缩进，newline用于换行
    if element:  # 判断element是否有子元素
        if element.text == None or element.text.isspace():  # 如果element的text没有内容
            element.text = newline + indent * (level + 1)
        else:
            element.text = newline + indent * (level + 1) + element.text.strip() + newline + indent * (level + 1)
    # else:                                                       # 此处两行如果把注释去掉，Element的text也会另起一行
    # element.text = newline + indent * (level + 1) + element.text.strip() + newline + indent * level
    temp = list(element)  # 将elemnt转成list
    for subelement in temp:
        if temp.index(subelement) < (len(temp) - 1):  # 如果不是list的最后一个元素，说明下一个行是同级别元素的起始，缩进应一致
            subelement.tail = newline + indent * (level + 1)
        else:  # 如果是list的最后一个元素， 说明下一行是母元素的结束，缩进应该少一个
            subelement.tail = newline + indent * level
        prettyXml(subelement, indent, newline, level=level + 1)  # 对子元素进行递归操作

# 批量替换数据
def replace_data(file_path, m_folder_name, rd_data, j, parameters_name):
    # tree = ET.parse(work_dir + '\\' + original_file_name_type)
    #变化j参数在同时进行添加和替换操作时应用
    if file_path == work_dir + '\\' + original_file_name_type:
        n = range(0, N)
    else:
        n = range(j, j + 1)
    
    tree = ET.parse(file_path)
    root = tree.getroot()                                  # 获取XML文件根节点
    Geometries = root.find('Geometries')                   # 获取子节点Geometries
    Materials = root.find('Materials')                     # 获取子节点Materials
    StabilityItems = root.find('StabilityItems')           # 获取子节点StabilityItems
    d = [round(saturation_line_options[1]['浸润线纵坐标上限'][k] - saturation_line_options[1]['浸润线纵坐标下限'][k], 3) for k in range(len(saturation_line_options[1]['浸润线纵坐标下限']))]  # 两水位线各点的间距

    for i in n:
        if area_coordinates_options[0]['替换区域坐标'] == '是' and area_coordinates_options[0]['添加区域坐标'] == '否':            
            for Geometry in Geometries.findall('Geometry'):
                Points = Geometry.find('Points')
                for parameter_name in parameters_name:
                    for Point in Points.findall('Point'):
                        if '横坐标' in parameter_name:
                            if '1' in parameter_name and '1' in Point.attrib['ID']:
                                Point.attrib['X'] = str(rd_data[parameter_name][i])
                            elif '2' in parameter_name and '2' in Point.attrib['ID']:
                                Point.attrib['X'] = str(rd_data[parameter_name][i])
                            elif '3' in parameter_name and '3' in Point.attrib['ID']:
                                Point.attrib['X'] = str(rd_data[parameter_name][i])
                            elif '4' in parameter_name and '4' in Point.attrib['ID']:
                                Point.attrib['X'] = str(rd_data[parameter_name][i])
                            elif '5' in parameter_name and '5' in Point.attrib['ID']:
                                Point.attrib['X'] = str(rd_data[parameter_name][i])
                            elif '6' in parameter_name and '6' in Point.attrib['ID']:
                                Point.attrib['X'] = str(rd_data[parameter_name][i])
                        if '纵坐标' in parameter_name:
                            if '1' in parameter_name and '1' in Point.attrib['ID']:
                                Point.attrib['Y'] = str(rd_data[parameter_name][i])
                            elif '2' in parameter_name and '2' in Point.attrib['ID']:
                                Point.attrib['Y'] = str(rd_data[parameter_name][i])
                            elif '3' in parameter_name and '3' in Point.attrib['ID']:
                                Point.attrib['Y'] = str(rd_data[parameter_name][i])
                            elif '4' in parameter_name and '4' in Point.attrib['ID']:
                                Point.attrib['Y'] = str(rd_data[parameter_name][i])
                            elif '5' in parameter_name and '5' in Point.attrib['ID']:
                                Point.attrib['Y'] = str(rd_data[parameter_name][i])
                            elif '6' in parameter_name and '6' in Point.attrib['ID']:
                                Point.attrib['Y'] = str(rd_data[parameter_name][i])
            for StabilityItem in StabilityItems.findall('StabilityItem'):
                Entry = StabilityItem.find('Entry')
                SlipSurface = Entry.find('SlipSurface')
                EntryExit = SlipSurface.find('EntryExit')
                LeftSideLeftPt = EntryExit.find('LeftSideLeftPt')
                LeftSideRightPt = EntryExit.find('LeftSideRightPt')
                RightSideLeftPt = EntryExit.find('RightSideLeftPt')
                RightSideRightPt = EntryExit.find('RightSideRightPt')
                for parameter_name in parameters_name:
                    if 'LeftSideLeftPt' in parameter_name and 'X' in parameter_name:
                        LeftSideLeftPt.attrib['X'] = str(rd_data[parameter_name][i])
                    elif 'LeftSideLeftPt' in parameter_name and 'Y' in parameter_name:
                        LeftSideLeftPt.attrib['Y'] = str(rd_data[parameter_name][i])
                    elif 'LeftSideRightPt' in parameter_name and 'X' in parameter_name:
                        LeftSideRightPt.attrib['X'] = str(rd_data[parameter_name][i])
                    elif 'LeftSideRightPt' in parameter_name and 'Y' in parameter_name:
                        LeftSideRightPt.attrib['Y'] = str(rd_data[parameter_name][i])
                    elif 'RightSideLeftPt' in parameter_name and 'X' in parameter_name:
                        RightSideLeftPt.attrib['X'] = str(rd_data[parameter_name][i])
                    elif 'RightSideLeftPt' in parameter_name and 'Y' in parameter_name:
                        RightSideLeftPt.attrib['Y'] = str(rd_data[parameter_name][i])
                    elif 'RightSideRightPt' in parameter_name and 'X' in parameter_name:
                        RightSideRightPt.attrib['X'] = str(rd_data[parameter_name][i])
                    elif 'RightSideRightPt' in parameter_name and 'Y' in parameter_name:
                        RightSideRightPt.attrib['Y'] = str(rd_data[parameter_name][i])
        
        if soil_parameter_options[0]['替换土体参数'] == '是' and soil_parameter_options[0]['添加土体参数'] == '否':
            for Material in Materials.findall('Material'):
                ID = Material.find('ID').text
                StressStrain = Material.find('StressStrain')
                for parameter_name in parameters_name:
                    if '内摩擦角' + ID in parameter_name:
                        StressStrain.find('PhiPrime').text = str(rd_data[parameter_name][i])
                    elif '粘聚力' + ID in parameter_name:
                        StressStrain.find('CohesionPrime').text = str(rd_data[parameter_name][i])
                    elif '重度' + ID in parameter_name:
                        StressStrain.find('UnitWeight').text = str(rd_data[parameter_name][i])

        if saturation_line_options[0]['替换浸润线'] == '是' and saturation_line_options[0]['添加浸润线'] == '否':
            for StabilityItem in StabilityItems.findall('StabilityItem'):
                Entry = StabilityItem.find('Entry')
                DataPoints = Entry.find('DataPoints')
                DataPoint = DataPoints.findall('DataPoint')
                for k in DataPoint:
                    k.attrib['X'] = saturation_line_options[1]['浸润线横坐标'][int(k.attrib['Number']) - 1]
                    k.attrib['Y'] = str(
                        saturation_line_options[1]['浸润线纵坐标下限'][i][int(k.attrib['Number']) - 1] + d[int(k.attrib['Number']) - 1] * replace_data['Ywn'][i])
        
        location = work_dir + '\\' + m_folder_name
        if not os.path.exists(location):
            make_dir(m_folder_name)
            tree.write(location + '/' + original_file_name + str(i + 1) + '.xml', encoding='utf-8',
                    xml_declaration=True)
        else:
            tree.write(location + '/' + original_file_name + str(i + 1) + '.xml', encoding='utf-8',
                    xml_declaration=True)

# 批量添加浸润线
def add_data(file_path, m_folder_name, rd_data, j):
    if file_path == work_dir + '\\' + original_file_name_type:
        n = range(0, N)
    else:
        n = range(j, j + 1)
    d = [round(saturation_line_options[1]['浸润线纵坐标上限'][k] - saturation_line_options[1]['浸润线纵坐标下限'][k], 3) for k in range(len(saturation_line_options[1]['浸润线纵坐标下限']))]  # 两水位线各点的间距
    # tree = ET.parse(work_dir + '\\' + original_file_name_type)
    tree = ET.parse(file_path)
    root = tree.getroot()                         # 获取XML文件根节点
    StabilityItems = root.find('StabilityItems')  # 获取子节点StabilityItems
    WaterItems = root.find('WaterItems')          #获取子节点
    for i in n:
        if saturation_line_options[0]['替换浸润线'] == '否' and saturation_line_options[0]['添加浸润线'] == '是':
            for StabilityItem in StabilityItems.findall('StabilityItem'):
                Entry1 = StabilityItem.find('Entry')
                SubElement_Entry0 = ET.SubElement(Entry1, 'DataPoints',
                                  attrib={'Len': str(len(saturation_line_options[1]['浸润线横坐标']))})                # 添加含attrib的标签，atrib后面接的是字典格式的
                for k in range(len(saturation_line_options[1]['浸润线横坐标'])):
                    ET.SubElement(SubElement_Entry0, 'DataPoint',
                                attrib={'Number': str(k + 1), 'X': saturation_line_options[1]['浸润线横坐标'][k], 'Y': str(saturation_line_options[1]['浸润线纵坐标下限'][k] + d[k] * rd_data['浸润线Ywn1'][i])})                
                
                SubElement_Entry1 = ET.SubElement(Entry1, 'PiezometricLines',
                                  attrib={'Len': '1'})                        # 添加含attrib的标签，atrib后面接的是字典格式的
                SubElement_Entry1_PiezometricLines = ET.SubElement(SubElement_Entry1, 'PiezometricLine',
                                                                attrib={})    # 添加含attrib的标签，atrib后面接的是字典格式的
                SubElement_PiezometricLine0_ID = ET.SubElement(SubElement_Entry1_PiezometricLines, 'ID')
                SubElement_PiezometricLine0_ID.text = '1'                     # 配置text，注意不能直接用int类型的
                SubElement_PiezometricLine0_DataPoints = ET.SubElement(SubElement_Entry1_PiezometricLines, 'DataPoints',
                                                                    attrib={'Len': str(len(saturation_line_options[1]['浸润线横坐标']))})
                for k in reversed(range(len(saturation_line_options[1]['浸润线横坐标']))):
                    SubElement_PiezometricLine0_DataPoint0 = ET.SubElement(SubElement_PiezometricLine0_DataPoints,
                                                                        'DataPoint')
                    SubElement_PiezometricLine0_DataPoint0.text = str(k + 1)  # 配置text，注意不能直接用int类型的
                SubElement_Entry2 = ET.SubElement(Entry1, 'MaterialUsesPiezs',
                                                attrib={'Len': '1'})          # 添加含attrib的标签，atrib后面接的是字典格式的
                ET.SubElement(SubElement_Entry2, 'MaterialUsesPiez', attrib={'ID': '1', 'UsesID': '1'})
                prettyXml(Entry1, '    ', '\n')                                # 美化标签
                
            for WaterItem in WaterItems.findall('WaterItem'):
                Entry2 = WaterItem.find('Entry')
                SubElement_Entry3 = ET.SubElement(Entry2, 'ResultInputInfo', attrib={})  # 添加含attrib的标签，atrib后面接的是字典格式的
                SubElement_ResultInputInfo_Option = ET.SubElement(SubElement_Entry3,
                                                                'Option')                # 添加含attrib的标签，atrib后面接的是字典格式的
                SubElement_ResultInputInfo_Option.text = 'PiezoLine'                     # 配置text，注意不能直接用int类型的
                prettyXml(Entry2, '    ', '\n')
                
            location = work_dir + '\\' + m_folder_name
            if not os.path.exists(location):
                make_dir(m_folder_name)
                tree.write(location + '/' + original_file_name + str(i + 1) + '.xml', encoding='utf-8',
                        xml_declaration=True)
            else:
                tree.write(location + '/' + original_file_name + str(i + 1) + '.xml', encoding='utf-8',
                        xml_declaration=True)


def replace_add_data(m_folder_name, rd_data, parameters_name):
    files_folder_path = work_dir + '\\' + m_folder_name
    original_file_path = work_dir + '\\' + original_file_name_type
    if area_coordinates_options[0]['替换区域坐标'] == '否' and area_coordinates_options[0]['添加区域坐标'] == '是':
        pass
    if soil_parameter_options[0]['替换土体参数'] == '否' and soil_parameter_options[0]['添加土体参数'] == '是':
        pass
    if saturation_line_options[0]['替换浸润线'] == '否' and saturation_line_options[0]['添加浸润线'] == '是':
        for i in range(N):
            files_path = files_folder_path + '\\' + original_file_name + str(i + 1) + '.xml'
            if os.path.exists(files_path):
                tree = ET.parse(files_path)
                root = tree.getroot()  # 获取XML文件根节点
                StabilityItems = root.find('StabilityItems')
                StabilityItem = StabilityItems.find('StabilityItem')
                Entry = StabilityItem.find('Entry')
                DataPoints = Entry.find('DataPoints')
                if DataPoints:
                    replace_data(files_path, m_folder_name, rd_data, i, parameters_name)
            else:
                add_data(original_file_path, m_folder_name, rd_data, 0)
                replace_data(files_path, m_folder_name, rd_data, i, parameters_name)

#鼠标相关操作，由B站UP主 不高兴就喝水 提供
def mouseClick(clickTimes,lOrR,img):
    while True:
        location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
        if location is not None:
            pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
            break
        print("未找到匹配图片,0.1秒后重试")
        time.sleep(0.1)

#用来按键还是文字输入,由B站up主 尔茄无双 提供。
def presskey(hk_g_inputValue):
    keys = hk_g_inputValue.split(',')
    for key in keys:
        if isinstance(key, str):
            win32api.keybd_event(key_map[key.upper()], win32api.MapVirtualKey(key_map[key.upper()], 0), 0, 0)
            win32api.keybd_event(key_map[key.upper()], win32api.MapVirtualKey(key_map[key.upper()], 0), win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.1)
    print("执行了：",hk_g_inputValue)
    time.sleep(0.1)

#判断热键组合个数还是文字输入,由B站up主 尔茄无双 提供
def hotkey_get(hk_g_inputValue):
    try:
        keys = hk_g_inputValue.split(',')
        for key in keys:
            if isinstance(key, str):
                win32api.keybd_event(key_map[key.upper()], win32api.MapVirtualKey(key_map[key.upper()], 0), 0, 0)
        for key in keys:
            if isinstance(key, str):
                win32api.keybd_event(key_map[key.upper()], win32api.MapVirtualKey(key_map[key.upper()], 0), win32con.KEYEVENTF_KEYUP, 0)
        print("执行了：",hk_g_inputValue)
        time.sleep(0.1)
    except:
        pyperclip.copy(hk_g_inputValue)
        pyautogui.hotkey('ctrl', 'v')

#获取窗口标题
def get_all_hwnd(hwnd, mouse):
    if (win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd)):
        hwnd_title.update({hwnd: win32gui.GetWindowText(hwnd)})

def instruction_set_execution(bis):
    i = 0
    while i < len(bis):
        cmdType = bis[i]['指令'+ str(i+1)]     #读取指令号
        cmdContent = bis[i]['内容']            #读取指令内容
        cycles = bis[i]['循环次数']                #读取指令循环次数
        if cmdType == 1:           
            for j in range(cycles):
                if '单击左键' in cmdContent:
                    mouseClick(1,"left", cmdContent)
                    print("单击左键", cmdContent)
                elif '双击左键' in cmdContent:
                    mouseClick(2,"left", cmdContent)
                    print("双击左键", cmdContent)    
                elif '单击右键' in cmdContent:
                    mouseClick(1,"right", cmdContent)
                    print("单击右键", cmdContent) 
        #2鼠标滚轮操作
        elif cmdType == 2:
            for j in range(cycles):
                pyautogui.scroll(cmdContent)
                print("滚轮滑动",int(cmdContent),"距离")     
        #3键盘按键
        elif cmdType == 3:
            for j in range(cycles):
                presskey(cmdContent)
                time.sleep(0.5)
        #4键盘热键组合
        elif cmdType == 4:
            for j in range(cycles):
                hotkey_get(cmdContent)
                time.sleep(0.5)
        #5输入字符串
        elif cmdType == 5:
            for j in range(cycles):
                pyperclip.copy(cmdContent)
                pyautogui.hotkey('ctrl','v')
                print("输入:",cmdContent) 
                time.sleep(0.5)                                       
        #6等待
        elif cmdType == 6:
            for j in range(cycles):
                time.sleep(cmdContent)
                print("等待",cmdContent,"秒")
        #7粘贴当前时间
        elif cmdType == 7:      
            for j in range(cycles):
                localtime = time.strftime("%Y-%m-%d %H：%M：%S", time.localtime())  #设置本机当前时间。
                pyperclip.copy(localtime)
                pyautogui.hotkey('ctrl','v')
                print("粘贴了本机时间:",localtime)
                time.sleep(0.5)
        #8系统命令集
        elif cmdType == 8: 
            for j in range(cycles):
                os.system(cmdContent)
                print("运行系统命令:",cmdContent)
                time.sleep(0.5) 
        #9利用指定程序打开指定文件或者打开某程序
        elif cmdType == 9:
            cmdContent = cmdContent.split(',') 
            for j in range(cycles):
                if len(cmdContent) == 1:
                    win32api.ShellExecute(0, 'open', cmdContent[0], "", "", 1)                                      #只打开指定程序，1个参数：指定程序所在路径
                elif len(cmdContent) == 2:
                    win32api.ShellExecute(0, 'open', cmdContent[0], cmdContent[1], "", 1)  #利用指定程序打开默认文件路径下指定文件，2个参数：指定程序所在路径、文件名称（包括格式后缀）
                elif len(cmdContent) == 3:
                    win32api.ShellExecute(0, 'open', cmdContent[0], cmdContent[1], cmdContent[2], 1)  #利用指定程序打开指定文件下指定文件，3个参数：指定程序所在路径、文件名称（包括格式后缀）、文件所在文件夹路径
                print("运行系统命令:",cmdContent)
                time.sleep(0.5) 
        #10 获得桌面上所有打开窗口的标题，可选项：匹配窗口标题
        elif cmdType == 10:
            cmdContent = cmdContent.split(',')
            for j in range(cycles): 
                win32gui.EnumWindows(get_all_hwnd, 0)
                print('桌面存在窗口:', hwnd_title)
                for k in cmdContent:
                    for m in hwnd_title.values():
                        if k == m:
                            命令10匹配存在的窗口标题.append(k)
                            print('匹配窗口:', k, '存在')
                time.sleep(0.5) 
        i += 1

# 查找中断位置并计算或者进行批处理，操纵GeoStudio进行计算
def check_blank_calculate(m_folder_name):
    file_dir = work_dir + '\\' + m_folder_name
    # print(file_dir)
    file_handle = open(work_dir + '\\' + '计算位置.txt', mode='w')
    location = []
    for i in range(N):
        cur_dir1 = file_dir + '\\' + str(i + 1) + '_Runs\\SLOPE&3W Analysis\\001'
        cur_dir2 = file_dir + '\\' + str(i + 1) + '_Runs\\SLOPE&3W 分析\\001'
        # print(cur_dir)
        if os.path.exists(cur_dir1):
            cur_dir = cur_dir1
        elif os.path.exists(cur_dir2):
            cur_dir = cur_dir2
        if not os.path.exists(cur_dir):
            location.append(i + 1)
            file_handle.write(str(i + 1) + '\n')

    for i in location:
        Path = work_dir + '\\' + m_folder_name + '\\' + str(i) + '_Runs\\'
        file = original_file_name + str(i) + '.xml'
        print(Path, file)
        # 利用指定程序打开指定文件夹下的指定文件，等待并匹配操作窗口，匹配的窗口出现后进行快捷键操作
        batch_instruction_set1 = [
            {'指令1':9,'内容':批处理操作路径 + ',' + file + ',' + Path,'循环次数':1},
            {'指令2':6,'内容':10,'循环次数':1},
            {'指令3':10,'内容':original_file_name + str(i) + '.xml - GeoStudio 2018 R2 (SLOPE/W Definition)','循环次数':1}
        ]
        instruction_set_execution(batch_instruction_set1)
        for j in 命令10匹配存在的窗口标题:
            if j == original_file_name + str(i) + '.xml - GeoStudio 2018 R2 (SLOPE/W Definition)':
                instruction_set_execution(batch_instruction_set2)

        # 等待并匹配是否存在结果窗口，若存在则关闭程序，否则继续
        batch_instruction_set3 = [
            {'指令1':6,'内容':2,'循环次数':1},
            {'指令2':10,'内容':original_file_name + str(i) + '.xml - GeoStudio 2018 R2 (SLOPE/W Results)','循环次数':1},
            {'指令3':10,'内容':original_file_name + str(i) + '.xml - GeoStudio 2018 R2 (SLOPE/W Definition)','循环次数':1}
        ]
        instruction_set_execution(batch_instruction_set3)
        for j in 命令10匹配存在的窗口标题:
            if j == original_file_name + str(i) + '.xml - GeoStudio 2018 R2 (SLOPE/W Results)':
                instruction_set_execution(batch_instruction_set4)
            elif j == original_file_name + str(i) + '.xml - GeoStudio 2018 R2 (SLOPE/W Definition)':
                instruction_set_execution(batch_instruction_set4)
            else:
                continue

# 获取安全系数和滑动体积并输出
def get_data(d_name, m_main_folder_name):
    file_dir = work_dir + '\\' + m_main_folder_name
    # print(file_dir)
    fs_data = ['安全系数']
    sv_data = ['滑动体积']
    for i in range(N):
        cur_dir1 = file_dir + '\\' + str(i + 1) + '_Runs\\SLOPE&3W Analysis\\001'
        cur_dir2 = file_dir + '\\' + str(i + 1) + '_Runs\\SLOPE&3W 分析\\001'
        # print(cur_dir)
        if os.path.exists(cur_dir1):
            cur_dir = cur_dir1
        elif os.path.exists(cur_dir2):
            cur_dir = cur_dir2
        msp_number = 1
        for root, dirs, files in os.walk(cur_dir):
            for file in files:
                if 'lambdafos_' in file and file.endswith('.csv'):
                    msp_number = int(re.findall('\d+', file)[0]) - 1
                    cur_files1 = root + '\\' + file
                    data1 = pd.read_csv(cur_files1)
                    if len(data1) == 0:
                        fs_data.append(0)
                    else:
                        fs_data.append(data1['FOSByMoment'][len(data1) - 1])
                if 'slip_surface' in file and file.endswith('.csv'):
                    cur_files2 = root + '\\' + file
                    data2 = pd.read_csv(cur_files2)
                    sv_data.append(data2['SlipVolume'][msp_number])
    wb = openpyxl.load_workbook(work_dir + '\\' + d_name)  # 生成一个已存在的wookbook对象
    wb1 = wb.active  # 激活sheet
    tcl_number = 1
    if area_coordinates_options[0]['替换区域坐标'] == '是' or area_coordinates_options[0]['添加区域坐标'] == '是':
        tcl_number += area_coordinates_options[1]['横坐标X'] + area_coordinates_options[2]['纵坐标Y'] + area_coordinates_options[3]['其它坐标X'] + area_coordinates_options[4]['其它坐标Y']
    elif soil_parameter_options[0]['替换土体参数'] == '是' or soil_parameter_options[0]['添加土体参数'] == '是':
        tcl_number += soil_parameter_options[1]['内摩擦角'] + soil_parameter_options[2]['粘聚力'] + soil_parameter_options[3]['重度']
    elif saturation_line_options[0]['替换浸润线'] == '是' or saturation_line_options[0]['添加浸润线'] == '是':
        tcl_number += 1
    for i in range(len(fs_data)):
        wb1.cell(i + 1, tcl_number).value = fs_data[i]
    for i in range(len(sv_data)):
        wb1.cell(i + 1, tcl_number + 1).value = sv_data[i]
    wb.save(work_dir + '\\' + d_name)  # 保存



if __name__ == "__main__":
    main()

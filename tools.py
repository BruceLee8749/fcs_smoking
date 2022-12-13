from shutil import copyfile
from configparser import ConfigParser
from time import strftime
import os
from openpyxl import load_workbook
import logging
from PIL import Image
import shutil
import math
import operator
from functools import reduce

LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
DATE_FORMAT = "%m/%d/%Y %H:%M:%S %p"
logging.basicConfig(filename='critical.log', level=logging.CRITICAL, format=LOG_FORMAT, datefmt=DATE_FORMAT)
project_path = f"{os.path.dirname(__file__)}"


def critical(text):
    logging.critical(text)


def dict_in(dict1, dict2):  # 检查dict2是否包含dict1
    try:
        lis1, lis2 = [], []
        for key in dict1:
            lis1.append(dict1[key])
            lis2.append(dict_get(dict2, key, None))
        return lis1 == lis2
    except KeyError:
        return False


def dict_get(dic, key, default):  # 嵌套字典检查
    for k, v in dic.items():
        if k == key:
            return v
        else:
            if isinstance(v, dict):
                ret = dict_get(v, key, default)
                if ret is not default:
                    return ret
    return default


def get_all_xlsx_file(folder):
    file_list = []
    for file in os.listdir(folder):
        if file.endswith('.xlsx'):
            file_list.append(file)
    return file_list


def get_conf(sec, key):
    cf = ConfigParser()
    pro_path = f"{os.path.dirname(__file__)}"
    config_path = f"{pro_path}/conf.ini"
    if not os.path.exists(config_path):
        cf.add_section('FCS')
        cf.set('FCS', '接口地址', 'http://192.168.80.116/fcsserver/composite/upload')
        cf.set('FCS', '指定文件', '')
        cf.set('FCS', '指定工作表', '')
        cf.set('FCS', '存放文件夹', '')
        cf.write(open('conf.ini', 'w'))
    cf.read(config_path)
    return cf.get(sec, key)


def set_conf(sec, opt, value):
    cf = ConfigParser()
    config_file = f"{project_path}/conf.ini"
    cf.read(config_file)
    cf.set(sec, opt, value)
    cf.write(open(config_file, "w"))


def copy_test_excel(proj, case_excel):
    case_path = os.path.join(proj + '测试用例', case_excel)
    result = case_excel[:-5] + '_result_%s.xlsx' % strftime('%Y_%m_%d_%H_%M_%S')  # 结果表命名，附加开始测试时间的后缀
    result_path = os.path.join(proj + '测试结果', result)
    copyfile(case_path, result_path)
    return result_path


def get_case_sheet_name(folder, case_excel):
    sheet_names = []
    book = load_workbook(os.path.join(folder, case_excel), read_only=True)
    for sheet_name in book.sheetnames:
        sheet_names.append(sheet_name)
    return sheet_names


def get_sheet_name(fcs_result_p):
    sheet_names = []
    book = load_workbook(os.path.join(fcs_result_p), read_only=True)
    for sheet_name in book.sheetnames:
        sheet_names.append(sheet_name)
    return sheet_names


def get_target_cases(proj):
    """
    1、如果conf.ini指定文件为空，运行测试用例文件夹下所有xlsx文件
    2、如果conf.ini指定文件想为多个值，用“空格”连接，如：aaa.xlsx bbb.xlsx
    3、如果conf.ini指定文件为一个值，指定工作表为空，运行所有工资表
    4、如果conf.ini指定文件为一个值，指定工作表为想为多个值，如：sheet1 sheet2
    5、如果conf.ini指定文件为空，指定工作表为多个值，此时是无效的
    概括：空时，运行所有的，多个值就用“空格”连接
    :return: {'aaa.xlsx': ["sheet1","sheet2"], 'bbb.xlsx':["sheet1","sheet2"]}
    """
    target_cases = {}  # 目标测试用例的字典
    all_case_excel = get_conf(proj, '指定文件')  # 获得指定文件下的数据
    if all_case_excel == '':  # 如果指定为空
        for ss in get_all_xlsx_file(proj + '测试用例'):  # 空的话就是运行所有，遍历测试用例文件夹
            target_cases[ss] = get_case_sheet_name(proj + '测试用例',
                                                   ss)  # 字典格式{"表格1",["sheet1","sheet2"],"表格2",["sheet1","sheet2"]}
    else:  # 指定文件不为空时
        all_case_excel_list = all_case_excel.split()  # 拆分指定文件，如果有多个值，就拆分为多个值组成的列表
        if len(all_case_excel_list) == 1:  # 如果只有一个值时，指定工作表生效
            if get_conf(proj, '指定工作表') == '':  # 如果指定工作表为空时
                target_cases[all_case_excel] = get_case_sheet_name(proj + '测试用例', all_case_excel)  # 获得该excel表下所有工作表
            else:  # 指定工作表不为空时
                # target_cases[all_case_excel] = get_case_sheet_name(proj + '测试用例', get_conf(proj, '指定工作表').split())  # 不为空为按照指定的来写入字典
                target_cases[all_case_excel] = get_conf(proj, '指定工作表').split()  # 不为空为按照指定的来写入字典
        else:  # 指定为多个值时
            for ss in all_case_excel_list:  # 遍历指定的多个表格
                target_cases[ss] = get_case_sheet_name(proj + '测试用例', ss)  # 获取指定的表格的工作表
    return target_cases


def get_cell(file_path, row_num, column_num, sheet_name):
    workbook = load_workbook(file_path)
    sheet = workbook[sheet_name]
    cell_value = sheet.cell(row=row_num, column=column_num).value
    return cell_value


def get_max_row(file_path, sheet_name):
    workbook = load_workbook(file_path)
    sheet = workbook[sheet_name]
    max_num = sheet.max_row
    return max_num


def get_result(file_path, row_num, sheet_name):
    workbook = load_workbook(file_path)
    sheet = workbook[sheet_name]
    result = sheet.cell(row=row_num, column=7).value
    return result


def set_cell(file_path, row_num, value, sheet_name):
    workbook = load_workbook(file_path)
    sheet = workbook[sheet_name]
    sheet.cell(row=row_num, column=12, value=value)
    workbook.save(file_path)


def image_contrast(img1, img2):
    """用于对比两张截图的差异，当结果等于0时，两张图相同，结果越大，差异越大"""
    image1 = Image.open(img1)
    image2 = Image.open(img2)
    h1 = image1.histogram()
    h2 = image2.histogram()
    result = math.sqrt(reduce(operator.add, list(map(lambda a, b: (a - b) ** 2, h1, h2))) / len(h1))  # 算法
    return result


def get_project_path():  # 获取当前项目的路径
    return os.getcwd()


def get_download_path():  # 文件下载地址的储存
    download_path = get_project_path() + '\\download'
    return download_path


def del_download_file(path=get_download_path()):  # 直接删除该文件夹，然后再去新建该文件夹就行了
    try:
        shutil.rmtree(path)
    except PermissionError:
        pass
    os.mkdir(path)


def download_file_exist(filename):
    """
    判断项目download文件夹中是否存在被下载的的文件
    :param filename: 文件名
    :return: flag 1.ture表示存在  2.false表示不存在
    """
    if os.path.exists(get_download_path() + '\\' + filename):
        return True
    else:
        return False


def assert_pic_exist_or_not(pic_path):
    """
    判断图片路径
    :param pic_path: 文件路径
    :return: flag 1.ture表示存在  2.false表示不存在
    """
    flag = os.path.lexists(pic_path)
    return flag


def get_file_size(file_path):
    """
    获取文件大小，转换为M
    :param file_path: 文件路径
    :return: 返回文件大小，以M为单位
    """
    stat_info = os.stat(file_path)
    size = stat_info.st_size / 1024 / 1024  # 获取文件的size，单位为M
    return size

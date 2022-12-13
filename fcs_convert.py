# coding=utf-8
from main import Test
from tools import *
import traceback
from time import sleep


def convert_run():
    """接口转换"""
    try:
        for xlsx, sheets in get_target_cases('FCS').items():
            result_excel = copy_test_excel('FCS', xlsx)
            set_conf('FCS', '测试结果文件夹', result_excel)
            print('--------------------------正在运行%s--------------------------------------' % xlsx)
            for sheet in sheets:
                Test('FCS', result_excel, sheet).run()
    except:
        print(traceback.format_exc())
        critical(traceback.format_exc())


def convert_rerun():
    """接口转换失败的用例重跑"""
    try:
        rerun_xlsx = get_conf('FCS', '测试结果文件夹')
        print('------------------------正在重跑接口转换失败的用例%s------------------------------------' % rerun_xlsx)
        for sheet in get_sheet_name(rerun_xlsx):
            Test('FCS', rerun_xlsx, sheet).run()
    except:
        print(traceback.format_exc())
        critical(traceback.format_exc())


if __name__ == '__main__':
    convert_run()
    sleep(5)
    convert_rerun()

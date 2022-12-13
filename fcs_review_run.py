# coding=utf-8
from main import BrowserAction, pro_path
from time import sleep
from main import Test
from tools import *
import traceback


def convert_run():
    """参数接口转换"""
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


def get_except_screen_shoot():
    """打开预览链接后，进行期望截图"""
    for sheet_name in sheet_names:
        for i in range(2, get_max_row(fcs_result_path, sheet_name) + 1):
            case_id = get_cell(fcs_result_path, i, 1, sheet_name)
            if get_result(fcs_result_path, i, sheet_name) == "通过":
                url = get_cell(fcs_result_path, i, 8, sheet_name)
                driver = BrowserAction()
                driver.open_bro(url)
                sleep(2)
                file_name = driver.image_path_preview(fcs_result_path, i, 10, sheet_name, capture_type='期望截图')
                print(
                    f'--------------------------正在运行{sheet_name}下用例:{case_id}的期望截图---------------------------')
                data = eval(get_cell(fcs_result_path, i, 6, sheet_name))  # 获取返回的body
                src_filename = data["data"]["srcFileName"]  # 获取文件名
                suffix = src_filename.split(".")[1]  # 去掉文件名，只留后缀
                if suffix in ["zip", "rar", "7z", "gz", "tar"]:  # 如果是压缩包格式
                    driver.click_ele("//*[@class='fcon ico_file_dir']")  # 先点击压缩包按钮
                    num = case_id.split("_")[1]
                    driver.click_ele(f"//*[@id='files']/a[{num}]")
                    sleep(5)
                    driver.capture_image_function(file_name)  # 打开压缩文件后截图
                    driver.click_ele("//*[@class='pica-close ac-pica-close']")  # 点击右上角关闭按钮，返回压缩文件列表页面
                else:  # 如果是其他格式的文件，不根据元素直接截图
                    driver.capture_image_function(file_name)
            else:
                print(f"{case_id}接口转换失败,不执行截图")


def result_run():
    """进行实际截图，并进行对比"""
    for sheet_name in sheet_names:
        for i in range(2, get_max_row(fcs_result_path, sheet_name) + 1):
            case_id = get_cell(fcs_result_path, i, 1, sheet_name)
            if get_result(fcs_result_path, i, sheet_name) == "通过":
                url = get_cell(fcs_result_path, i, 8, sheet_name)
                driver = BrowserAction()
                driver.open_bro(url)
                sleep(2)
                file_name = driver.image_path_preview(fcs_result_path, i, 11, sheet_name, capture_type='实际截图')
                data = eval(get_cell(fcs_result_path, i, 6, sheet_name))  # 获取返回的body
                src_filename = data["data"]["srcFileName"]  # 获取文件名
                suffix = src_filename.split(".")[1]  # 去掉文件名，只留后缀
                if suffix in ["zip", "rar", "7z", "gz", "tar"]:  # 如果是压缩包格式
                    driver.click_ele("//*[@class='fcon ico_file_dir']")  # 先点击压缩包按钮
                    num = case_id.split("_")[1]
                    driver.click_ele(f"//*[@id='files']/a[{num}]")
                    sleep(5)
                    driver.capture_image_function(file_name)  # 打开压缩文件后截图
                    driver.click_ele("//*[@class='pica-close ac-pica-close']")  # 点击右上角关闭按钮，返回压缩文件列表页面
                else:  # 如果是其他格式的文件，不根据元素直接截图
                    driver.capture_image_function(file_name)
                # 图片对比
                except_image_path = pro_path + "/期望截图/" + case_id + ".png"
                actual_image_path = pro_path + "/实际截图/" + case_id + ".png"
                result = image_contrast(except_image_path, actual_image_path)
                print(
                    f'--------------------------正在进行{sheet_name}下用例:{case_id}截图对比---------------------------')
                print(f"对比值为{result}")
                set_cell(fcs_result_path, i, result, sheet_name)
            else:
                print(f"{case_id}接口转换失败,不执行截图对比")


if __name__ == '__main__':
    convert_run()  # 接口转换
    sleep(5)
    convert_rerun()  # 接口转换失败的用例重跑
    sleep(5)
    fcs_result_path = get_conf('FCS', '测试结果文件夹')
    sheet_names = get_sheet_name(fcs_result_path)  # 获得该excel表下所有工作表
    # get_except_screen_shoot()#期望截图
    # sleep(5)
    result_run()  # 实际截图并对比

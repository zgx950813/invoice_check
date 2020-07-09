#!/usr/bin/python
# -*- coding: UTF-8 -*-

import time
import sys
import os
import datetime

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from invoice import Invoice, Result


def isElementPresent(driver, by, value):
    from selenium.common.exceptions import NoSuchElementException
    try:
        driver.find_element(by=by, value=value)
    except NoSuchElementException as e:
        print(e)
        # 发生了NoSuchElementException异常，说明页面中未找到该元素
        return False
    else:
        # 没有发生异常，表示在页面中找到了该元素
        return True


def init_driver(driver_path):
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--incognito')        # 隐身模式（无痕模式）
    chrome_options.add_argument('--start-maximized')  # 设置浏览器分辨率（全屏）
    driver = webdriver.Chrome(
        executable_path=driver_path, chrome_options=chrome_options)
    driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
        'source': 'Object.defineProperty(navigator, "webdriver", {get: () => undefined})'
    })  # 查票平台会根据js字段进行屏蔽，添加此句进行反屏蔽
    print('[INFO] 环境初始化完成')
    return driver


def input_invoice_info(driver, invoice):
    # get方法，打开指定网址
    driver.get('https://inv-veri.chinatax.gov.cn/index.html')

    # 选择网页元素
    driver.find_element_by_id('fpdm').send_keys(invoice.invoice_code)  # 发票代码
    driver.find_element_by_id('fphm').send_keys(invoice.invoice_num)   # 发票号码
    driver.find_element_by_id('kprq').send_keys(invoice.date_issued)   # 开票日期
    driver.find_element_by_id('kjje').send_keys(invoice.check_price)   # 开具金额
    # 利用js将为元素设置焦点
    driver.execute_script(
        "arguments[0].focus();", driver.find_element_by_id('yzm'))     # 验证码


def handle_response(driver, invoice):
    # 截图路径
    project_path = os.path.dirname(os.path.abspath(__file__))
    file_path = project_path + '\\ScreenShot\\' + invoice.vin + '.png'

    # 所有可能结果是 1. 查询次数超过一天5次；2. 查询失败，查无此票； 3. 查询成功
    # 1. 查询次数超过限制
    if isElementPresent(driver, 'id', 'popup_message'):
        popup_message = driver.find_element_by_id('popup_message')
        print(popup_message.text)
        result = Result(check_time=datetime.datetime.now().strftime("%Y-%m-%d %X"),
                        invoice_status='超过可查次数', invoice_code='无', invoice_num='无', date_issued='无', check_price='无')
    else:
        driver.switch_to.frame("dialog-body")
        cysj = driver.find_element_by_id('cysj')  # 查验时间

        # 2. 查无此票/不一致
        if isElementPresent(driver, 'id', 'cyjg'):
            cyjg = driver.find_element_by_id('cyjg')    # 查验结果
            fp_dm = driver.find_element_by_id('fp_dm')  # 发票代码
            fp_hm = driver.find_element_by_id('fp_hm')  # 发票号码
            kp_rq = driver.find_element_by_id('kp_rq')  # 开票日期
            kj_je = driver.find_element_by_id('kj_je')  # 不含税价
            result = Result(check_time=cysj.text, invoice_status=cyjg.text,
                            invoice_code=fp_dm.text, invoice_num=fp_hm.text, date_issued=kp_rq.text, check_price=kj_je.text)

        # 3. 查询成功
        elif isElementPresent(driver, 'id', 'cycs'):
            cycs = driver.find_element_by_id('cycs')                # 查验次数
            # 机动车销售统一发票
            if isElementPresent(driver, 'id', 'fpdm_jdcfp'):
                fpdm = driver.find_element_by_id('fpdm_jdcfp')      # 机动车发票代码
                fphm = driver.find_element_by_id('fphm_jdcfp')      # 机动车发票号码
                kprq = driver.find_element_by_id('kprq_jdcfp')      # 机动车开票日期
                cjfy = driver.find_element_by_id('cjfy_jdcfp')      # 机动车不含税价
                cjhm = driver.find_element_by_id('cjhm_jdcfp')      # 机动车车架号码
                ghdw = driver.find_element_by_id('ghdw_jdcfp')      # 机动车买方单位
                sfzhm = driver.find_element_by_id('sfzhm_jdcfp')    # 机动车买方证号
                xhdwmc = driver.find_element_by_id('xhdwmc_jdcfp')  # 机动车卖方单位
                nsrsbh = driver.find_element_by_id('nsrsbh_jdcfp')  # 机动车卖方证号
                result = Result(check_time=cysj.text, invoice_status=cycs.text, invoice_code=fpdm.text,
                                invoice_num=fphm.text, date_issued=kprq.text, check_price=cjfy.text, vin=cjhm.text,
                                buyer_entity=ghdw.text, buyer_identity=sfzhm.text, seller_entity=xhdwmc.text, seller_identity=nsrsbh.text)

            # 二手车销售统一发票
            elif isElementPresent(driver, 'id', 'fpdm_escfp'):
                fpdm = driver.find_element_by_id('fpdm_escfp')      # 二手车发票代码
                fphm = driver.find_element_by_id('fphm_escfp')      # 二手车发票号码
                kprq = driver.find_element_by_id('kprq_escfp')      # 二手车开票日期
                cjfy = driver.find_element_by_id('cjhjxx_escfp')    # 二手车车价合计
                cjhm = driver.find_element_by_id('clsbdm_escfp')    # 二手车车架号码
                ghdw = driver.find_element_by_id('mfmc_escfp')      # 二手车买方单位
                sfzhm = driver.find_element_by_id('mfdm_escfp')     # 二手车买方证号
                xhdwmc = driver.find_element_by_id('xfmc_escfp')    # 二手车卖方单位
                nsrsbh = driver.find_element_by_id('xfdm_escfp')    # 二手车卖方证号
                result = Result(check_time=cysj.text, invoice_status=cycs.text, invoice_code=fpdm.text,
                                invoice_num=fphm.text, date_issued=kprq.text, check_price=cjfy.text, vin=cjhm.text,
                                buyer_entity=ghdw.text, buyer_identity=sfzhm.text, seller_entity=xhdwmc.text, seller_identity=nsrsbh.text)

    driver.get_screenshot_as_file(file_path)
    print("[INFO] 截图成功，保存于%s" % (file_path))
    return result


if __name__ == '__main__':
    invoice = Invoice(vin='L2CAB3B2XKG102680',
                      invoice_code='143001720660',
                      invoice_num='01741862',
                      date_issued='20180327',
                      check_price='264786.32')
    # 初始化环境
    driver = init_driver('chromedriver.exe')
    # 输入发票信息
    input_invoice_info(driver, invoice)
    # 抓取键盘输入
    while True:
        key_input = input()
        if key_input == 'ok' or key_input == 'OK':
            print('[INFO] 程序继续执行')
            break
        elif key_input == 'exit' or key_input == 'EXIT':
            print('[INFO] 手动停止执行')
            driver.quit()
            exit(0)
    # 处理返回
    result = handle_response(driver, invoice)
    result.print_info()
    # 退出浏览器
    driver.quit()
    print('[INFO] 欢迎再次使用！Kimmy (✿◡‿◡)')

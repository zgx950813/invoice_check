#!/usr/bin/python
# -*- coding: UTF-8 -*-


class Invoice:
    def __init__(self, vin, invoice_code, invoice_num, date_issued, check_price):
        self.vin = vin  # 命名截图文件用
        self.invoice_code = invoice_code
        self.invoice_num = invoice_num
        self.date_issued = date_issued
        self.check_price = check_price

    def print_info(self):
        print('\n=========输入信息开始==========')
        print('车架号码: ', self.vin)
        print('发票代码: ', self.invoice_code)
        print('发票号码: ', self.invoice_num)
        print('开票日期: ', self.date_issued)
        print('查询价格: ', self.check_price)
        print('=========输入信息结束==========\n')


class Result:
    def __init__(self, check_time, invoice_status, invoice_code, invoice_num,
                 date_issued, check_price, vin='无', buyer_entity='无', buyer_identity='无', seller_entity='无', seller_identity='无'):
        self.check_time = check_time
        self.invoice_status = invoice_status
        self.invoice_code = invoice_code
        self.invoice_num = invoice_num
        self.date_issued = date_issued
        self.check_price = check_price
        self.vin = vin
        self.buyer_entity = buyer_entity
        self.buyer_identity = buyer_identity
        self.seller_entity = seller_entity
        self.seller_identity = seller_identity

    def print_info(self):
        print('\n=========输出信息开始==========')
        print('查询时间: ', self.check_time)
        print('结果状态: ', self.invoice_status)
        print('发票代码: ', self.invoice_code)
        print('发票号码: ', self.invoice_num)
        print('开票日期: ', self.date_issued)
        print('查询价格: ', self.check_price)
        print('车架号码: ', self.vin)
        print('买方单位: ', self.buyer_entity)
        print('买方证号: ', self.buyer_identity)
        print('卖方单位: ', self.seller_entity)
        print('卖方证号: ', self.seller_identity)
        print('=========输出信息结束==========\n')

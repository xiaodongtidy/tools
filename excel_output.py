#!/usr/bin/env python
# -*- coding:utf-8 -*-
import io
from xlwt import *


class BaseExcelOutput:
    """
    基础单表数据导出
    """

    def run(self, data: [dict, dict], format_data: dict) -> bytes:
        """
        :param data:        example——[
            {'number': '123', 'status': 1, 'is_pay': False},
            {'number': '124', 'status': 3, 'is_pay': True}
        ]
        :param format_data: example——{
            'title_data': [['单号', 5918], ['状态', 3000], ['是否结算', 3400]],
            'format': {'status': ['', '待付款', '待发货', '已完成', '已撤单'], 'is_pay': ['否', '是']}
        }
        ①title_data中“单号”、“状态”、“是否结算”和data中列表中字典的键按顺序一一对应；
        ②5918等数字表示当前列宽；
        ③4个中文字符取3400左右的值即可；
        ④format中“status”可以将格式为int的id转换成对应的中文选项，
        本例的选项为((1, 待付款), (2, 待发货), (3, 已完成), (4, 已撤单))。“is_pay”可以将Boolean数据格式转换成是或否；
        ⑤title_data、format为固定格式，当format为空时设置成空字典“{}”即可；
        :return:            excel文件二进制流，可保存为.xls格式
        """
        # 开始制表
        ws = Workbook(encoding="utf-8")
        style = XFStyle()  # 全局初始化样式
        title_style = XFStyle()  # 标题样式
        # 对齐样式
        al = Alignment()
        al.horz = 0x02  # 设置水平居中
        al.vert = 0x01  # 设置垂直居中
        style.alignment = al
        title_style.alignment = al
        # 加粗字体样式
        font = Font()
        font.name = '微软雅黑'
        font.height = 20 * 12
        font.bold = True
        title_style.font = font
        # 普通字体样式
        font = Font()
        font.name = '微软雅黑'
        font.bold = False
        style.font = font
        w = ws.add_sheet('sheet1', cell_overwrite_ok=True)
        self.excel_format(w, style, title_style, data, format_data)
        bio = io.BytesIO()
        ws.save(bio)
        bio.seek(0)
        return bio.getvalue()

    @staticmethod
    def excel_format(w, style, title_style, data, format_data):
        if len(data) == 0:
            return
        title_data = format_data['title_data']
        for index, title in enumerate(title_data):
            w.write(0, index, title[0], title_style)
        index = 1
        for purchase_order in data:
            for col_index, info in enumerate(purchase_order):
                if info in format_data['format']:
                    w.write(index, col_index, format_data['format'][info][purchase_order[info]], style)
                else:
                    w.write(index, col_index, purchase_order[info], style)
            index += 1
        for index, title in enumerate(title_data):
            w.col(index).width = title[1]


class DetailExcelOutput(BaseExcelOutput):
    """
    订单+详情类数据的导出
    data: [
        {'number': '1', 'status': 1, 'is_pay': False, 'details': [{'名称': '门', '数量': 2}, {'名称': '窗', '数量': 4}]},
        {'number': '2', 'status': 3, 'is_pay': True, 'details': [{'名称': '墙纸', '数量': 2}, {'名称': '龙头', '数量': 3}]}
    ]
    format_data: {
        'title_data': [['单号', 5918], ['状态', 3000], ['是否结算', 3400], ['名称', 4000], ['数量', 3000]],
        'format': {'status': ['', '待付款', '待发货', '已完成', '已撤单'], 'is_pay': ['否', '是']}
    }
    title_data中“单号”、“状态”、“是否结算”和data列表中字典的键按顺序一一对应，
    后面的“名称”、“数量”和details列表中字典的键按顺序一一对应。
    """

    @staticmethod
    def excel_format(w, style, title_style, data, format_data):
        if len(data) == 0:
            return
        information, details = len(data[0]) - 1, len(data[0]['details'][0])
        w.write_merge(0, 0, 0, information - 1, '单据信息', title_style)
        w.write_merge(0, 0, information, information - 1 + details, '单据详情', title_style)
        title_data = format_data['title_data']
        for index, title in enumerate(title_data):
            w.write(1, index, title[0], title_style)
        index = 2
        for purchase_order in data:
            start = index
            if purchase_order['details']:
                for detail in purchase_order['details']:
                    for col_index, key in enumerate(detail):
                        w.write(index, col_index + information, detail[key], style)
                    index += 1
            else:
                index += 1
            del purchase_order['details']
            for col_index, info in enumerate(purchase_order):
                if info in format_data['format']:
                    w.write_merge(
                        start, index - 1, col_index, col_index, format_data['format'][info][purchase_order[info]], style
                    )
                else:
                    w.write_merge(start, index - 1, col_index, col_index, purchase_order[info], style)
        for index, title in enumerate(title_data):
            w.col(index).width = title[1]


with open('a.xls', 'wb') as fp:
    a = DetailExcelOutput()
    data_ = [
        {'number': '1', 'status': 1, 'is_pay': False, 'details': [{'名称': '门', '数量': 2}, {'名称': '窗', '数量': 4}]},
        {'number': '2', 'status': 3, 'is_pay': True, 'details': [{'名称': '墙纸', '数量': 2}, {'名称': '龙头', '数量': 3}]}
    ]
    format_data_ = {
        'title_data': [['单号', 5918], ['状态', 3000], ['是否结算', 3400], ['名称', 4000], ['数量', 3000]],
        'format': {'status': ['', '待付款', '待发货', '已完成', '已撤单'], 'is_pay': ['否', '是']}
    }
    s = a.run(data=data_, format_data=format_data_)
    fp.write(s)

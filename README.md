环境依赖：
python3.6
pip install xlwt

小试牛刀：
```python
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
```
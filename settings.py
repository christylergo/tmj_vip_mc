# 设置运行的默认参数
import win32con
import win32api
import datetime


# 获取桌面路径
def get_desktop():
    key = win32api.RegOpenKey(
        win32con.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders', 0,
        win32con.KEY_READ)
    return win32api.RegQueryValueEx(key, 'Desktop')[0]


# 占位符,用于列簇层级结构
placeholder = None

# 表格生成后是否打开, True表示'是',False表示'否'
SHOW_DOC_AFTER_GENERATED = True

REAL_VIRTUAL = {
    'virtual_inventory': [0, True, placeholder],
    'real_inventory': [1, True, placeholder],
}

# 各仓显示优先级,数字越小显示越靠前, True表示显示此列，False表示不显示
WAREHOUSE_PRIORITY = {
    'HanChuan': [1, True, REAL_VIRTUAL],  # 汉川仓
    'ChengDong': [2, True, REAL_VIRTUAL],  # 城东仓
    'LingDing': [3, True, REAL_VIRTUAL],  # 岭顶仓
    'YueZhong': [4, True, REAL_VIRTUAL],  # 越中小件仓
    'LinDa': [5, True, REAL_VIRTUAL],  # 琳达仓
    'PiFa': [6, True, REAL_VIRTUAL],  # 五夫批发仓
    'KunShan': [7, True, REAL_VIRTUAL],  # 昆山仓
    'adjustment': [0, False, placeholder],  # 库存修正，不需要显示，但是数据优先级最高
}

daily_sales_week_title = [f'{datetime.date.today()- datetime.timedelta(days=i):%m/%d}' for i in range(7, 0, -1)]
daily_sales_week_priority = [[i, True, placeholder] for i in range(1, 8)]
# 各属性显示优先级,数字越小显示越靠前, True表示显示此列，False表示不显示
FEATURE_PRIORITY = {
    'nu': [0, True, placeholder],  # 序号
    'platform': [1, True, placeholder],  # 在售平台
    'status': [2, True, placeholder],  # 上下架状态
    'vip_barcode': [3, True, placeholder],  # 唯品条码
    'vip_commodity': [4, False, placeholder],  # 唯品款号
    'vip_item': [5, True, placeholder],  # 唯品货号
    'vip_item_name': [6, True, placeholder],  # 商品名称
    'tmj_barcode': [7, True, placeholder],  # 货品条码明细
    'vip_category': [8, True, placeholder],  # 唯品分类
    'month_sales': [9, True, placeholder],  # 月销量
    'week_sales': [10, True, placeholder],  # 周销量
    'site_DSI': [11, True, placeholder],  # 页面库存周转
    'site_stock': [12, True, placeholder],  # 页面库存
    'disassemble': [13, True, placeholder],  # 组合分解为单品,在此列填写数量重新运行后会生成分解后的单品数量。
    'stock_inventory': [14, True, WAREHOUSE_PRIORITY],  # 仓组库存，一个或多个仓
    'cost': [15, True, placeholder],  # 成本
    'weight': [16, True, placeholder],  # 重量
    'annotation': [17, True, placeholder],  # 备注，包括缺货的货品信息，缺货情况下如果有可以替换的货品也会写在备注里
    # 按天显示最近一周的销量,注意此处是字典生成式(iterable)
    'daily_sales_week': [18, True, zip(daily_sales_week_title, daily_sales_week_priority)],
}


COLUMN_PROPERTY = [
    {'identity': 'nu', 'name': '序号', 'refer_doc': placeholder, 'floating_title': placeholder},
    {'identity': 'platform', 'name': , 'refer_doc': , 'floating_title': },
    {'identity': 'status', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'vip_barcode', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'vip_commodity', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'vip_item', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'vip_item_name', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'tmj_barcode', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'vip_category', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'month_sales', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'week_sales', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'site_DSI', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'site_stock', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'disassemble', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'HanChuan', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'ChengDong', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'LingDing', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'YueZhong', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'LinDa', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'PiFa', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'KunShan', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'adjustment', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'cost', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'weight', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': 'annotation', 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': daily_sales_week_title[0], 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': daily_sales_week_title[1], 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': daily_sales_week_title[2], 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': daily_sales_week_title[3], 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': daily_sales_week_title[4], 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': daily_sales_week_title[5], 'name': , 'refer_doc': , 'floating_title':: },
    {'identity': daily_sales_week_title[6], 'name': , 'refer_doc': , 'floating_title':: },
]
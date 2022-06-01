# -*- coding: utf-8 -*-
import sys
import win32con
import win32api
import datetime


# 表格生成后是否打开, True表示'是',False表示'否'
SHOW_DOC_AFTER_GENERATED = True
# 唯品销量显示的天数,1~30
VIP_SALES_INTERVAL = 7
# 猫超销量的天数,1~30
MC_SALES_INTERVAL = 7
# 占位符,用于列簇层级结构
placeholder = None
# ---------------------文件夹路径(填写在引号内)-------------------------
# 网上导出数据文件夹路径
DOCS_PATH = r'C:\Users\Administrator\Desktop\tmj_vip_mc\vip_docs'
# 代码文件夹路径
sys.path.append(r'C:\Users\Administrator\Desktop\tmj_vip_mc')
# 库存显示方面的设置
warehouses = [
    'HanChuan', 'ChengDong', 'LingDing', 'YueZhong', 'LinDa', 'PiFa', 'KunShan', 'adjustment'
]
warehouses_key_name = ['汉川', '城东', '岭顶', '越中', '琳达', '批发', '昆山', '修正']
# 下面这行列出来的是各个库存文件名称的关键字,用于识别是哪个仓的库存
warehouses_key_re = ['汉川', '城东', '岭顶', '越中', '琳达', '批发', '昆山', '修正']

# 各仓显示优先级,数字越小显示越靠前, True表示显示此列，False表示不显示
WAREHOUSE_PRIORITY = {
    'HanChuan': [1, True],  # 汉川仓
    'ChengDong': [2, True],  # 城东仓
    'LingDing': [3, True],  # 岭顶仓
    'YueZhong': [4, True],  # 越中小件仓
    'LinDa': [5, True],  # 琳达仓
    'PiFa': [6, True],  # 五夫批发仓
    'KunShan': [7, True],  # 昆山仓
    'adjustment': [0, True],  # 库存修正，不需要显示，但是数据优先级最高
}

# 各属性显示优先级,数字越小显示越靠前, True表示显示此列，False表示不显示
FEATURE_PRIORITY = {
    'row_nu': [0, True],  # 序号
    'platform': [1, True],  # 在售平台
    'mc_week_sales': [2, True],  # 猫超周销量
    'status': [3, True],  # 上下架状态
    'vip_barcode': [4, True],  # 唯品条码
    'vip_commodity': [5, False],  # 唯品货号
    'vip_item_name': [6, True],  # 商品名称
    'tmj_barcode': [7, True],  # 货品条码明细
    'vip_category': [8, True],  # 唯品分类
    'month_sales': [9, True],  # 月销量
    'stock_DSI': [10, True],  # 可用库存周转
    'site_DSI': [11, True],  # 页面库存周转
    'site_inventory': [12, True],  # 页面库存
    'disassemble': [13, True],  # 组合分解为单品,在此列填写数量重新运行后会生成分解后的单品数量。
    'WAREHOUSE_PRIORITY': [14, True],  # 仓组库存作为一个集合的优先级，一个或多个仓
    'cost': [15, True],  # 成本
    'weight': [16, True],  # 重量
    'annotation': [17, True],  # 备注，包括缺货的货品信息，缺货情况下如果有可以替换的货品也会写在备注里
    'site_link': [18, True],  # 网页链接
    # 按天显示最近一周的销量,注意此处是字典生成式(iterable)
    'DAILY_SALES_WEEK': [19, True],
}

for value in FEATURE_PRIORITY.values():
    value[0] *= 100

for value in WAREHOUSE_PRIORITY.values():
    value[0] *= 10
    value[0] += FEATURE_PRIORITY['WAREHOUSE_PRIORITY'][0]

daily_sales_week_title = [
    f'{datetime.date.today() - datetime.timedelta(days=i):%m/%d}' for i in range(VIP_SALES_INTERVAL, 0, -1)
]

daily_sales_week_priority = [
    [FEATURE_PRIORITY['DAILY_SALES_WEEK'][0] + i,
     FEATURE_PRIORITY['DAILY_SALES_WEEK'][1]]
    for i in range(0, VIP_SALES_INTERVAL)
]

DAILY_SALES_WEEK = zip(daily_sales_week_title, daily_sales_week_priority)
FEATURE_PRIORITY.update(DAILY_SALES_WEEK)

WAREHOUSE_PRIORITY_REAL_VIRTUAL = ({
    warehouses[i] + '_virtual': [WAREHOUSE_PRIORITY[warehouses[i]][0] + 1, WAREHOUSE_PRIORITY[warehouses[i]][1]],
    warehouses[i]: [WAREHOUSE_PRIORITY[warehouses[i]][0] + 2, WAREHOUSE_PRIORITY[warehouses[i]][1]],
} for i in range(0, len(warehouses))
)

for i in range(0, len(warehouses)):
    FEATURE_PRIORITY.update(next(WAREHOUSE_PRIORITY_REAL_VIRTUAL))

# 定义全部可能会用到的列,用生成式来定义特性一致的列，如库存列以及日销列
COLUMN_PROPERTY = [
    {'identity': 'row_nu', 'name': '序号',
     'refer_doc': 'self', 'floating_title': placeholder},
    {'identity': 'platform', 'name': '在售平台', 'refer_doc': 'arrAtom',
     'floating_title': 'platform'},  # 1唯品，2猫超，3共用
    {'identity': 'status', 'name': '在架状态',
     'refer_doc': 'vip_routine_operation', 'floating_title': '尺码状态'},
    {'identity': 'vip_barcode', 'name': '唯品条码',
     'refer_doc': 'vip_fundamental_collections', 'floating_title': '唯品后台条码'},
    {'identity': 'vip_commodity', 'name': '唯品货号', 'refer_doc': 'vip_fundamental_collections',
     'floating_title': '唯品会货号'},
    {'identity': 'vip_item_name', 'name': '商品名称',
     'refer_doc': 'vip_fundamental_collections', 'floating_title': '商品名称'},
    {'identity': 'tmj_barcode', 'name': '旺店通编码明细',
     'refer_doc': 'arrAtom', 'floating_title': 'tmj_barcode'},
    {'identity': 'vip_category', 'name': '类别',
     'refer_doc': 'vip_fundamental_collections', 'floating_title': '类别'},
    {'identity': 'month_sales', 'name': '月销量',
     'refer_doc': 'vip_daily_sales', 'floating_title': 'month_sales'},
    {'identity': 'mc_week_sales', 'name': '猫超周销量',
     'refer_doc': 'mc_daily_sales', 'floating_title': 'mc_week_sales'},
    {'identity': 'site_DSI', 'name': '',
     'refer_doc': 'self', 'floating_title': placeholder},
    {'identity': 'site_inventory', 'name': '页面库存余量',
     'refer_doc': 'vip_routine_site_stock', 'floating_title': '可扣库存'},
    {'identity': 'disassemble', 'name': '组合分解',
     'refer_doc': 'self', 'floating_title': placeholder},
    {'identity': 'cost', 'name': '成本',
     'refer_doc': 'tmj_atom', 'floating_title': '会员价'},
    {'identity': 'weight', 'name': '重量',
     'refer_doc': 'tmj_atom', 'floating_title': '重量'},
    {'identity': 'annotation', 'name': '备注',
     'refer_doc': 'arrAtom', 'floating_title': 'annotation'},
    {'identity': 'site_link', 'name': '商品链接',
     'refer_doc': 'vip_daily_sales', 'floating_title': '商品链接'},
]

vip_daily_sales_columns = [{
    'identity': daily_sales_week_title[i], 'name': daily_sales_week_title[i],
    'refer_doc': 'vip_daily_sales', 'floating_title': daily_sales_week_title[i]
} for i in range(0, VIP_SALES_INTERVAL)]

warehouses_stock = [{
    'identity': warehouses[i], 'name': warehouses_key_name[i] + '仓库存',
    'refer_doc': warehouses[i].lower() + '_stock', 'floating_title': ['可发库存', '可用库存']
} for i in range(0, len(warehouses))
]

warehouses_stock_virtual = [{
    'identity': warehouses[i] + '_virtual', 'name': warehouses_key_name[i] + '虚拟仓库存',
    'refer_doc': warehouses[i].lower() + '_stock_virtual', 'floating_title': ['可发库存', '可用库存']
} for i in range(0, len(warehouses))
]

COLUMN_PROPERTY.extend(warehouses_stock)
COLUMN_PROPERTY.extend(warehouses_stock_virtual)
COLUMN_PROPERTY.extend(vip_daily_sales_columns)

doc_stock = [{
    'identity': warehouses[i].lower() + '_stock', 'name': '',
    'key_words': '^[^虚拟].*' + warehouses_key_re[i] + '[^虚拟].*$', 'key_pos': ['商家编码', ], 'val_pos': ['可发库存', '可用库存'],
    'val_type': ['INT', 'INT'],
    'importance': 'optional', 'mode': None,
} for i in range(0, len(warehouses))
]

doc_stock_virtual = [{
    'identity': warehouses[i].lower() + '_stock_virtual', 'name': '',
    'key_words': warehouses_key_re[i] + '.*虚拟|虚拟.*' + warehouses_key_re[i],
    'key_pos': ['商家编码', ], 'val_pos': ['可发库存', '可用库存'],
    'val_type': ['INT', 'INT'],
    'importance': 'optional', 'mode': None,
} for i in range(0, len(warehouses))
]

# 文件重要性的程度分为三类,'required'是必须的,'caution'是不必须,缺少的情况下会提示,'optional'是可选
DOC_REFERENCE = [
    {
        'identity': 'vip_routine_site_stock', 'name': '',  # 页面库存文件
        'key_words': '常态可扣减', 'key_pos': ['条码', ], 'val_pos': ['可扣库存', ], 'val_type': ['INT', ],
        'importance': 'caution', 'mode': None,
    },
    {
        'identity': 'vip_routine_operation', 'name': '',  # 上下架状态文件
        'key_words': '常态商品运营', 'key_pos': ['条码', ], 'val_pos': ['尺码状态', ], 'val_type': ['TEXT', ],
        'importance': 'caution', 'mode': None,
    },
    {
        'identity': 'vip_daily_sales', 'name': '',  # 日销量、商品链接
        'key_words': '条码粒度', 'key_pos': ['条码', '日期', ], 'val_pos': ['销售量', '商品链接', ],
        'val_type': ['INT', 'TEXT', ],
        'importance': 'caution', 'mode': 'merge',
    },
    {
        'identity': 'vip_fundamental_collections', 'name': '',  # 唯品总货表
        'key_words': '唯品会十月总货表', 'key_pos': ['唯品后台条码', '旺店通条码'],
        'val_pos': ['类别', '商品名称', '唯品会货号', '日常券', '自主分类'], 'val_type': ['TEXT', 'TEXT', 'TEXT', 'TEXT', 'TEXT'],
        'importance': 'required', 'mode': None,
    },
    {
        'identity': 'tmj_atom', 'name': '',  # 单品信息表，包含名称、货号、成本、重量等信息
        'key_words': '单品列表', 'key_pos': ['商家编码'], 'val_pos': ['货品编号', '货品名称', '规格名称', '会员价', '重量'],
        'val_type': ['TEXT', 'TEXT', 'TEXT', 'REAL', 'REAL'],
        'importance': 'required', 'mode': None,
    },
    {
        'identity': 'tmj_combination', 'name': '',  # 组合表
        'key_words': '组合装明细', 'key_pos': ['商家编码', ], 'val_pos': ['单品名称', '单品货品编号', '单品商家编码', '数量'],
        'val_type': ['TEXT', 'TEXT', 'TEXT', 'INT'],
        'importance': 'required', 'mode': None,
    },
    {
        'identity': 'mc_item', 'name': '',
        'key_words': 'export-', 'key_pos': ['货品编码', '条码'], 'val_pos': ['采购负责人', '上下架状态'], 'val_type': ['TEXT', 'TEXT'],
        'importance': 'caution', 'mode': None,
    },
    {
        'identity': 'mc_daily_sales', 'name': '',
        'key_words': '业务库存出入库流水', 'key_pos': ['货品ID', '出入库时间', '业务类型'], 'val_pos': ['库存变动'], 'val_type': ['INT'],
        'importance': 'caution', 'mode': 'merge',
    },
    {
        'identity': 'vip_bench_player', 'name': '',
        'key_words': '替换|替代|备选', 'key_pos': ['商家编码_首发', ],
        'val_pos': [
            '商家编码_替代1', '商家编码_替代2', '商家编码_替代3', '商家编码_替代4',
            '交换比_替代1', '交换比_替代2', '交换比_替代3', '交换比_替代4'],  # 交换比是指一个替代品可以替换多少个首发商品
        'val_type': ['TEXT', 'TEXT', 'TEXT', 'TEXT', 'REAL', 'REAL', 'REAL', 'REAL', ],
        'importance': 'optional', 'mode': None,
    },
    {
        'identity': 'vip_summary', 'name': '',  # 生成的统计最终表,当需要分解组合的时候读取.
        'key_words': '唯品库存统计分析', 'key_pos': ['唯品条码'], 'val_pos': ['组合分解'], 'val_type': ['INT'],
        'importance': 'optional', 'mode': None,
    },
]

DOC_REFERENCE.extend(doc_stock)
DOC_REFERENCE.extend(doc_stock_virtual)

print('settings->tracing...')


# 获取桌面路径
def get_desktop() -> str:
    desktop_key = win32api.RegOpenKey(
        win32con.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders', 0,
        win32con.KEY_READ)
    return win32api.RegQueryValueEx(desktop_key, 'Desktop')[0]

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

warehouses = ['HanChuan', 'ChengDong', 'LingDing', 'YueZhong', 'LinDa', 'PiFa', 'KunShan', 'adjustment']
warehouses_key_name = ['汉川', '城东', '岭顶', '越中', '琳达', '批发', '昆山', '修正']
REAL_VIRTUAL = (
    {
        warehouses[i]+'_virtual': [0, True, placeholder],
        warehouses[i]+'_real': [1, True, placeholder],
    }
    for i in range(0, 7)
)
# 各仓显示优先级,数字越小显示越靠前, True表示显示此列，False表示不显示
WAREHOUSE_PRIORITY = {
    'HanChuan': [1, True, next(REAL_VIRTUAL)],  # 汉川仓
    'ChengDong': [2, True, next(REAL_VIRTUAL)],  # 城东仓
    'LingDing': [3, True, next(REAL_VIRTUAL)],  # 岭顶仓
    'YueZhong': [4, True, next(REAL_VIRTUAL)],  # 越中小件仓
    'LinDa': [5, True, next(REAL_VIRTUAL)],  # 琳达仓
    'PiFa': [6, True, next(REAL_VIRTUAL)],  # 五夫批发仓
    'KunShan': [7, True, next(REAL_VIRTUAL)],  # 昆山仓
    'adjustment': [0, False, placeholder],  # 库存修正，不需要显示，但是数据优先级最高
}

daily_sales_week_title = [f'{datetime.date.today() - datetime.timedelta(days=i):%m/%d}' for i in range(7, 0, -1)]
daily_sales_week_priority = [[i, True, placeholder] for i in range(1, 8)]
# 各属性显示优先级,数字越小显示越靠前, True表示显示此列，False表示不显示
FEATURE_PRIORITY = {
    'nu': [0, True, placeholder],  # 序号
    'platform': [1, True, placeholder],  # 在售平台
    'status': [2, True, placeholder],  # 上下架状态
    'vip_barcode': [3, True, placeholder],  # 唯品条码
    'vip_item_name': [4, True, placeholder],  # 商品名称
    'tmj_barcode': [5, True, placeholder],  # 货品条码明细
    'vip_commodity': [6, False, placeholder],  # 唯品货号
    'vip_category': [7, True, placeholder],  # 唯品分类
    'month_sales': [8, True, placeholder],  # 月销量
    'mc_week_sales': [9, True, placeholder],  # 周销量
    'stock_DSI': [10, True, placeholder],  # 可用库存周转
    'site_DSI': [10, True, placeholder],  # 页面库存周转
    'site_inventory': [11, True, placeholder],  # 页面库存
    'disassemble': [12, True, placeholder],  # 组合分解为单品,在此列填写数量重新运行后会生成分解后的单品数量。
    'stock_inventory': [13, True, WAREHOUSE_PRIORITY],  # 仓组库存，一个或多个仓
    'cost': [14, True, placeholder],  # 成本
    'weight': [15, True, placeholder],  # 重量
    'annotation': [16, True, placeholder],  # 备注，包括缺货的货品信息，缺货情况下如果有可以替换的货品也会写在备注里
    'site_link': [17, True, placeholder],  # 网页链接
    # 按天显示最近一周的销量,注意此处是字典生成式(iterable)
    'daily_sales_week': [18, True, zip(daily_sales_week_title, daily_sales_week_priority)],
}

vip_daily_sales_columns = (
    {'identity': daily_sales_week_title[i], 'name': daily_sales_week_title[i],
     'refer_doc': 'vip_daily_sales', 'floating_title': daily_sales_week_title[i]} for i in range(0, 7)
)

warehouses_stock = (
    {'identity': warehouses[i], 'name': warehouses_key_name[i]+'仓库存',
     'refer_doc': warehouses[i].lower()+'_stock', 'floating_title': ['可发库存', '可用库存']}
    for i in range(0, 8)
)
warehouses_stock_virtual = (
    {'identity': warehouses[i]+'_virtual', 'name': warehouses_key_name[i]+'虚拟仓库存',
     'refer_doc': warehouses[i].lower()+'_stock_virtual', 'floating_title': ['可发库存', '可用库存']}
    for i in range(0, 7)
)
COLUMN_PROPERTY = [
    {'identity': 'nu', 'name': '序号', 'refer_doc': placeholder, 'floating_title': placeholder},
    {'identity': 'platform', 'name': '在售平台', 'refer_doc': 'arrAtom', 'floating_title': 'platform'},  # 1唯品，2猫超，3共用
    {'identity': 'status', 'name': '在架状态', 'refer_doc': 'vip_routine_operation', 'floating_title': '尺码状态'},
    {'identity': 'vip_barcode', 'name': '唯品条码', 'refer_doc': 'vip_fundamental_collections', 'floating_title': '唯品后台条码'},
    {'identity': 'vip_commodity', 'name': '唯品货号', 'refer_doc': 'vip_fundamental_collections', 'floating_title': '唯品会货号'},
    {'identity': 'vip_item_name', 'name': '商品名称', 'refer_doc': 'vip_fundamental_collections', 'floating_title': '商品名称'},
    {'identity': 'tmj_barcode', 'name': '旺店通编码明细', 'refer_doc': 'arrAtom', 'floating_title': 'tmj_barcode'},
    {'identity': 'vip_category', 'name': '类别', 'refer_doc': 'vip_fundamental_collections', 'floating_title': '类别'},
    {'identity': 'month_sales', 'name': '月销量', 'refer_doc': 'vip_daily_sales', 'floating_title': 'month_sales'},
    {'identity': 'mc_week_sales', 'name': '猫超周销量', 'refer_doc': 'mc_daily_sales', 'floating_title': 'mc_week_sales'},
    {'identity': 'site_DSI', 'name': '', 'refer_doc': 'self', 'floating_title': ''},
    {'identity': 'site_inventory', 'name': '页面库存余量', 'refer_doc': 'vip_routine_site_stock', 'floating_title': '可扣库存'},
    {'identity': 'disassemble', 'name': '组合分解', 'refer_doc': 'self', 'floating_title': ''},
    next(warehouses_stock), next(warehouses_stock_virtual),
    next(warehouses_stock), next(warehouses_stock_virtual),
    next(warehouses_stock), next(warehouses_stock_virtual),
    next(warehouses_stock), next(warehouses_stock_virtual),
    next(warehouses_stock), next(warehouses_stock_virtual),
    next(warehouses_stock), next(warehouses_stock_virtual),
    next(warehouses_stock), next(warehouses_stock_virtual),
    next(warehouses_stock),
    {'identity': 'cost', 'name': '成本', 'refer_doc': 'tmj_atom', 'floating_title': '会员价'},
    {'identity': 'weight', 'name': '重量', 'refer_doc': 'tmj_atom', 'floating_title': '重量'},
    {'identity': 'annotation', 'name': '备注', 'refer_doc': 'arrAtom', 'floating_title': 'annotation'},
    {'identity': 'site_link', 'name': '商品链接', 'refer_doc': 'vip_daily_sales', 'floating_title': '商品链接'},
    next(vip_daily_sales_columns),
    next(vip_daily_sales_columns),
    next(vip_daily_sales_columns),
    next(vip_daily_sales_columns),
    next(vip_daily_sales_columns),
    next(vip_daily_sales_columns),
    next(vip_daily_sales_columns),
]

DOC_REFERENCE = [
    {'identity': 'vip_routine_site_stock',  # 页面库存文件
     'key_words': ['常态可扣减', '剩余可售库存'], 'key_pos': ['条码', ], 'val_pos': ['可扣库存', ], 'val_type': ['int', ]},

    {'identity': 'vip_routine_operation',  # 上下架状态文件
     'key_words': ['常态商品运营', '导出尺码'], 'key_pos': ['条码', ], 'val_pos': ['尺码状态', ], 'val_type': ['str', ]},

    {'identity': 'vip_daily_sales',  # 日销量、商品链接
     'key_words': ['商品明细', '条码粒度'], 'key_pos': ['条码', '日期', ], 'val_pos': ['销售量', '商品链接', ],
     'val_type': ['int', 'str', ]},

    {'identity': 'vip_fundamental_collections',  # 唯品总货表
     'key_words': ['唯品会十月总货表'], 'key_pos': ['唯品后台条码', '旺店通条码'],
     'val_pos': ['类别', '商品名称', '唯品会货号', '日常券', '自主分类'], 'val_type': ['str', 'str', 'str', 'str', 'str']},

    {'identity': 'tmj_atom',  # 单品信息表，包含名称、货号、成本、重量等信息
     'key_words': ['单品列表'], 'key_pos': ['商家编码'], 'val_pos': ['货品编号', '货品名称', '规格名称', '会员价', '重量'],
     'val_type': ['str', 'str', 'str', 'float', 'float']},

    {'identity': 'tmj_combination',  # 组合表
     'key_words': ['组合装明细'], 'key_pos': ['商家编码', ], 'val_pos': ['单品名称', '单品货品编号', '单品商家编码', '数量'],
     'val_type': ['str', 'str', 'str', 'int']},

    {'identity': 'hanchuan_stock',
     'key_words': ['汉川仓', ], 'key_pos': ['商家编码', ], 'val_pos': ['可发库存', '可用库存'], 'val_type': ['int', 'int']},

    {'identity': 'chengdong_stock',
     'key_words': ['城东仓', ], 'key_pos': ['商家编码', ], 'val_pos': ['可发库存', '可用库存'], 'val_type': ['int', 'int']},

    {'identity': 'lingding_stock',
     'key_words': ['岭顶仓', '领顶仓'], 'key_pos': ['商家编码', ], 'val_pos': ['可发库存', '可用库存'], 'val_type': ['int', 'int']},

    {'identity': 'yuezhong_stock',
     'key_words': ['越中仓', '越中小件仓'], 'key_pos': ['商家编码', ], 'val_pos': ['可发库存', '可用库存'], 'val_type': ['int', 'int']},

    {'identity': 'linda_stock',
     'key_words': ['琳达仓', '琳达妈咪仓'], 'key_pos': ['商家编码', ], 'val_pos': ['可发库存', '可用库存'], 'val_type': ['int', 'int']},

    {'identity': 'pifa_stock',
     'key_words': ['批发仓', '批发专用', '五夫'], 'key_pos': ['商家编码', ], 'val_pos': ['可发库存', '可用库存'],
     'val_type': ['int', 'int']},

    {'identity': 'mc_item',
     'key_words': ['export-', ], 'key_pos': ['货品编码', '条码'], 'val_pos': ['采购负责人', '上下架状态'], 'val_type': ['str', 'str']},

    {'identity': 'mc_daily_sales',
     'key_words': ['业务库存出入库流水', ], 'key_pos': ['货品编码', '出入库时间'], 'val_pos': ['库存变动', ], 'val_type': ['int', ]},

    {'identity': 'adjustment_stock',
     'key_words': ['修正', '库存修正', '在途'], 'key_pos': ['商家编码', ], 'val_pos': ['可发库存', '可用库存'], 'val_type': ['int', 'int']},
    {'identity': 'vip_bench_player',
     'key_words': ['替换', '替代', '备选'], 'key_pos': ['商家编码(首发)', ],
     'val_pos': ['商家编码(替代1)', '商家编码(替代2)', '商家编码(替代3)', '商家编码(替代4)',
                 '交换比(替代1)', '交换比(替代2)', '交换比(替代3)', '交换比(替代4)'],  # 交换比是指一个替代品可以替换多少个首发商品
     'val_type': ['str', 'str', 'str', 'str', 'float', 'float', 'float', 'float', ]},
]

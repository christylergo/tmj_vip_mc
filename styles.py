# -*- coding: utf-8 -*-
import time
import xlwings as xw
import settings as st


def add_styles(dataframe):
    df = dataframe
    wb = xw.Book()
    ws = wb.sheets.active
    ws.name = 'VipInventory'
    old_time = time.time()
    # ------------------------------------
    columns = df.columns.get_level_values(0)
    data_type = df.columns.get_level_values(3)
    number_formats = ['0']
    width = [5]
    text_wrap = [False]
    alignment = [-4108]
    columns_size = df.columns.size
    for ii in range(columns_size):
        for jj in st.COLUMN_PROPERTY:
            if jj['name'] == columns[ii]:
                w = jj.get('width', 6)
                nf = 'General'
                wrap = jj.get('text_wrap', False)
                alm = jj.get('alignment', -4108)
                if data_type[ii] == 'str':
                    nf = '@'
                elif data_type[ii] == 'int':
                    nf = '0'
                number_formats.append(nf)
                width.append(w)
                text_wrap.append(wrap)
                alignment.append(alm)
    a_upper = 65
    font_size = 11
    font = '微软雅黑'
    h_center = -4108
    h_left = -4131
    v_center = -4108
    # 最多循环操作13 次, 尽量避开使用xlwings, openpyxl可以批量设置
    column = ':'.join([chr(a_upper), chr(a_upper + columns_size+1)])
    # column = 'A:B,E:H'
    # ws.range(column).column_width = width[ii]
    ws.range(column).number_format = '@'
    # ws.range(column).wrap_text = text_wrap[ii]
    ws.range(column).font.size = font_size
    ws.range(column).font.name = font
    ws.range(column).api.VerticalAlignment = v_center
    ws.range(column).api.HorizontalAlignment = h_center
    # time.sleep(0.2)
    # ------------------------------------
    ws.range('A1').value = df
    print(time.time() - old_time)
    wb.save(r'C:\Users\Administrator\Downloads\path_to_pandas.xlsx')


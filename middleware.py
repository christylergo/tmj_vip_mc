# -*- coding: utf-8 -*-
import re
import time

import numpy as np
import pandas as pd
import functools
import datetime

import settings as st


class MiddlewareArsenal:
    """
    -*- 容器 -*-
    使用类作为容器放置各种数据处理中间件函数, 函数名称对应doc_reference中的identity.
    然后把函数转换成字典形式保存, 需要处理数据时直接用identity来索引调用.
    注: 不需要实例化此类, 调用实例方法时, self指定为None即可
    中间件函数的形参统一命名为: data_ins
    data_ins = {'identity': self.identity, 'doc_ref': self.doc_ref, 'data_frame': data_frame,
    'to_sql_df': sql_df, 'mode': self.from_sql}
    """

    @staticmethod
    def __rectify_daily_sales(data_ins):
        # 这个日期格式的针对性操作应该放进middleware里, 之前放在reading_docs里
        sales_date_head = datetime.date.today() - datetime.timedelta(days=st.VIP_SALES_INTERVAL)
        sales_date_tail = datetime.date.today()
        date_col = data_ins['doc_ref']['key_pos'][1]
        data_frame = data_ins['data_frame']
        data_frame[date_col] = pd.to_datetime(data_frame[date_col]).dt.date
        data_frame = data_frame.loc[lambda df: df[date_col] >= sales_date_head, :]
        data_frame = data_frame[data_frame[date_col] < sales_date_tail]
        # ---------------------------------------------
        data_frame[date_col] = data_frame[date_col].astype('str')
        # -----------------------------------------------
        # 如果不先判断就进行筛选, 可能会报错
        data_frame = data_frame.sort_index(level=0, kind='mergesort')
        source = data_frame.index.get_level_values(0)
        sql_df = data_frame.loc['sql_df'] if 'sql_df' in source else None
        doc_df = data_frame.loc['doc_df'] if 'doc_df' in source else None
        if sql_df is not None and doc_df is not None:
            sql_date = sql_df.drop_duplicates(
                subset=[date_col], keep='first')[date_col]
            # pandas不要使用 for x in df的形式, 效率很低
            # mask = [False if x in list(
            #     sql_date) else True for x in doc_df[date_col]]
            date_list = sql_date.to_list()
            # 在numpy中扩展, 这样也是可行的
            mask = ~doc_df[date_col].isin(date_list).to_numpy()
            mask = np.hstack([mask, np.array([True] * sql_df.index.size)])
            # 默认ascending=True, 默认使用quicksort, 稳定算法要选mergesort
            data_frame = data_frame[mask]
            source = data_frame.index.get_level_values(0)
            doc_df = data_frame.loc['doc_df'] if 'doc_df' in source else None
        to_sql_df = None
        if doc_df is not None:
            # 不需要reset index, 这个操作很耗时
            # to_sql_df = doc_df.reset_index(drop=True)
            to_sql_df = doc_df
        data_frame = data_frame
        return data_frame, to_sql_df

    # --------------------------------------------------
    @staticmethod
    def __pivot_daily_sales(data_ins):
        key_col = data_ins['doc_ref']['key_pos'][0]
        date_col = data_ins['doc_ref']['key_pos'][1]
        sales_col = data_ins['doc_ref']['val_pos'][0]
        data_frame = data_ins['data_frame']
        sales_date_head = datetime.datetime.today() - datetime.timedelta(days=st.VIP_SALES_INTERVAL)
        data_frame = data_frame.loc[
                     lambda df: pd.to_datetime(df[date_col]) >= sales_date_head, :]
        data_frame = pd.pivot_table(
            data_frame, index=[key_col], columns=[date_col], values=sales_col,
            aggfunc=np.sum, fill_value=0)  # 自定义agg func很便捷但是会严重降低运行速度, 所以尽量使用np.sum .mean等原生函数方法
        data_frame = data_frame.applymap(func=abs)
        data_frame.columns = data_frame.columns.map(lambda x: f"{pd.to_datetime(x):%m/%d}")
        data_frame = data_frame.reset_index()
        return data_frame

    # ---------------------------------------------
    def mc_daily_sales(self, data_ins) -> None:
        while self is not None:
            print('eliminate the weak warnings')
        origin_df, to_sql_df = MiddlewareArsenal.__rectify_daily_sales(data_ins)
        criteria_col = data_ins['doc_ref']['key_pos'][2]
        criterion = pd.concat(
            [origin_df[criteria_col] == 'SO0', origin_df[criteria_col] == 'SO4'], axis=1).any(axis=1)
        data_ins['data_frame'] = origin_df[criterion]
        pivoted_df = MiddlewareArsenal.__pivot_daily_sales(data_ins)
        data_ins['data_frame'] = pivoted_df
        data_ins['to_sql_df'] = to_sql_df

    # ------------------------------------------------
    def vip_daily_sales(self, data_ins) -> None:
        while self is not None:
            print('eliminate the weak warnings')
        # -----------------------------------------
        origin_df, to_sql_df = MiddlewareArsenal.__rectify_daily_sales(data_ins)
        old_time = time.time()
        key_col = data_ins['doc_ref']['key_pos'][0]
        link_col = data_ins['doc_ref']['val_pos'][1]
        origin_df = origin_df[[key_col, link_col]]
        pivoted_df = MiddlewareArsenal.__pivot_daily_sales(data_ins)
        # print('pivot table 耗时: ', time.time() - old_time)
        old_time = time.time()
        merged_df = pd.merge(pivoted_df, origin_df, how='left', on=key_col)
        data_ins['data_frame'] = merged_df
        data_ins['to_sql_df'] = to_sql_df
        print('left join 耗时: ', time.time() - old_time)

    # -------------------------------------------------
    def vip_routine_site_stock(self, data_ins) -> None:
        """
        剔除val_pos列中的无效值, 目前是"-"
        """
        while self is not None:
            print('eliminate the weak warnings')
        val_col = data_ins['doc_ref']['val_pos'][0]
        data_frame = data_ins['data_frame']
        criterion = data_frame[val_col].map(lambda x: str(x).find('-') == -1)
        data_ins['data_frame'] = data_frame[criterion]

    # 这样首尾单下划线的名称结构可以避免和内部属性雷同造成混淆
    def _warehouse_stock_(self, data_ins) -> None:
        while self is not None:
            print('eliminate the weak warnings')
        criteria_col = data_ins['doc_ref']['val_pos'][0]
        data_frame = data_ins['data_frame']
        data_ins['data_frame'] = data_frame[data_frame[criteria_col] != '是']


#  以字典构建dataframe处理函数集合, 后续直接用各个df的identity来调用,
#  注意partial必须传递key argument的限制
middleware_dict = MiddlewareArsenal.__dict__
middleware_arsenal = {}
for func_name, func in middleware_dict.items():
    if re.match(r'^(?=[^_])\w+(?<=[^_])$', func_name):  # 排除系统属性, 如果有大量的regular匹配需求, 最好先调用compile
        middleware_arsenal[func_name] = functools.partial(func, self=None)
    if re.match(r'^(?=_[^_])\w*(?<=[^_]_)$', func_name):
        warehouse_func = functools.partial(func, self=None)
        stock_func_dict = {
            warehouse.lower() + '_stock': warehouse_func
            for warehouse in st.warehouses}
        middleware_arsenal.update(stock_func_dict)
        virtual_stock_func_dict = {
            warehouse.lower() + '_stock_virtual': warehouse_func
            for warehouse in st.warehouses}
        middleware_arsenal.update(virtual_stock_func_dict)


class AssemblyLines:
    """
    -*- 容器 -*-
    各个dataframe之间的整合所需的加工函数在此类的内部类中定义.
    dataframe之间有主、从的区分, 1主单/多从的方式调用.
    主从索引都是identity, 通过内部类的类属性来定义操作method的实参
    所有的内部类的操作method统一命名为assemble, 因此内部类的method定义为class method会更方便调用.
    注: 不需要实例化此类, 直接调用类方法
    """

    class VipElementWiseStockInventory:
        """
        匹配每个唯品条码对应的各仓库存, 首先应把唯品条码map到tmj组合及单品.
        data_ins = {'identity': self.identity, 'doc_ref': self.doc_ref, 'data_frame': data_frame,
        'to_sql_df': sql_df, 'mode': self.from_sql}
        """
        tmj_combination = None
        tmj_atom = None
        vip_fundamental_collections = None

        @classmethod
        def assemble(cls) -> pd.DataFrame():
            master = cls.tmj_combination['data_frame']
            slave = cls.vip_fundamental_collections['data_frame']
            master_key = cls.tmj_combination['doc_ref']['key_pos'][0]
            foreign_key = cls.vip_fundamental_collections['doc_ref']['key_pos'][1]
            master = pd.merge(
                master, slave, how='left', left_on=master_key, right_on=foreign_key
            )
            print(master.head())
            return master

    class VipElementWiseDailySales:

        @classmethod
        def assemble(cls):
            pass

        pass

    class MCElementWiseDailySales:

        @classmethod
        def assemble(cls):
            pass

        pass

    class VipCombinedWithMc:

        @classmethod
        def assemble(cls):
            pass

        pass

    class VipNotes:
        ccc = 100
        ddd = 'good!'

        @classmethod
        def assemble(cls):
            print('vip_notes')

        pass

    class FinalAssembly:
        subassembly = None

        @classmethod
        def assemble(cls):
            if cls.subassembly is None:
                return None
            pass


for x in st.doc_stock_real_and_virtual:
    setattr(AssemblyLines.VipElementWiseStockInventory, x['identity'], None)
# ----------------------------------------------------------
assembly_lines = {}
for attr, attr_value in AssemblyLines.__dict__.items():
    if re.match(r'^(?=[^_])\w+(?<=[^_])$', attr):
        assembly_lines.update({attr: attr_value})
#
# aaa = assembly_lines['VipNotes']
#
# aaa.ddd = 'very good!'
# eee = aaa.__dict__
# print('ccc' in aaa.__dict__)
# for x in assembly_lines:
#     print(x)
# aaa.assemble()

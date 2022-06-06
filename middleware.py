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

    def mc_daily_sales(self, data_ins) -> None:
        while self is not None:
            print('eliminate the weak warnings')
        origin_df = data_ins['data_frame']
        criteria_col = data_ins['doc_ref']['key_pos'][2]
        criterion = pd.concat(
            [origin_df[criteria_col] == 'SO0', origin_df[criteria_col] == 'SO4'], axis=1).any(axis=1)
        data_ins['data_frame'] = origin_df[criterion]
        pivoted_df = MiddlewareArsenal.__pivot_daily_sales(data_ins)
        data_ins['data_frame'] = pivoted_df

    def vip_daily_sales(self, data_ins) -> None:
        while self is not None:
            print('eliminate the weak warnings')
        # -----------------------------------------
        old_time = time.time()
        key_col = data_ins['doc_ref']['key_pos'][0]
        link_col = data_ins['doc_ref']['val_pos'][1]
        origin_df = data_ins['data_frame']
        origin_df = origin_df[[key_col, link_col]]
        pivoted_df = MiddlewareArsenal.__pivot_daily_sales(data_ins)
        print('pivot table 耗时: ', time.time() - old_time)
        old_time = time.time()
        merged_df = pd.merge(pivoted_df, origin_df, how='left', on=key_col)
        data_ins['data_frame'] = merged_df
        print('left join 耗时: ', time.time() - old_time)

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
        # print(data_ins['data_frame'].head())


#  以字典构建dataframe处理函数集合, 后续直接用各个df的identity来调用
middleware_dict = MiddlewareArsenal.__dict__
middleware_arsenal = {}
for func_name, func in middleware_dict.items():
    if re.match(r'^(?=[^_])\w+(?<=[^_])$', func_name):  # 排除系统属性, 如果有大量的regular匹配需求, 最好先调用compile
        middleware_arsenal[func_name] = functools.partial(func, self=None)


# aaa = 123456
# middleware_arsenal["mc_daily_sales"](data_ins=aaa)


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

        """

        @classmethod
        def assemble(cls):
            pass

        pass

    class VipElementWiseDailySales:

        @classmethod
        def assemble(cls):
            pass

        pass

    class McElementWiseDailySales:

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

        @classmethod
        def assemble(cls):
            pass

        pass

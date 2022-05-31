# -*- coding:utf-8 -*-
import multiprocessing
import sys
import re
import time
import datetime
import numpy as np
import pandas as pd
import threading
import multiprocessing
import queue
import warnings
from pathlib import Path
import sqlite3 as sqlite

import settings as st
import sqlite_init


class DocumentIO(threading.Thread):
    """
    基于多线程读取写入文件,判断文件来源.
    实例化此类实现多线程,一个分类文件开启一个线程读取.
    数据库sqlite不支持多线程, 且conn不能存在于多线程中，必须使用mutex进行保护.
    轻量化的application使用sqlite速度已经足够快, mysql显然功能更好, 但是需要单独安装配置数据库.
    """
    # sql_mark标明两个文件是必须从sqldb中读取，一部分然后文件中读取合并在一起.
    # 其他文件都是根据更新时间选择读取来源
    sql_mark = [
        {'identity': 'mc_daily_sales', 'mode': 'merge'},
        {'identity': 'vip_daily_sales', 'mode': 'merge'}
    ]
    sql_db = sqlite_init.sql_db
    files = None  # 需读取的文件信息, 包括identity, file_name, mtime等
    thread_num = 0  # 保存所需线程数, 用于和thread_counter进行比对
    thread_counter = 0  # 统计开启的线程数, 全部读取之后关闭conn
    queue = queue.Queue()
    mutex = threading.Lock()

    @classmethod
    def count_threads(cls):
        cls.thread_counter += 1

    @classmethod
    def get_files_list(cls) -> None:
        files = Path(st.DOCS_PATH)
        files_list = [{
            'identity': None,
            'file_name': str(file),
            'file_mtime': file.stat().st_mtime,
            'file_mtime_in_sqlite': None,
            'read_doc': True,
            'updated_sqlite': False
        } for file in files.glob('*')]
        cls.files = files_list

    @classmethod
    def check_files_list(cls) -> list:
        cls.get_files_list()
        files = str.join(',', [file['file_name'] for file in cls.files])
        for doc in st.DOC_REFERENCE:
            existence = re.search(doc['key_words'], files)
            if existence is None:
                if doc['importance'] == 'required':
                    print(f"缺少必需重要数据表格: {doc['name']}\n")
                    sys.exit()
                elif doc['importance'] == 'caution':
                    print(f"缺少数据表格: {doc['identity']}\n")
                else:
                    pass  # optional文件不存在时不需要提醒
            for file in cls.files:
                matched = re.search(doc['key_words'], file['file_name'])
                if matched is not None:
                    file['identity'] = doc['identity']
                    cls.thread_num += 1
        cls.mutex.acquire()
        # sqlite是单线程,不能线程共用一个conn  # 直接引用类属性中的conn, 不用反复开启连接, 方便适应sqlite连接的单线程特点
        conn = sqlite.connect(cls.sql_db)
        cursor = conn.cursor()
        cursor_data = cursor.execute(
            "SELECT identity, file_name, file_mtime FROM tmj_files_info;")
        # print(files_list)
        for row in cursor_data:  # 把查询到的sqlite中的文件更新时间放入files_list中,后续对比会用到
            for file in cls.files:
                # print(file)
                if file['file_name'] == row[1]:
                    file['file_mtime_in_sqlite'] = row[2]
                    if file['file_mtime'] == row[2]:
                        file['read_doc'] = False
                    # print('修改了')
        cursor.close()
        conn.close()
        cls.mutex.release()
        return cls.files

    def __init__(self, doc_reference: dict):
        super().__init__()
        self.identity = doc_reference['identity']
        self.doc_ref = doc_reference
        self.file = None  # 准备读取的文件名称列表
        self.from_sql = None
        self.to_sql = False
        self.to_sql_df = None  # pandas.DataFrame if not None
        if self.files is None:  # 类属性
            self.files = self.check_files_list()
        self.check_file()

    def check_file(self) -> None:
        for doc in DocumentIO.sql_mark:
            if self.identity == doc['identity']:
                self.from_sql = doc['mode']
        file_name = []
        for file in self.files:  # files是类属性,全部文件夹中的文件信息列表
            if file['identity'] == self.identity:
                file_name.append(file['file_name'])
                if file['file_mtime'] == file['file_mtime_in_sqlite']:
                    self.from_sql = 'substitute'  # 是否从sqlite读取的最终依据是文件是否更新过
        if len(file_name) > 0:
            self.file = file_name
            self.count_threads()  # 存在文件就会开启线程进行读取, thread_counter加1

    def read_doc(self) -> pd.DataFrame():
        doc_df = pd.DataFrame()
        for file in self.file:  # file是实例属性,将要读取的文件信息,也是列表,因为同一性质文件可能有多个
            matched_csv = re.match(r'^.*\.csv$', file)
            matched_excel = re.match(r'^.*\.xlsx?$', file)
            pd_cols = self.doc_ref['key_pos'].copy()  # 直接引用后使用extend方法导致一系列问题
            x = self.doc_ref['val_pos'].copy()  # 避免list出现异常,需要使用copy方法
            pd_cols.extend(x)
            if matched_csv:
                one_df = pd.read_csv(file, usecols=lambda col: col in pd_cols)
                doc_df = pd.concat([doc_df, one_df], ignore_index=True, axis=0)
            if matched_excel:
                # 默认引擎是openpyxl,使用xlrd比openpyxl速度更快,但是必须是新版,pip install xlrd==1.2.0
                # 把关于xlrd的warnings进行捕获, 避免大量的关于xlrd版本的warnings干扰正常提示
                with warnings.catch_warnings(record=True):
                    one_df = pd.read_excel(
                        file, engine='xlrd', usecols=lambda col: col in pd_cols)
                    doc_df = pd.concat([doc_df, one_df], ignore_index=True, axis=0)
        doc_df = doc_df.dropna()  # 剔除空行, 很重要
        if self.from_sql == 'merge':
            sales_date_head = datetime.datetime.today(
            ) - datetime.timedelta(days=st.VIP_SALES_INTERVAL)
            date_col = self.doc_ref['key_pos'][1]
            doc_df = doc_df.loc[lambda df: pd.to_datetime(
                df[date_col]) >= sales_date_head, :]
        return doc_df

    def read_sqlite(self) -> pd.DataFrame():
        pd_cols = self.doc_ref['key_pos'].copy()  # 直接引用后使用extend方法导致一系列问题
        x = self.doc_ref['val_pos'].copy()  # 避免list出现异常,需要使用copy方法
        pd_cols.extend(x)
        sql_constraint = ''
        if self.from_sql == 'merge':
            sales_date_head = datetime.datetime.today(
            ) - datetime.timedelta(days=st.VIP_SALES_INTERVAL)
            # vip和mc日销文件的date列名不同
            sql_constraint = f" WHERE {self.doc_ref['key_pos'][1]} >= '{sales_date_head}'"
        self.mutex.acquire()
        conn = sqlite.connect(self.sql_db)
        # sqlite是单线程,不能线程共用一个conn
        # sql_cursor = conn.cursor()
        # print(pd_cols, '**********')  # tiaoshi
        sql_query = f"SELECT {str.join(',', pd_cols)} FROM {self.identity}{sql_constraint}"
        # 要实现两个df的concat,两者的index列也要相同
        sql_df = pd.read_sql_query(sql_query, con=conn)
        conn.close()
        self.mutex.release()
        return sql_df

    def get_data(self) -> pd.DataFrame():
        if self.from_sql == 'merge':
            doc_df = self.read_doc()
            sql_df = self.read_sqlite()
            sql_date = pd.DataFrame()
            date_col = self.doc_ref['key_pos'][1]
            if not doc_df.empty:
                # doc_df[date_col] = pd.to_datetime(doc_df[date_col])
                self.to_sql_df = doc_df
            if not sql_df.empty:
                # sql_df[date_col] = pd.to_datetime(sql_df[date_col])
                sql_date = sql_df.drop_duplicates(
                    subset=[date_col], keep='first')[date_col]
            if not (doc_df.empty or sql_df.empty):
                mask = [False if x in list(
                    sql_date) else True for x in doc_df[date_col]]
                doc_masked_df = doc_df[mask]
                if doc_masked_df.empty:
                    self.to_sql_df = None
                    merged_df = sql_df
                else:
                    self.to_sql_df = doc_masked_df
                    merged_df = pd.concat(
                        [doc_masked_df, sql_df], ignore_index=True, keys=['doc', 'sqlite'])
                return merged_df
            else:
                # 从数据库中读取的的df为空时, 包含无效的index, 会在concat时报错, 避免使用
                merged_df = doc_df if not doc_df.empty else sql_df
                return merged_df
        elif self.from_sql == 'substitute':
            sql_df = self.read_sqlite()
            return sql_df
        else:
            doc_df = self.read_doc()
            self.to_sql_df = doc_df
            return doc_df

    def run(self) -> None:
        old_time = time.time()
        tracing = f"reading_thread: {self.thread_counter} ({self.identity})is initialized! \r\n" + \
                  f"mode: {self.from_sql} start at: {time.ctime()}\r\n^_^"
        if self.file is None:
            print(f"{self.identity}'s initialization is dispensable!")
        else:
            data_frame = self.get_data()
            sql_df = self.to_sql_df
            df_dict = {'identity': self.identity, 'data_frame': data_frame,
                       'to_sql_df': sql_df, 'mode': self.from_sql}
            self.mutex.acquire()
            self.queue.put(df_dict)
            self.mutex.release()
            # self.to_sqlite()
            print(tracing, f'get it done at: {time.ctime()}  total cost: {time.time() - old_time}\r\n')

    @classmethod
    def update_to_sqlite(cls, queue_ins):
        conn = sqlite.connect(cls.sql_db)  # sqlite是单线程,不能线程共用一个conn
        cursor = conn.cursor()
        while not queue_ins.empty():
            to_sql = queue_ins.get()
            if to_sql['to_sql_df'] is not None:
                if to_sql['mode'] != 'merge':
                    sql_query = f"DELETE FROM {to_sql['identity']};"
                    cursor.execute(sql_query)
                    to_sql['to_sql_df'].to_sql(
                        to_sql['identity'], conn, if_exists='append', index=False, chunksize=1000)
                else:
                    # 需要特别留意DataFrame.to_sql()的参数,必须明确这些参数
                    to_sql['to_sql_df'].to_sql(
                        to_sql['identity'], conn, if_exists='append', index=False, chunksize=1000)
        query_data = []
        for file in cls.files:
            if file['identity'] is not None and file['read_doc']:
                query_data.append(
                    (file['identity'], file['file_name'], file['file_mtime']))
                # 把最新的文件信息写进sqlite中,用于下一次比对,旧信息全部删除.
                cursor.execute(
                    f"DELETE FROM tmj_files_info WHERE identity = '{file['identity']}';")
        cursor.executemany(
            "INSERT INTO tmj_files_info(identity, file_name, file_mtime) VALUES(?,?,?);", query_data)
        conn.commit()
        cursor.close()
        conn.close()

# -*- coding:utf-8 -*-

import sys
import re
import datetime
import numpy as np
import pandas as pd
import threading
import queue
# import multiprocessing
from pathlib import Path
import sqlite3 as sqlite

import settings as st
import sqlite_init


# CPUS = os.cpu_count()
class DocumentIO(threading.Thread):
    """s
    基于多线程读取写入文件,判断文件来源.
    实例化此类实现多线程,外部使用2个进程,每个实例是1个线程,进程内部多线程读取.
    """
    # sql_mark标明两个文件是必须从sqldb中读取，一部分然后文件中读取合并在一起。
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
    conn = sqlite.connect(sql_db)  # 所有线程共用一个conn, 线程计数满足后才关闭conn
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
        conn = cls.conn  # 直接引用类属性中的conn, 不用反复开启连接, 方便适应sqlite连接的单线程特点
        cursor = conn.cursor()
        cursor_data = cursor.execute("SELECT identity, file_name, file_mtime FROM tmj_files_info;")
        # print(files_list)
        for row in cursor_data:  # 把查询到的sqlite中的文件更新时间放入files_list中,后续对比会用到
            for file in cls.files:
                # print(file)
                if file['file_name'] == row[1]:
                    file['file_mtime_in_sqlite'] = row[3]
                    # print('修改了')
        cursor.close()
        cls.mutex.release()
        return cls.files

    def __init__(self, doc_reference):
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
        file_name = []
        for file in self.files:  # files是类属性,全部文件夹中的文件信息列表
            if file['identity'] == self.identity:
                file_name.append(file['file_name'])
                if file['file_mtime'] == file['file_mtime_in_sqlite']:
                    self.from_sql = 'substitute'
        if len(file_name) > 0:
            self.file = file_name
            self.count_threads()  # 存在文件就会开启线程进行读取, thread_counter加1
        for doc in DocumentIO.sql_mark:
            if self.identity == doc['identity']:
                self.from_sql = doc['mode']

    def read_doc(self) -> pd.DataFrame():
        doc_df = pd.DataFrame()
        for file in self.file:  # file是实例属性,将要读取的文件信息,也是列表,因为同一性质文件可能有多个
            matched_csv = re.match(r'^.*\.csv$', file)
            matched_excel = re.match(r'^.*\.xlsx?$', file)
            pd_cols = self.doc_ref['key_pos']
            pd_cols.extend(self.doc_ref['val_pos'])
            if matched_csv:
                one_df = pd.read_csv(file, usecols=lambda col: col in pd_cols)
                doc_df = pd.concat([doc_df, one_df], axis=0)
            if matched_excel:
                # 默认引擎是openpyxl,使用xlrd比openpyxl速度更快,但是必须是新版,pip install xlrd==1.2.0
                one_df = pd.read_excel(file, engine='xlrd', usecols=lambda col: col in pd_cols)
                doc_df = pd.concat([doc_df, one_df], axis=0)
        return doc_df

    def read_sqlite(self) -> pd.DataFrame():
        pd_cols = self.doc_ref['key_pos']
        pd_cols.extend(self.doc_ref['val_pos'])
        sql_constraint = ''
        if self.from_sql == 'merge':
            sales_date_head = datetime.datetime.today() - datetime.timedelta(days=st.VIP_SALES_INTERVAL)
            sql_constraint = f" WHERE {self.doc_ref['key_pos'][1]} >= '{sales_date_head}'"  # vip和mc日销文件的date列名不同
        self.mutex.acquire()
        conn = self.conn  # 引用类属性中的conn
        # sql_cursor = conn.cursor()
        sql_query = f"SELECT {str.join(',', pd_cols)} FROM {self.identity}{sql_constraint}"
        sql_df = pd.read_sql_query(sql_query, con=conn)  # 要实现两个df的concat,两者的index列也要相同
        # cursor.close()
        self.mutex.release()
        return sql_df

    def get_data(self) -> pd.DataFrame():
        if self.from_sql == 'merge':
            doc_df = self.read_doc()
            sql_df = self.read_sqlite()
            sql_date = pd.DataFrame()
            date_col = self.doc_ref['key_pos'][1]
            if not doc_df.empty:
                doc_df[date_col] = pd.to_datetime(doc_df[date_col])
                self.to_sql_df = doc_df
                # doc_date = doc_df.drop_duplicates(subset=[date_col], keep='first')[date_col]
            if not sql_df.empty:
                sql_df[date_col] = pd.to_datetime(sql_df[date_col])
                sql_date = sql_df.drop_duplicates(subset=[date_col], keep='first')[date_col]
            if not (doc_df.empty or sql_df.empty):
                mask = [False if x in sql_date else True for x in doc_df[date_col]]
                doc_masked_df = doc_df[mask]
                if doc_masked_df.empty:
                    merged_df = sql_df
                    self.to_sql_df = None
                else:
                    merged_df = pd.concat([doc_masked_df, sql_df], keys=['doc_df', 'sql_df'])
                    self.to_sql_df = doc_masked_df
                return merged_df
            else:
                valid_df = doc_df if not doc_df.empty else sql_df
                return valid_df
        elif self.from_sql == 'substitute':
            sql_df = self.read_sqlite()
            return sql_df
        else:
            doc_df = self.read_doc()
            self.to_sql_df = doc_df
            return doc_df

    def to_sqlite(self):
        self.mutex.acquire()
        conn = self.conn
        cursor = conn.cursor()
        if self.to_sql_df is not None:
            if self.from_sql != 'merge':
                sql_query = f"DELETE FROM {self.identity};"
                cursor.execute(sql_query)
            else:
                self.to_sql_df.to_sql(self.identity, conn)
            self.to_sql = True

        query_data = []
        for file in self.files:
            if file['identity'] == self.identity:
                query_data.append((file['identity'], file['file_name'], file['file_mtime']))
        # 把最新的文件信息写进sqlite中,用于下一次比对,旧信息全部删除.
        cursor.execute(f"DELETE FROM tmj_files_info WHERE identity = '{self.identity}';")
        cursor.executemany(
            "INSERT INTO tmj_files_info(identity, file_name, file_mtime) VALUES(?,?,?,?);", query_data)
        conn.commit()
        cursor.close()
        if self.thread_counter == self.thread_num:
            conn.close()
        self.mutex.release()

    def run(self) -> None:
        if self.file is None:
            print(f"{self.identity}'s initialization is dispensable!")
        else:
            data_frame = self.get_data()
            df_dict = {'identity': self.identity, 'data_frame': data_frame}
            self.mutex.acquire()
            self.queue.put(df_dict)
            self.mutex.release()
            self.to_sqlite()





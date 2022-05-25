# -*- coding:utf-8 -*-

import sys
import re
import datetime
import numpy as np
import pandas as pd
import threading
import multiprocessing
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
    files = None
    mutex = threading.Lock()
    queue = multiprocessing.Queue()

    @classmethod
    def get_files_list(cls):
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
        cls.mutex.aquaire()
        conn = sqlite.connect(cls.sql_db)
        cursor = conn.cursor()
        cursor_data = cursor.execute("SELECT identity, file_name, file_mtime FROM tmj_files_info;")
        # print(files_list)
        for row in cursor_data:  # 把查询到的sqlite中的文件更新时间放入files_list中,后续对比会用到
            for file in cls.files:
                # print(file)
                if file['identity'] == row[1]:
                    file['file_mtime_in_sqlite'] = row[3]
                    # print('修改了')
        conn.close()
        cls.mutex.release()
        return cls.files

    def __init__(self, doc_reference):
        super().__init__()
        self.identity = doc_reference['identity']
        self.doc_ref = doc_reference
        self.file = None
        self.from_sql = None
        self.queue = self.queue
        self.mutex = self.mutex
        if DocumentIO.files is None:
            self.files = DocumentIO.check_files_list()
        self.check_file()

    def check_file(self):
        file_name = []
        for file in self.files:
            if file['identity'] == self.identity:
                file_name.append(file['file_name'])
                if file['file_mtime'] == file['file_mtime_in_sqlite']:
                    self.from_sql = 'substitute'
        if len(file_name) > 0:
            self.file = file_name
        for doc in DocumentIO.sql_mark:
            if self.identity == doc['identity']:
                self.from_sql = doc['mode']

    def read_doc(self) -> pd.DataFrame:
        doc_df = pd.DataFrame()
        for file in self.file:
            matched_csv = re.match(r'^.*\.csv$', file)
            matched_excel = re.match(r'^.*\.xlsx?$', file)
            pd_cols = self.doc_ref['key_pos'].extend(self.doc_ref['val_pos'])
            if matched_csv:
                df = pd.read_csv(file, usecols=lambda col: col in pd_cols)
                doc_df = pd.concat([doc_df, df], axis=0)
            if matched_excel:
                df = pd.read_excel(file, usecols=lambda col: col in pd_cols)  # 在read_excel中使用index_col=[]报错,不知道原因
                doc_df = pd.concat([doc_df, df], axis=0)
        return doc_df

    def read_sqlite(self) -> pd.DataFrame:
        pd_cols = self.doc_ref['key_pos'].extend(self.doc_ref['val_pos'])
        sql_constraint = ''
        if self.from_sql == 'merge':
            sales_date_head = datetime.datetime.today() - datetime.timedelta(days=st.VIP_SALES_INTERVAL)
            sql_constraint = f' WHERE 日期 >= {sales_date_head}'
        self.mutex.aquaire()
        conn = sqlite.connect(self.sql_db)
        # sql_cursor = conn.cursor()
        sql_query = f"SELECT {str.join(',', pd_cols)}, FROM {self.identity}{sql_constraint}"
        sql_df = pd.read_sql_query(sql_query, con=conn, index_col=self.doc_ref['key_pos'])
        conn.close()
        self.mutex.release()
        return sql_df

    def get_data(self):
        if self.from_sql == 'merge':
            doc_df = self.read_doc()
            sql_df = self.read_sqlite()
            if not doc_df.empty:
                doc_df.assign(newdate=lambda x: pd.to_datetime(x['日期']))
            if not sql_df.empty:
                sql_df.assign(newdate=lambda x: pd.to_datetime(x['日期']))
        elif self.from_sql == 'substitute':
            self.read_sqlite()

        pass

    def to_sqlite(self):
        self.mutex.aquaire()
        conn = sqlite.connect(self.sql_db)
        cursor = conn.cursor()
        pass
        '''

        '''
        query_data = []
        count = 1
        for file in self.files:
            if file['identity'] is not None:
                query_data.append((count, file['identity'], file['file_name'], file['file_mtime']))
                count += 1
        # 把最新的文件信息写进sqlite中,用于下一次比对,旧信息全部删除.
        print(self.files)
        cursor.execute("DELETE FROM tmj_files_info;")
        cursor.executemany(
            "INSERT INTO tmj_files_info(id, identity, file_name, file_mtime) VALUES(?,?,?,?);", query_data)
        conn.commit()
        conn.close()

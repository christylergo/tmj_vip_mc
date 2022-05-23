# -*- coding:utf-8 -*-

import sys
import re
import numpy as np
import pandas as pd
import threading
import multiprocessing
from pathlib import Path
import sqlite3 as sqlite

import settings as st

# CPUS = os.cpu_count()

def get_files_list(files_path)->list:
    files = Path(files_path)
    files_list = [{
        'identity': None,
        'file_name': str(file),
        'file_mtime': file.stat().st_mtime,
        'file_mtime_in_sqlite': None
        } for file in files.glob('*')]
    return files_list

def check_file(doc_reference, files_list)->None:
    files = str.join(files_list['file_name'],',')
    for doc in doc_reference:
        existence = re.search(doc['key_words'], files)
        if existence is None:
            if doc['importance'] == 'required':
                print(f"缺少必需重要数据表格: {doc['name']}\n")
                sys.exit()
            elif doc['importance'] == 'caution':
                print(f"缺少数据表格: {doc['name']}\n")
            else:
                pass
        for file in files_list:
            matched = re.search(doc['key_words'], files_list['file_name'])
            if matched is not None:
                file['identity'] = doc['identity']

    conn = sqlite.connect('tmj_sqlite.db')
    sql_cursor = conn.cursor()
    sql_cursor = sql_cursor.execute("SELECT id, identity, file_name, file_mtime, FROM tmj_files_info")
    conn.close()
    for file in files_list:  # 把查询到的sqlite中的文件更新时间放入files_list中,后续对比会用到
        for row in sql_cursor:
            if  file['identity'] == row[1]:
                file['file_mtime_in_sqlite'] = row[3]
            

FILES_LIST = get_files_list(st.DOCS_PATH)
DOC_REFERENCE = st.DOC_REFERENCE
check_file(DOC_REFERENCE, FILES_LIST)


class DocumentIO(threading.Thread):
    """
    基于多线程读取写入文件,判断文件来源.
    实例化此类实现多线程,外部使用2个进程,每个实例是1个线程,进程内部多线程读取.
    """
    # sql_mark标明两个文件是必须从sqldb中读取，一部分然后文件中读取合并在一起。
    # 其他文件都是根据更新时间选择读取来源
    sql_mark = [
        {'identity': 'mc_daily_sales', 'mode': 'merge'},
        {'identity': 'vip_daily_sales', 'mode': 'merge'}]
    sql_db = 'tmj_sqlite.db'

    def __init__(self, doc_reference, files_list, data_queue: multiprocessing.Queue):
        super().__init__()
        self.identity = doc_reference['identity']
        self.doc_ref = doc_reference
        self.files = files_list
        self.file = None
        self.from_sql = None
        self.doc_data = None
        self.data_queue = data_queue
        self.check_file()

    def check_file(self):
        for file in self.files:
            if file['identity'] == self.identity:
                self.file = file['file_name']
                if file['file_mtime'] == file['file_mtime_in_sqlite']:
                    self.from_sql = 'substitute'
        for doc in DocumentIO.sql_mark:
            if self.identity == doc['identity']:
                self.from_sql = doc['mode']

    def doc_io(self):
        matched_csv = re.match('^.*\.csv$', self.file)
        matched_excel = re.match('^.*\.xlsx?$', self.file)
        pd_cols = self.doc_ref['key_pos'].extend(self.doc_ref['val_pos'])
        if matched_csv:
            self.doc_data = pd.read_csv(self.file, index_col=self.doc_ref['key_pos'], usecols=lambda col: col in pd_cols)
        if matched_excel:
            self.doc_data = pd.read_excel(self.file, index_col=self.doc_ref['key_pos'], use_cols=lambda col: col in pd_cols)

    def sqlite_io(self):
        pass

    def get_data(self):
        if self.sql == 'merge':
            self.doc_io()
            self.sqlite_io()
        elif self.from_sql == 'substitute':
            pass

        pass

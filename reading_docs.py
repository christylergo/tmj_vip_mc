# -*- coding:utf-8 -*-

import os
import re
import numpy as np
import pandas as pd
import threading
from multiprocessing import process
from multiprocessing import Queue
from pathlib import Path

import settings as st
import sqlite_io

CPUS = os.cpu_count()
# -----------------------------------------------
# 读取sqlite中的存档信息，唯品总货表，唯品猫超30天销量。
pass
preserved_file_name = '唯品会十月总货表'
preserved_file_mtime = ''
pass
# ------------------------------------------------
# print(st.DOC_REFERENCE)
vip_fundamental_collections = re.compile(preserved_file_name)


# docs = Path(st.DOCS_PATH)
# for doc in docs.glob('*'):
#     cc = doc.stat().st_mtime
#     print(doc)


def get_files_list(files_path):
    docs = Path(files_path)
    files_list = [
        {'file_name': str(doc),
         'file_mtime': doc.stat().st_mtime,
         } for doc in docs.glob('*')
    ]


class ReadDocument:
    def __init__(self, doc_reference, files_list):
        self.identity = doc_reference['identity']
        self.doc_ref = doc_reference
        self.files = files_list
        self.sql_mark = [{'identity': 'mc_daily_sales', 'mode': 'append'},
                         {'identity': 'vip_daily_sales', 'mode': 'append'},
                         {'identity': 'vip_fundamental_collections', 'mode': 'replace'}]
        self.sql = None

    def check_file(self):
        for doc in self.sql_mark:
            if self.identity == doc['identity']:
                self.sql = doc['mode']

    def from_doc(self):
        pass

    def from_sqlite(self):
        pass

    def get_data(self):
        self.check_file()
        if self.sql == 'append':
            self.from_doc()
            self.from_sqlite()
        elif self.sql == ''

        pass

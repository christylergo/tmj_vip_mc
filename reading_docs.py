# -*- coding:utf-8 -*-

import os
import sys
import re
import time
import datetime
import queue
import pandas as pd
import threading
import multiprocessing
import warnings
from pathlib import Path
import sqlite3 as sqlite

import settings as st
import sqlite_init

CPUS = os.cpu_count()


class DocumentIO(threading.Thread):
    """
    基于多线程读取写入文件,判断文件来源.
    实例化此类实现多线程,一个分类文件开启一个线程读取, 放入队列中的数据结构在实例方法run中定义.
    如果主要是读取sqlite则不开启多进程, 如果有较多的文件要读取, 则每2类文件开启一个进程去读取, 多核处理器提速很明显.
    数据库sqlite不支持多线程, 且conn不能存在于多线程中，必须使用mutex进行保护.
    轻量化的application使用sqlite速度已经足够快, mysql显然功能更好, 但是需要单独安装配置数据库.
    """
    # sql_mark标明两个文件是必须从sqldb中读取，一部分然后文件中读取合并在一起.
    # 其他文件都是根据更新时间选择读取来源
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
        self.from_sql = doc_reference['mode']
        self.to_sql = False
        self.to_sql_df = None  # pandas.DataFrame if not None
        if self.files is None:  # 类属性
            self.files = self.check_files_list()
        self.check_file()

    def check_file(self) -> None:
        file_name = []
        read_doc = False
        for file in self.files:  # files是类属性,全部文件夹中的文件信息列表
            if file['identity'] == self.identity:
                file_name.append(file['file_name'])
                read_doc = read_doc or file['read_doc']
        if not read_doc:
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
                if self.identity == 'mc_daily_sales':  # 这个针对性操作应该放进middleware里
                    doc_df[date_col] = pd.to_datetime(doc_df[date_col]).dt.date
                    doc_df[date_col] = doc_df[date_col].astype('str')
                self.to_sql_df = doc_df
            if not sql_df.empty:
                if self.identity == 'mc_daily_sales':
                    sql_df[date_col] = pd.to_datetime(sql_df[date_col]).dt.date
                    sql_df[date_col] = sql_df[date_col].astype('str')
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
                  f"mode: {self.from_sql}   start at: {time.ctime()} ^_^\r\n"
        if self.file is None:
            print(f"{self.identity}'s initialization is dispensable!")
        else:
            data_frame = self.get_data()
            sql_df = self.to_sql_df
            #  放入queue中的数据的结构
            df_dict = {'identity': self.identity, 'data_frame': data_frame,
                       'to_sql_df': sql_df, 'mode': self.from_sql}
            self.mutex.acquire()
            self.queue.put(df_dict)
            self.mutex.release()
            # self.to_sqlite()
            tracing = tracing + f'get it done at: {time.ctime()}  total cost: {time.time() - old_time}\r\n'
            print(tracing)

    @classmethod
    def update_to_sqlite(cls, list_ins: list) -> None:
        """
        因为多进程读写同一个sqlite表很可能会出现连接冲突,所以单独定义写入sqlite的类方法.
        此方法最后单独执行, 避免冲突. tmj_files_info表是各个线程或进程共用的.
        最后单独写入, 可以避免信息混乱. 还能规避连接冲突.
        :param list_ins:
        :return:
        """
        conn = sqlite.connect(cls.sql_db)  # sqlite是单线程,不能线程共用一个conn
        cursor = conn.cursor()
        query_data = []
        for to_sql in list_ins:
            if to_sql['to_sql_df'] is not None:
                if to_sql['mode'] == 'merge':
                    # 需要特别留意DataFrame.to_sql()的参数,必须明确这些参数
                    to_sql['to_sql_df'].to_sql(
                        to_sql['identity'], conn, if_exists='append', index=False, chunksize=1000)
                    print(to_sql['to_sql_df'].head())
                else:
                    sql_query = f"DELETE FROM {to_sql['identity']};"
                    cursor.execute(sql_query)
                    to_sql['to_sql_df'].to_sql(
                        to_sql['identity'], conn, if_exists='append', index=False, chunksize=1000)
                # 写入sqlite的文件更新信息, 避免出现线程执行失败, 但是文件信息却更新了的情况
                for file in cls.files:
                    if file['identity'] == to_sql['identity']:
                        query_data.append(
                            (file['identity'], file['file_name'], file['file_mtime']))
                        # 把最新的文件信息写进sqlite中,用于下一次比对,旧信息全部删除.
                        cursor.execute(
                            f"DELETE FROM tmj_files_info WHERE identity = '{file['identity']}';")
        #  --------------------------------
        cursor.executemany(
            "INSERT INTO tmj_files_info(identity, file_name, file_mtime) VALUES(?,?,?);", query_data)
        conn.commit()
        cursor.close()
        conn.close()


# ----------------------------------分隔线, 之后是功能函数---------------------------------------


def reading_worker(process_queue=None, doc_refer=None, /) -> None:
    if doc_refer is None:
        doc_refer = st.DOC_REFERENCE
    rds_ins = []
    for xx in doc_refer:
        temp = DocumentIO(xx)
        if temp.file is not None:
            rds_ins.append(temp)
            temp.start()
    for ins in rds_ins:
        ins.join()
    while not DocumentIO.queue.empty():
        data_ins = DocumentIO.queue.get()
        process_queue.put(data_ins)


def multiprocessing_reader() -> list:
    """
    返回值是字典列表
    {'identity': identity, 'data_frame': pd.DataFrame(), to_sql_df': pd.DataFrame(), 'mode': 'merge'/'substitute'/None}
    :return:
    """
    global CPUS
    files_list = DocumentIO.check_files_list()
    doc_reference = []
    sql_reference = []
    for doc in st.DOC_REFERENCE:
        zzz = None
        for x in files_list:
            if x['identity'] == doc['identity']:
                if x['read_doc']:
                    zzz = doc
        if zzz is None:
            sql_reference.append(doc)
        else:
            doc_reference.append(zzz)
    len_doc = len(doc_reference)
    if len_doc > 2:
        print('multiprocessing is initialized.')
        pool = multiprocessing.Pool(CPUS)
        queue_ins = multiprocessing.Manager().Queue()
        for i in range(len_doc // 2):  # 每2个文档读取需求开启一个进程
            doc_group = [doc_reference[i * 2], doc_reference[i * 2 + 1]]
            pool.apply_async(reading_worker, (queue_ins, doc_group))
            # print(doc_group)
        doc_group = sql_reference  # 把需要从sqlite中读取的需求也加进最后一个进程
        if len_doc % 2 == 1:
            doc_group.append(doc_reference[-1])  # 把奇数末尾一个文档读取的需求也加进最后一个进程
        pool.apply_async(reading_worker, (queue_ins, doc_group))
        pool.close()
        pool.join()
    else:
        queue_ins = queue.Queue()
        reading_worker(queue_ins)  # 将数据放入便于读取的queue中
    data_ins_list = []
    while not queue_ins.empty():
        data_ins = queue_ins.get()
        data_ins_list.append(data_ins)
    DocumentIO.update_to_sqlite(data_ins_list)  # 最后更新文件信息,避免干扰读取
    return data_ins_list

    # print('CPU_CORES: ', CPUS)
    # print('********* all things are done! *********')

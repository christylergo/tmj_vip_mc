# -*- coding:utf-8 -*-

import numpy as np
import pandas as pd
import openpyxl as xl

import settings as st
import reading_docs as rds

rds_ins = []
for doc in st.DOC_REFERENCE:
    temp = rds.DocumentIO(doc)
    if temp.file is not None:
        rds_ins.append(temp)
        temp.start()
for ins in rds_ins:
    ins.join()
for i in range(20):
    if not rds.DocumentIO.queue.empty():
        data_ins = rds.DocumentIO.queue.get()
        # if data_ins['identity'] == 'vip_routine_site_stock':
        print(data_ins['identity'])
        print(data_ins['data_frame'].head())
        print(rds.DocumentIO.queue.qsize())

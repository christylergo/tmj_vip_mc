# -*- coding:utf-8 -*-

import numpy as np
import pandas as pd
import openpyxl as xl

import settings as st
import reading_docs as rds


for doc in st.DOC_REFERENCE:
    rds_ins = rds.DocumentIO(doc)
    if rds_ins.file is not None:
        rds_ins.start()
        rds_ins.join()
if not rds.DocumentIO.queue.empty():
    data_ins = rds.DocumentIO.queue.get()
    print(data_ins['identity'])
    print(data_ins['data_frame'].head())

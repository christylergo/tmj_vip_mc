# -*- coding:utf-8 -*-

import multiprocessing

import settings as st
import reading_docs as rds


def reading_process(process_queue=None, doc_refer=None):
    if doc_refer is None:
        doc_refer = st.DOC_REFERENCE
    rds_ins = []
    for xx in doc_refer:
        temp = rds.DocumentIO(xx)
        if temp.file is not None:
            rds_ins.append(temp)
            temp.start()
    for ins in rds_ins:
        ins.join()
    for i in range(20):
        if not rds.DocumentIO.queue.empty():
            data_ins = rds.DocumentIO.queue.get()
            if process_queue is not None:
                process_queue.put(data_ins)
            # if data_ins['identity'] == 'vip_routine_site_stock':
            print(data_ins['identity'])
            print(data_ins['data_frame'].head())
            print(rds.DocumentIO.queue.qsize())


queue = multiprocessing.Queue()
from_doc_list = rds.DocumentIO.check_files_list()
doc_reference = []
for x in from_doc_list:
    if x['read_doc']:
        for doc in st.DOC_REFERENCE:
            if doc['identity'] == x['identity']:
                doc_reference.append(doc)
len_doc = len(doc_reference)
if len_doc > 3:
    p_list = []
    for i in range(len_doc//2):
        doc_group = [doc_reference[i*2], doc_reference[i*2+1]]
        if len_doc == i*2+3:
            doc_group = [doc_reference[i * 2], doc_reference[i * 2 + 1], doc_reference[i * 2 + 2]]
        p = multiprocessing.Process(target=reading_process, args=(queue, doc_group))
        p.start()
        p_list.append(p)
    for p in p_list:
        p.join()
else:
    reading_process()



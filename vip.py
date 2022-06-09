# -*- coding:utf-8 -*-
import time
import reading_docs as rds
from middleware import middleware_arsenal

if __name__ == '__main__':
    raw_data = rds.multiprocessing_reader()
    # 对已读取的dataframe进行预处理
    processed_data = []
    for data_ins in raw_data:
        identity = data_ins['identity']
        if identity in middleware_arsenal.keys():
            preprocess_func = middleware_arsenal[identity]  # 可能会出现KeyError
            preprocess_func(data_ins=data_ins)  # partial对象需要传递key argument
    processed_data = raw_data
    old_time = time.time()
    rds.DocumentIO.update_to_sqlite(processed_data)  # 最后更新文件信息,避免干扰读取
    print('写入sqlite耗时: ', time.time() - old_time)


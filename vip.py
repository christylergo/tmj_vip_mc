# -*- coding:utf-8 -*-
import reading_docs as rds
from middleware import middleware_arsenal

if __name__ == '__main__':
    full_data = rds.multiprocessing_reader()
    # 对已读取的dataframe进行预处理
    for data_ins in full_data:
        identity = data_ins['identity']
        try:
            preprocess_func = middleware_arsenal[identity]
            preprocess_func(data_ins=data_ins)  # partial对象需要传递key argument
            # print(data_ins['data_frame'].head(25))
        except KeyError:
            pass


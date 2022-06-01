# -*- coding:utf-8 -*-
import reading_docs as rds

if __name__ == '__main__':
    full_data = rds.multiprocessing_reader()
    for xx in full_data:
        p_str = f"identity: {xx['identity']}\r\n{xx['data_frame'].head(15)}\r\n"
        print(p_str)

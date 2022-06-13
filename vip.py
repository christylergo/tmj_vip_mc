# -*- coding:utf-8 -*-
import time
import reading_docs as rds
from middleware import middleware_arsenal
from middleware import assembly_lines

if __name__ == '__main__':
    raw_data = rds.multiprocessing_reader()
    # 对已读取的dataframe进行预处理
    for data_struct in raw_data:
        identity = data_struct['identity']
        # 避免可能会出现KeyError. try except是最差的方式, 会掩盖其他类型的error
        preprocess_func = middleware_arsenal.get(identity, lambda data_ins: data_ins)
        preprocess_func(data_ins=data_struct)  # partial对象需要传递key argument
    processed_data = raw_data
    rds.DocumentIO.update_to_sqlite(processed_data)  # 最后更新文件信息,避免干扰读取
    # ---------------------------------------------------------
    data_dict = {x['identity']: x for x in processed_data}
    assembled_data = {}
    for inner_class in assembly_lines:
        for identity in data_dict:
            if identity in inner_class.__dict__:
                setattr(inner_class, identity, data_dict[identity])
        df = inner_class.assemble()
        assembled_data.update({inner_class: df})
    final_assembly = assembly_lines['FinalAssembly']
    final_assembly.subassembly = assembled_data
    final_data = final_assembly.assemble()

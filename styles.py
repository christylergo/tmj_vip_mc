# -*- coding: utf-8 -*-
import openpyxl as xl
import pandas as pd
import numpy as np

df = pd.DataFrame({
    "Name": ["Braund, Mr. Owen Harris", "Allen, Mr. William Henry", "Bonnell, Miss. Elizabeth", ],
    "Age": [22, 35, 58],
    "Sex": ["male", "male", "female"],
}
)
zzz = df[df['Age']>30]['Sex']
df['Sex'] = 'unknown'
print(np.ceil(0.01))
"""
A value is trying to be set on a copy of a slice from a DataFrame.
Try using .loc[row_indexer,col_indexer] = value instead
"""
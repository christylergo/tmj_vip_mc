aaa = [{'a1': 1, 'a2': 2}, {'a1': 1, 'b2': 2}]


def func(aa):
    print('func: ',aa)


ccc = ['a', 'b', 'c']

ddd = {x: func for x in ccc}
ddd['a'](9)
ddd['c'](100)



print(aaa[1].keys())

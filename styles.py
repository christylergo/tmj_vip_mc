aaa = [{'a1': 1, 'a2': 2}, {'a1': 1, 'b2': 2}]


def func(aa):
    print('func: ',aa)


ccc = ['a', 'b', 'c']
fff = [1, 2]
ddd = {x: func for x in ccc}
ddd['a'](9)
ddd['c'](100)

fff.extend(ccc)

print(fff, ccc)

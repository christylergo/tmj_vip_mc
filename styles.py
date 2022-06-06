aaa = [{'a1': 1, 'a2': 2}, {'a1': 1, 'b2': 2}]


def func(aa):
    for a in aa:
        a['a1'] = 9


func(aaa)
print(aaa)

a = -1
if bool(a):
    print('-1true')

ccc = [1, -2, -9]

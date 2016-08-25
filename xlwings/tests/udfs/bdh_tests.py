import xlwings as xw

import random


@xw.func
@xw.ret(expand='table')
def test_bdh():
    n = random.randint(1, 10)
    m = random.randint(1, 10)
    return [
        [random.random() for j in range(m)]
        for i in range(n)
    ]
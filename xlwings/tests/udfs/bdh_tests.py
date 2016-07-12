import xlwings as xw

import random


@xw.func
@xw.ret(expand=True)
def test_bdh():
    n = random.randint(0, 32)
    return [
        [random.random()]
        for i in range(n)
    ]
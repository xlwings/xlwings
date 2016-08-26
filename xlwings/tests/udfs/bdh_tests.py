import xlwings as xw

import random


@xw.func
@xw.arg("n", numbers=int)
@xw.arg("m", numbers=int)
@xw.ret(expand='table')
def test_bdh(n, m):
    return [
        [ "%s x %s : %s, %s" % (n, m, i, j) for j in range(m)]
        for i in range(n)
    ]


@xw.func
@xw.arg("n", numbers=int)
@xw.arg("m", numbers=int)
def test_af(n, m):
    return [
        [ "%s x %s : %s, %s" % (n, m, i, j) for j in range(m)]
        for i in range(n)
    ]



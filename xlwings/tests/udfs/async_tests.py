import xlwings as xw
import time


@xw.func(async='threading')
def async1(a):
    time.sleep(2)
    return a


@xw.func(async='threading')
@xw.arg("n", numbers=int)
@xw.arg("m", numbers=int)
@xw.ret(expand='table')
def async2(n, m):
    time.sleep(2)
    return [
        ["%s x %s : %s, %s" % (n, m, i, j) for j in range(m)]
        for i in range(n)
    ]

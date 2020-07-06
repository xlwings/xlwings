import xlwings as xw
import time
import numpy as np
import pandas as pd


@xw.func(async_mode='threading')
@xw.arg("n", numbers=int)
@xw.arg("m", numbers=int)
@xw.ret(expand='table')
def threading(n, m):
    print('1 - CALLED THREADING FUNCTION')
    time.sleep(1)
    return [
        ["%s x %s : %s, %s" % (n, m, i, j) for j in range(m)]
        for i in range(n)
    ]


@xw.arg("n", numbers=int)
@xw.arg("m", numbers=int)
@xw.ret(expand='table')
def nothreading(n, m):
    print('2 - CALLED NOTHREADING FUNCTION')
    return [
        ["%s x %s : %s, %s" % (n, m, i, j) for j in range(m)]
        for i in range(n)
    ]


@xw.func
def simple(x):
    print('3 - CALLED SIMPLE FUNCTION')
    return x


@xw.func(async_mode='threading')
def simple_threading(x):
    print('4 - CALLED SIMPLE THREADING FUNCTION')
    time.sleep(1)
    return x


@xw.func(async_mode='threading')
@xw.arg("i", numbers=int)
@xw.arg("j", numbers=int)
@xw.ret(expand='table')
def numpy_async(i, j):
    print('5 - CALLED NUMPY ASYNC FUNCTION')
    time.sleep(1)
    return np.arange(i * j).reshape((i, j))


if __name__ == '__main__':
    xw.serve()

import asyncio
import time

import numpy as np

import xlwings as xw


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


@xw.arg("n", numbers=int)
@xw.arg("m", numbers=int)
@xw.ret(expand='table')
async def coro(n, m):
    print('3 - CALLED COROUTINE')
    return [
        ["%s x %s : %s, %s" % (n, m, i, j) for j in range(m)]
        for i in range(n)
    ]


@xw.func
def simple(x):
    print('4 - CALLED SIMPLE FUNCTION')
    return x


@xw.func(async_mode='threading')
def simple_threading(x):
    print('5 - CALLED SIMPLE THREADING FUNCTION')
    time.sleep(1)
    return x


@xw.func
async def simple_coro(x):
    print('6 - CALLED SIMPLE COROUTINE')
    await asyncio.sleep(1)
    return x


@xw.func(async_mode='threading')
@xw.arg("i", numbers=int)
@xw.arg("j", numbers=int)
@xw.ret(expand='table')
def numpy_async(i, j):
    print('7 - CALLED NUMPY ASYNC FUNCTION')
    time.sleep(1)
    return np.arange(i * j).reshape((i, j))


@xw.func(async_mode='threading')
@xw.arg('behavior', numbers=int)
@xw.ret(expand='table')
def formula_erased_2(behavior):
    """
    GH1010 Call 1 then as soon as you can call 2,
    then 3 -- Formula will disappear as you will
    call 3 before 2 even returned
    """
    print("formula_erased_2", behavior)
    if behavior == 1:
        return [['value'] * 20] * 300
    if behavior == 2:
        time.sleep(5)
        print("2 returns")
        return [['value'] * 20] * 250
    return [['value'] * 20] * 200


@xw.func
@xw.arg('nb_rows', numbers=int)
@xw.ret(expand='table')
def hello(nb_rows):
    print("hello", nb_rows)
    return ['value'] * nb_rows


@xw.func
@xw.arg('data', expand='table')
@xw.ret(expand='table', index=False)
def print_table(data):
    """
    GH1164 Pointing this function at the top
    left cell of a table, the output only
    changes when the top left cell changes,
    and not any other cell in the table
    """
    print('print_table', np.array(data).shape)
    return data


if __name__ == '__main__':
    xw.serve()

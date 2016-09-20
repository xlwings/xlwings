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


@xw.func
@xw.arg("n", numbers=int)
@xw.ret(expand='down')
def write_list(n):
    """
    ``down`` for this use case is undocumented as the only advantage of using it instead of ``table`` is that if you have
    non-empty cells immediately to the right of the table, then those won't get cleared out by xlwings. But this might
    be resolved by introducing something like ``clear_border=False`` at some point.

    """
    return [
        [i, random.random()]
        for i in range(n)
    ]
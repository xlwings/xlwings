"""
Custom Functions (UDFs)
"""

import xlwings as xw


@xw.func
def add(first, second, third=None):
    if third:
        return first * second * third
    else:
        return first * second


@xw.func
async def add2(first, second):
    return first + second

"""
Custom Functions (UDFs)
"""

from xlwings import pro


@pro.func
def add(first, second, third=None):
    if third:
        return first + second + third
    else:
        return first + second


@pro.func
async def add2(first, second):
    return first + second

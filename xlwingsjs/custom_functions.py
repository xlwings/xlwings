"""
Custom Functions (UDFs)
"""

from xlwings.pro import func


@func
def add(first, second, third=None):
    if third:
        return first + second + third
    else:
        return first + second


@func
async def add2(first, second):
    return first + second

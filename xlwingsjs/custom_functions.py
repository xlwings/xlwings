from xlwings import server


@server.func
def add(first, second, third=None):
    if third:
        return first + second + third
    else:
        return first + second


@server.func
async def add2(first, second):
    return first + second

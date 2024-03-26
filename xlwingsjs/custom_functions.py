import asyncio

import httpx
import numpy as np
import pandas as pd

from xlwings import server


@server.func
def hello(name):
    return f"hello {name}"


@server.func
async def random_numbers1(first, second):
    rng = np.random.default_rng()
    while True:
        matrix = rng.standard_normal(size=(first, second))
        df = pd.DataFrame(matrix, columns=[f"col{i+1}" for i in range(matrix.shape[1])])
        yield df
        await asyncio.sleep(1)


@server.func
async def random_numbers2(first, second):
    rng = np.random.default_rng()
    while True:
        matrix = rng.standard_normal(size=(first, second))
        df = pd.DataFrame(matrix, columns=[f"col{i+1}" for i in range(matrix.shape[1])])
        yield df + 10
        await asyncio.sleep(1)


@server.func
async def reciprocal(x):
    while True:
        yield 1 / x
        await asyncio.sleep(1)


@server.func
@server.ret(date_format="hh:mm:ss", index=False)
async def btc_price(base_currency="USD"):
    while True:
        async with httpx.AsyncClient() as client:
            response = await client.get(
                f"https://cex.io/api/ticker/BTC/{base_currency}"
            )
        response_data = response.json()
        response_data["timestamp"] = pd.to_datetime(
            int(response_data["timestamp"]), unit="s"
        )
        df = pd.DataFrame(response_data, index=[0])
        df = df[["pair", "timestamp", "bid", "ask"]]
        yield df
        await asyncio.sleep(1)

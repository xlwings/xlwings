import xlwings as xw
import datetime as dt
import numpy as np
import pandas as pd

wb = xw.Workbook('Workbook1')

xw.Range('A6').options(transpose=True).value = [1,2,3]  # [1,2,3]

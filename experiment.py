import xlwings as xw
from time import time

import numpy as np

wb = xw.Workbook.active()

times = []

print("---default---")
for i in range(10):
    start_time = time()
    xw.Range('A1:Z20000').options().value
    end_time = time()
    print("%ims" % (1000 * (end_time - start_time)))
    times.append(1000 * (end_time - start_time))

print("avg %ims stdev %ims" % (np.mean(times), np.std(times)))

times = []

print("---raw---")
for i in range(10):
    start_time = time()
    xw.Range('A1:Z20000').options('raw').value
    end_time = time()
    print("%ims" % (1000 * (end_time - start_time)))
    times.append(1000 * (end_time - start_time))

print("avg %ims stdev %ims" % (np.mean(times), np.std(times)))

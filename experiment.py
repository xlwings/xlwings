import xlwings as xw

s = xw.active.sheet

cs = xw.active.sheet.charts

print(len(cs))

print(cs(1))
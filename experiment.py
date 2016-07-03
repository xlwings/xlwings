import xlwings as xw

s = xw.sheets.active

c = s.charts.add()

print(c.chart_type)
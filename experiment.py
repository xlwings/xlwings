import xlwings as xw

import matplotlib.pyplot as plt
fig = plt.figure()

plt.plot([1,2,3,1,4,23])
xw.show(fig)

plt.plot([-1,1,-2,2,-3,3])
xw.show(fig)

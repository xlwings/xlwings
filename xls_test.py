import numpy as np
import xlwings


xl = xlwings.Xl()

def rand_numbers():
    """ produces a standard normally distributed random numbers with dim (n,n)"""
    n = xl.get_cell('Sheet1', 1, 2)
    rand_num = np.random.randn(n,n)
    xl.set_range('Sheet1', 3, 3, rand_num)
    
if __name__ == '__main__':
    rand_numbers()
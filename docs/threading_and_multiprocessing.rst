Threading and Multiprocessing
=============================

.. versionadded:: 0.13.0

Threading
---------

While xlwings is not technically thread safe, it's still easy to use it in threads as long as you have at least v0.13.0
and stick to a simple rule: Do not pass xlwings objects to threads. This rule isn't a requirement on macOS, but it's 
still recommended if you want your programs to be cross-platform.


Consider the following example that will **NOT** work::

    import threading
    from queue import Queue
    import xlwings as xw
    
    num_threads = 4
    
    
    def write_to_workbook():
        while True:
            myrange = q.get()
            myrange.value = myrange.address
            print(myrange.address)
            q.task_done()
    
    
    q = Queue()
    
    for i in range(num_threads):
        t = threading.Thread(target=write_to_workbook)
        t.daemon = True
        t.start()
    
    for cell in ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9', 'A10']:
        # THIS DOESN'T WORK - passing xlwings objects to threads will fail!
        myrange = xw.Book('Book1.xlsx').sheets[0].range(cell)
        q.put(myrange)
    
    q.join()


To make it work, you simply have to fully qualify the cell reference in the thread instead of passing a ``Book`` object::


    import threading
    from queue import Queue
    import xlwings as xw
    
    num_threads = 4
    
    
    def write_to_workbook():
        while True:
            cell_ = q.get()
            xw.Book('Book1.xlsx').sheets[0].range(cell_).value = cell_
            print(cell_)
            q.task_done()
    
    
    q = Queue()
    
    for i in range(num_threads):
        t = threading.Thread(target=write_to_workbook)
        t.daemon = True
        t.start()
    
    for cell in ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9', 'A10']:
        q.put(cell)
    
    q.join()


Multiprocessing
---------------

.. note::
    Multiprocessing is only supported on Windows!

The same rules apply to multiprocessing as for threading, here's a working example::


    from multiprocessing import Pool
    import xlwings as xw
    
    
    def write_to_workbook(cell):
        xw.Book('Book1.xlsx').sheets[0].range(cell).value = cell
        print(cell)
    
    
    if __name__ == '__main__':
        with Pool(4) as p:
            p.map(write_to_workbook,
                  ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9', 'A10'])

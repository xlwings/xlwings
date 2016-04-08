.. _matplotlib:

Matplotlib
==========

:meth:`xlwings.Plot` allows for an easy integration of Matplotlib with Excel. The plot
is pasted into Excel as picture.

Getting started
---------------

The easiest sample boils down to::

    >>> import matplotlib.pyplot as plt
    >>> import xlwings as xw

    >>> fig = plt.figure()
    >>> plt.plot([1, 2, 3])

    >>> wb = xw.Workbook()
    >>> xw.Plot(fig).show('MyPlot')

.. figure:: images/mpl_basic.png
  :scale: 80%

.. note::
    You can now resize and position the plot on Excel: subsequent calls to ``show``
    with the same name (``'MyPlot'``) will update the picture without changing its position or size.


Full integration with Excel
---------------------------

Calling the above code with :ref:`RunPython <run_python>` and binding it e.g. to a button is
straightforward and works cross-platform.

However, on Windows you can make things feel even more integrated by setting up a
:ref:`UDF <udfs>` along the following lines::

    @xw.func
    def myplot(n):
        wb = xw.Workbook.caller()
        fig = plt.figure()
        plt.plot(range(int(n)))
        xw.Plot(fig).show('MyPlot')
        return 'Plotted with n={}'.format(n)

If you import this function and call it from cell B2, then the plot gets automatically
updated when cell B1 changes:

.. figure:: images/mpl_udf.png
  :scale: 80%

Properties
----------

Size, position and other properties can either be set as arguments within ``show``, see :meth:`xlwings.Plot.show`, or
by manipulating the picture object as returned by ``show``, see :meth:`xlwings.Picture`.

For example::

    >>> xw.Plot(fig).show('MyPlot', left=xw.Range('B5').left, top=xw.Range('B5').top)

or::

    >>> plot = xw.Plot(fig).show('MyPlot')
    >>> plot.height /= 2
    >>> plot.width /= 2

.. note:: Once the picture is shown in Excel, you can only change it's properties via the picture object and not within
    the ``show`` method.

Getting a matplotlib figure
---------------------------
Here are a few examples of how you get a matplotlib ``figure`` object:

* via PyPlot interface::

    import matplotlib.pyplot as plt
    fig = plt.figure()
    plt.plot([1, 2, 3, 4, 5])

  or::

    import matplotlib.pyplot as plt
    plt.plot([1, 2, 3, 4, 5])
    fig = plt.gcf()


* via object oriented interface::

    from matplotlib.figure import Figure
    fig = Figure(figsize=(8, 6))
    ax = fig.add_subplot(111)
    ax.plot([1, 2, 3, 4, 5])

* via Pandas::

    import pandas as pd
    import numpy as np

    df = pd.DataFrame(np.random.rand(10, 4), columns=['a', 'b', 'c', 'd'])
    ax = df.plot(kind='bar')
    fig = ax.get_figure()

Then show it in Excel as picture as seen above::

    plot = Plot(fig)
    plot.show('Plot1')
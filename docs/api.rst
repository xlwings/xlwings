API Documentation
=================

The xlwings object model is very similar to the one used by Excel VBA but the hierarchy is flattened. An example:

**VBA:**

.. code-block:: python

    Workbooks("Book1").Sheets("Sheet1").Range("A1").Value = "Some Text"

**xlwings:**

.. code-block:: python

    wb = Workbook("Book1")
    Range("Sheet1", "A1").value = "Some Text"

Top-level functions
-------------------

.. automodule:: xlwings
    :members: view

Application
-----------

.. autoclass:: Application
    :members:

Workbook
--------

In order to use xlwings, instantiating a workbook object is always the first thing to do:


.. autoclass:: Workbook
    :members:


.. _api_sheet:

Sheet
-----

Sheet objects allow you to interact with anything directly related to a Sheet.


.. autoclass:: Sheet
    :members:


Range
-----

The xlwings Range object represents a block of contiguous cells in Excel.


.. autoclass:: Range
    :members:

Shape
-----

.. autoclass:: Shape
    :members:

Chart
-----

.. note:: The chart object is currently still lacking a lot of important methods/attributes.


.. autoclass:: Chart
    :members:
    :inherited-members:
    :show-inheritance:
    :exclude-members: __delattr__, __format__, __getattribute__, __hash__, __reduce__, __reduce_ex__, __setattr__, __sizeof__, __str__

Picture
-------

.. autoclass:: Picture
    :members:
    :inherited-members:
    :show-inheritance:
    :exclude-members: __delattr__, __format__, __getattribute__, __hash__, __reduce__, __reduce_ex__, __setattr__, __sizeof__, __str__, __repr__

Plot
----

.. autoclass:: Plot
    :members:
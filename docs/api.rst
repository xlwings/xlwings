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

App
---

.. autoclass:: App
    :members:

Book
----

In order to use xlwings, instantiating a workbook object is always the first thing to do:


.. autoclass:: Book
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


Picture
-------

.. autoclass:: Picture
    :members:
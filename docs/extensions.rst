.. _extensions:

Extensions
==========

It's easy to extend the xlwings add-in with own code like UDF wrapper code or RunPython macros, so that they can be deployed
without end users having to import or write the functions themselves. Just add another VBA module to the xlwings addin
with the respective code.

In-Excel SQL
------------

The xlwings addin comes with a built-in extension that adds in-Excel SQL syntax (sqlite dialect). As soon as you add
a reference to the xlwings add-in (see :ref:`xlwings_addin`), you can use the ``sql`` function:

.. code::

    =sql(SQL Statement, table a, table b, ...)

.. figure:: images/sql.png
    :scale: 40%


As this extension uses UDFs, it's only available on Windows right now.
Extensions
==========

It's easy to extend the xlwings add-in with own code like UDFs or RunPython macros, so that they can be deployed
without end users having to import or write the functions themselves. Just add another VBA module to the xlwings addin
with the respective code.

UDF extensions can be used from every workbook without having to set a reference. 

In-Excel SQL
------------

The xlwings addin comes with a built-in extension that adds in-Excel SQL syntax (sqlite dialect):

.. code::

    =sql(SQL Statement, table a, table b, ...)

.. figure:: ./images/sql.png

As this extension uses UDFs, it's only available on Windows right now.

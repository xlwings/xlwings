.. _r_and_julia:

xlwings with R and Julia
========================

While xlwings is a pure Python package, there are cross-language packages that allow
for a relative straightforward use from/with other languages. This means, however, that you'll
always need to have Python with xlwings installed in addition to R or Julia. We recommend the
`Anaconda <https://store.continuum.io/cshop/anaconda/>`_ distribution, see also :ref:`installation`.

R
-
The R instructions are for Windows, but things work accordingly on Mac except that calling the R functions
as User Defined Functions is not supported at the moment (but ``RunPython`` works, see :ref:`run_python`).


Setup:

* Install R and Python
* Add ``R_HOME`` environment variable to base directory of installation, .e.g ``C:\Program Files\R\R-x.x.x``
* Add ``R_USER`` environment variable to user folder, e.g. ``C:\Users\<user>``
* Add ``C:\Program Files\R\R-x.x.x\bin`` to ``PATH``
* Restart Windows because of the environment variables (!)

Simple functions with R
***********************

Original R function that we want to access from Excel (saved in ``r_file.R``):

.. code::

    myfunction <- function(x, y){
        return(x * y)
    }


Python wrapper code:

.. code::

    import xlwings as xw
    import rpy2.robjects as robjects
    # you might want to use some relative path or place the file in R's current working dir
    robjects.r.source(r"C:\path\to\r_file.R")

    @xw.func
    def myfunction(x, y):
        myfunc = robjects.r['myfunction']
        return tuple(myfunc(x, y))

After importing this function (see: :ref:`udfs`), it will be available as UDF from Excel.

Array functions with R
**********************

Original R function that we want to access from Excel (saved in ``r_file.R``):

.. code::

    array_function <- function(m1, m2){
      # Matrix multiplication
      return(m1 %*% m2)
    }


Python wrapper code:

.. code::

    import xlwings as xw
    import numpy as np
    import rpy2.robjects as robjects
    from rpy2.robjects import numpy2ri

    robjects.r.source(r"C:\path\to\r_file.R")
    numpy2ri.activate()

    @xw.func
    @xw.arg("x", np.array, ndim=2)
    @xw.arg("y", np.array, ndim=2)
    def array_function(x, y):
        array_func = robjects.r['array_function']
        return np.array(array_func(x, y))

After importing this function (see: :ref:`udfs`), it will be available as UDF from Excel.

Julia
-----

Setup:

* Install Julia and Python
* Run ``Pkg.add("PyCall")`` from an interactive Julia interpreter

xlwings can then be called from Julia with the following syntax (the colons take care of
automatic type conversion):

.. code:: julia

    julia> using PyCall
    julia> @pyimport xlwings as xw

    julia> xw.Workbook()
    PyObject <Workbook 'Workbook1'>

    julia> xw.Range("A1")[:value] = "Hello World"
    julia> xw.Range("A1")[:value]
    "Hello World"


    julia> xw.Range("A1")[:value] = [1 2; 3 4]
    julia> xw.Range("A1")[:table][:value]
    2x2 Array{Int64,2}:
    1  2
    3  4




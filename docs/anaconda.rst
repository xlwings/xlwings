.. _anaconda:

Anaconda
=========

**Currently only works in Windows.**

xlwings can execute Python code within an Anaconda environment. You just need to tell xlwings that you are using Anaconda, where the activate script is kept, and which environment you prefer to use.

* ``Use Anaconda``: Check this if you are using a named Anaconda environment. The Anaconda Activate File and Environment Name fields don't work unless this is checked.
* ``Activate File``: This is the directory of the activate.bat file within the condabin of your Anaconda. Example: C:\Users\me\AppData\Local\Continuum\anaconda3\condabin\activate.bat
* ``Environment Name``: The name of the Anaconda environment you want to activate.

================================================================================
pyexcel-io - Let you focus on data, instead of file formats
================================================================================

.. image:: https://api.travis-ci.org/pyexcel/pyexcel-io.svg?branch=master
   :target: http://travis-ci.org/pyexcel/pyexcel-io

.. image:: https://codecov.io/github/pyexcel/pyexcel-io/coverage.png
    :target: https://codecov.io/github/pyexcel/pyexcel-io

.. image:: https://readthedocs.org/projects/pyexcel-io/badge/?version=latest
   :target: http://pyexcel-io.readthedocs.org/en/latest/

Known constraints
==================

Fonts, colors and charts are not supported.


**pyexcel-io** provides **one** application programming interface(API) to read
and write the data in excel format, import the data into and export the data
from database. It provides support for csv(z) format, django database and
sqlalchemy supported databases. Its supported file formats are extended to cover
"xls", "xlsx", "ods" by the following extensions:

.. _file-format-list:
.. _a-map-of-plugins-and-file-formats:

.. table:: A list of file formats supported by external plugins

   ================= ======================= ============= ==================
   Package name      Supported file formats  Dependencies  Python versions
   ================= ======================= ============= ==================
   `pyexcel-io`_     csv, csvz [#f1]_, tsv,                2.6, 2.7, 3.3,
                     tsvz [#f2]_                           3.4, 3.5, 3.6
                                                           pypy
   `pyexcel-xls`_    xls, xlsx(read only),   `xlrd`_,      same as above
                     xlsm(read only)         `xlwt`_
   `pyexcel-xlsx`_   xlsx                    `openpyxl`_   same as above
   `pyexcel-xlsxw`_  xlsx(write only)        `XlsxWriter`_ same as above
   `pyexcel-ods3`_   ods                     `ezodf`_,     2.6, 2.7, 3.3, 3.4
                                             lxml          3.5, 3.6
   `pyexcel-ods`_    ods                     `odfpy`_      same as above
   ================= ======================= ============= ==================

.. _pyexcel-io: https://github.com/pyexcel/pyexcel-io
.. _pyexcel-xls: https://github.com/pyexcel/pyexcel-xls
.. _pyexcel-xlsx: https://github.com/pyexcel/pyexcel-xlsx
.. _pyexcel-ods: https://github.com/pyexcel/pyexcel-ods
.. _pyexcel-ods3: https://github.com/pyexcel/pyexcel-ods3
.. _pyexcel-xlsxw: https://github.com/pyexcel/pyexcel-xlsxw

.. _xlrd: https://github.com/python-excel/xlrd
.. _xlwt: https://github.com/python-excel/xlwt
.. _openpyxl: https://bitbucket.org/openpyxl/openpyxl
.. _XlsxWriter: https://github.com/jmcnamara/XlsxWriter
.. _ezodf: https://github.com/T0ha/ezodf
.. _odfpy: https://github.com/eea/odfpy


.. rubric:: Footnotes

.. [#f1] zipped csv file
.. [#f2] zipped tsv file

If you need to manipulate the data, you might do it yourself or use its brother
library `pyexcel <https://github.com/pyexcel/pyexcel>`__ .

If you would like to extend it, you may use it to write your own
extension to handle a specific file format.




Installation
================================================================================
You can install it via pip:

.. code-block:: bash

    $ pip install pyexcel-io


or clone it and install it:

.. code-block:: bash

    $ git clone http://github.com/pyexcel/pyexcel-io.git
    $ cd pyexcel-io
    $ python setup.py install



Development guide
================================================================================

Development steps for code changes

#. git clone https://github.com/pyexcel/pyexcel-io.git
#. cd pyexcel-io

Upgrade your setup tools and pip. They are needed for development and testing only:

#. pip install --upgrade setuptools "pip==7.1"

Then install relevant development requirements:

#. pip install -r rnd_requirements.txt # if such a file exists
#. pip install -r requirements.txt
#. pip install -r tests/requirements.txt


In order to update test environment, and documentation, additional setps are
required:

#. pip install moban
#. git clone https://github.com/pyexcel/pyexcel-commons.git commons
#. make your changes in `.moban.d` directory, then issue command `moban`

What is rnd_requirements.txt
-------------------------------

Usually, it is created when a dependent library is not released. Once the dependecy is installed(will be released), the future version of the dependency in the requirements.txt will be valid.

What is pyexcel-commons
---------------------------------

Many information that are shared across pyexcel projects, such as: this developer guide, license info, etc. are stored in `pyexcel-commons` project.

What is .moban.d
---------------------------------

`.moban.d` stores the specific meta data for the library.

How to test your contribution
------------------------------

Although `nose` and `doctest` are both used in code testing, it is adviable that unit tests are put in tests. `doctest` is incorporated only to make sure the code examples in documentation remain valid across different development releases.

On Linux/Unix systems, please launch your tests like this::

    $ make test

On Windows systems, please issue this command::

    > test.bat

License
================================================================================

New BSD License

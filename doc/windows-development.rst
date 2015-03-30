Testing on Windows
==================


Although openpyxl itself is pure Python and should run on any Python, we do use some libraries that require compiling for tests and documentation. The setup for testing on Windows is somewhat different.


Getting started
---------------

Once you have installed the versions of Python (2.6, 2.7, 3.3, 3.4) you should setup a development environment for testing so that you do not adversely affect the system install.


Setting up a development environment
------------------------------------

First of all you should checkout a copy of the repository. Atlassian provides a nice GUI client `SourceTree <http://www.sourcetreeapp.com>`_ that allows you to do this with a single-click from the browser.

By default the repository will be installed under your user folder. eg. c:\Users\YOURUSER\openpyxl

Switch to the branch you want to work on by double-clicking it. The default branch should never be used for development work.

Creating a virtual environment
++++++++++++++++++++++++++++++

You will need to manually install virtualenv. This is best done by first installing pip. open a command line and download the script "get_pip.py" to your preferred Python folder::

    bitsadmin /transfer pip http://bootstrap.pypa.io/get-pip.py c:\python27\get-pip.py # change the path as necessary
    
Install pip (it needs to be at least pip 6.0)::
  
    python get_pip.py

Now you can install virtualenv::

    Scripts\pip install virtualenv
    Scripts\virtualenv c:\Users\YOURUSER\openpyxl

    
lxml
----

openpyxl needs `lxml` in order to run the tests. Unfortunately, automatic installation of lxml on Windows is tricky as pip defaults to try and compile it. This can be avoided by using pre-compiled versions of the library.

#. In the command line switch to your repository folder::

    cd c:\Users\YOURUSER\openpyxl
  
#. Activate the virtualenv::

    Scripts\activate

#. Install a development version of openpyxl::

    python setup.py develop

#. Download all the relevant `lxml Windows wheels <http://www.lfd.uci.edu/~gohlke/pythonlibs/#lxml>`_

#. Move all these files to a folder called "downloads" in your openpyxl checkout

#. Install the project requirements::

    pip install --download downloads -r requirements.txt
    pip install --no-index --find-links downloads -r requirements.txt

To run tests for the virtualenv::

    py.test -xrf openpyxl # the flag will stop testing at the first error

    
tox
---

We use `tox` to run the tests on different Python versions and configurations. Using it is as simple as::

    set PIP_FIND_LINKS=downloads
    tox openpyxl

Testing on Windows
==================


Although openpyxl itself is pure Python and should run on any Python, we do use some libraries that require compiling for tests and documentation. The setup for testing on Windows is somewhat different.


Getting started
---------------

Once you have installed the versions of Python (2.6, 2.7, 3.3, 3.4) you should setup virtual environments for testing so that you do not pollute the system install.


Setting up virtual environments
-------------------------------

First of all you should checkout a copy of the repository. Atlassian provides a nice GUI client SourceTree that allows you to do this with a single-click from the browser.

By default the repository will be installed under the Documents folder of your user. eg. c:\Users\YOURUSER\openpyxl


Python 2.6 & Python 2.7
+++++++++++++++++++++++

You will need to manually install virtualenv. This is best done by first installing pip. Download the script "get_pip.py" and copy it to the main folders: c:\python26 and c:\python27

In a DOS-box switch to the relevant folder and run::

cd c:\python26
python get_pip.py

This will install pip. Now you can install virtualenv

Check the version, it needs to be at least pip 6.0

scripts\pip install virtualenv
python Scripts\pyvenv.py c:\Users\YOURUSER\openpyxl


Python 3.3
++++++++++

In a DOS-box switch to the Python 3.3 folder and run::

cd c:\python33
python Tools\Scripts\pyvenv.py c:\Users\YOURUSER\openpyxl


Python 3.4
++++++++++

cd c:\python34
python Tools\Scripts\pyvenv.py c:\Users\YOURUSER\openpyxl


lxml
----

openpyxl needs `lxml` in order to run the tests. Unfortunately, automatic installation of lxml on Windows is tricky as pip defaults to try and compile it.

#. In a DOS-box switch to your repository follder::
cd c:\Users\YOURUSER\openpyxl

#. Activate the virtualenv::
scripts\activate

#. Install a development version of openpyxl::
python setup.py develop

#. Download all the relevant `lxml Windows installers from PyPI <https://pypi.python.org/pypi/lxml>`_

#. Move all these files to a folder called "downloads" in your openpyxl checkout

#. Install `wheel`::

pip install wheel

#. Convert the files to "wheels": Scripts\wheel convert --dest-dir downloads downloads\*.exe

You can now install lxml into any virtualenv using the following command::

#. Install the project requirements::

pip -U -r requirements.txt

To run tests for the virtualenv::

py.test -xrf openpyxl # the flag will stop tests at the first error so you're not overwhelmed


tox
---

We use `tox` to run the tests on different Python versions and configurations. Using it is as simple as::

tox openpyxl

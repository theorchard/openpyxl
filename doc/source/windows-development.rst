Testing on Windows
==================


Although openpyxl itself is pure Python and should run on any Python, we do use some libraries that require compiling for tests and documentation. As tox support on Windows is also limited the setup for testing on Windows is somewhat different.


Getting started
---------------


On Windows you must use the same compiler for the libraries as was used to create Python. In general, if you're using a binary release from python.org this will be Visual Studio. Microsoft provides free versions of Visual Studio which you can use to compile any libraries necessary. However, as not all architectures are supported it is recommended that you work only with the 32-bit versions of Python. As they work fine on 64-bit Windows this is not a real problem.

Python 2.6 and Python 2.7 were compiled with Visual Studio 2008 aka Visual Studio 9
http://download.microsoft.com/download/E/8/E/E8EEB394-7F42-4963-A2D8-29559B738298/VS2008ExpressWithSP1ENUX1504728.iso

Python 3.3 and Python 3.4 were compiled with Visual Studio 2010 aka Visual Studio 10
http://www.microsoft.com/download/en/details.aspx?id=23691


These are provided as ISO images so you will either have to create DVDs from them or use something like WinCDEmu to mount them. If you are using a Windows virtual machine then you may need to copy the ISOs to the Windows drive in order to be able to mount them.


Once you have installed the versions of Python and the relevant versions of VisualStudio you should setup virtual environments for testing so that you do not pollute the system install.


Setting up virtual environments
-------------------------------


First of all you should checkout a copy of the repository. Atlassian provides a nice GUI client SourceTree that allows you to do this with a single-click from the browser.

By default the repository will be installed under the Documents folder of your user. eg. c:\Users\YOURUSER\openpyxl


Python 2.6 & Python 2.7
+++++++++++++++++++++++


You will need to manually install virtualenv. This is best done by first installing pip. Download the script "get_pip.py" and copy it to the main folders: c:\python26 and c:\python27

In a DOS-box switch to the relevant folder and run::

python get_pip.py

This will install pip. Now you can install virtualenv

scripts\pip install virtualenv


Python 3.3
++++++++++

In a DOS-box switch to the Python 3.3 folder and run::

cd c:\python33
python Tools\Scripts\pyvenv.py c:\Users\YOURUSER\openpyxl33


Python 3.4
++++++++++


lxml
-------------

openpyxl needs lxml in order to run the tests. Unfortunately, automatic installation of lxml on Windows is tricky as pip defaults to try and compile it. There is a workaround for this.

1. Download all the relevant lxml Windows installers from PyPI
2. Move all these files to a specific directory such as c:\lxml-downloads
3. Install wheel: pip install wheel
3. Convert the files to "wheels": Scripts\wheel convert --dest-dir c:\lxml-downloads c:\lxml-downloads\*.exe

You can now install lxml into any virtualenv using the following command::

pip install -U --no-index --find-index=c:\downloads lxml

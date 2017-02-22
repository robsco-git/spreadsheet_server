# spreadsheet_server

## Introduction

spreadsheet_server was built to aid rapid web tool development where the logic
was already implemented in Microsoft Excel/LibreOffice Calc. Instead of
rewriting the logic from scratch, this tool was born.

The tool has been developed to work on a headless GNU/Linux or FreeBSD system 
where the server and client are on the same machine.

## Features

* 'Instant' access to cells in the spreadsheets as they open in LibreOffice Calc.
* All the function calculation support and power of LibreOffice Calc.
* A given spreadsheet is locked (within python, not on disk) when it is accessed to prevent state irregularities across multiple concurrent connections to the same spreadsheet.
* Monitoring of a directory with automatic loading and unloading of spreadsheets.
* By default, when a spreadsheet file changes on disk, it will be closed and
  opened in LibreOffice.
* Spreadsheets can be saved - useful for debugging purposes.

## Usage

Place your spreadsheets in './spreadsheets'.

Make sure you have a correctly set up virtualenv the required packages are installed
(see below), then run the server:
```
python server.py
```

You can also create a SpreadsheetServer object (from server.py) and then call
the method 'run' on the object to start the server. This allows the for
customisation of the default settings. Have a look in server.py to see what can
be set.

example_client.py is provided for an overview of how to use the functions
exposed by the client. This is a good place to start.

## How it works

* A LibreOffice instance is launched by 'server.py' in a headless state.
* By default, the './spreadsheets' directory is polled every 5 seconds for file
changes.
* New spreadsheets in the directory are opened with LibreOffice and removed
spreadsheets are closed in LibreOffice.
* The 'client.py' connects to the server and can update cells and retrieve
their calculated content.

## Requirements

### Server

* LibreOffice Calc
* Python3 or Python2
  * On Ubuntu Server, Python3 is recommended. On FreeBSD, Python2 is recommended. See the installation instructions below.
* Python packages:
  * pyoo (for the connection to LibreOffice)
  * psutil (for cross platform pid discovery for the soffice process)
  * werkzeug (for securing filenames when saving spreadsheets)

### Client

* Python2 or Python3

## Installation

### Ubuntu Server 14.04

```
sudo apt-get update && sudo apt-get upgrade
sudo apt-get install git python3 python3-dev python-virtualenv libreoffice-calc python3-uno
git clone https://github.com/robsco-git/spreadsheet_server.git
cd spreadsheet_server
virtualenv --system-site-packages -p python3 venv
. venv/bin/activate
pip install -r requirements.txt
python server.py
```

### Ubuntu Server 16.04

The same steps as for 14.04 and in addition you need GCC to build psutil
```
sudo apt-get install gcc
```

#### Running the server component with python2 on Ubuntu Server

First off, keep in mind that a python2 spreadsheet_server client can connect to a
python3 spreadsheet_server server.

The Ubuntu team has removed the 'python-uno' package in favour of the 'python3-uno'
package. Running the spreadsheet_server server component with python2 requires
the LibreOffice 'uno' module built for python2. If you really need to run the server
as python2 on Ubuntu Server, You will need to build LibreOffice with python2
support. Here is a script outlining how to do so for an older version of LibreOffice:
https://gist.github.com/hbrunn/6f4a007a6ff7f75c0f8b

Also, here is some more discussion on the topic:
http://askubuntu.com/a/418550

### FreeBSD 11.0

#### Using LibreOffice installed from pkg

At the time of writing, the Python-UNO module, is built using python2 when you
install the binary version of LibreOffice provided by pkg. If you decide to use the
pkg version of LibreOffice, you will need to run the spreadsheet_server server
component with python2:
```
freebsd-update fetch install
sudo pkg install git libreoffice py27-virtualenv
git clone https://github.com/robsco-git/spreadsheet_server.git
cd spreadsheet_server
virtualenv venv
. venv/bin/activate
pip install -r requirements.txt
python server.py
```

#### Using LibreOffice installed from the ports tree, compiling the LibreOffice python modules with python3.5

```
freebsd-update fetch install
pkg install git py27-virtualenv
cd /usr/ports/editors/libreoffice
setenv PYTHON_VERSION python3.5
make clean install
git clone https://github.com/robsco-git/spreadsheet_server.git
cd spreadsheet_server
virtualenv venv
. venv/bin/activate
pip install -r requirements.txt
python server.py
```

## Questions

What is UNO? What is Python-UNO? What is PyOO?
The first few paragraphs of the PyOO README should answer most of your questions:
https://github.com/seznam/pyoo/blob/master/README.rst

## Tests

You can run the all the current tests with 'python -m unittest discover'.
Use './coverage.sh' to run the coverage analysis of the current tests and have a look
in the generated htmlcov directory.
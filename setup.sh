#!/bin/bash

set -e

# install required packages
sudo apt-get update
sudo apt-get install python3-uno libreoffice-calc
pip install virtualenv

# set up the python3 virtualenv
virtualenv --system-site-packages -p python3 venv

# install pyoo
. venv/bin/activate
pip3 install pyoo
deactivate

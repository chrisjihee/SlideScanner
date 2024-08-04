#!/bin/bash
# basic
pip install -r requirements.txt

# chrisbase
rm -rf chrisbase*
git clone https://github.com/chrisjihee/chrisbase.git
pip install --editable chrisbase*

# list
pip list | grep -E "chris|pptx"

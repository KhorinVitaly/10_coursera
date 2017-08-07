# Coursera Dump

With this script you can save courses list from coursera.org into xlsx file

# How to Install

Python 3 should be already installed. Then use pip (or pip3 if there is a conflict with old Python 2 setup) to install dependencies:

```bash
pip install -r requirements.txt # alternatively try pip3
```

Remember, it is recommended to use [virtualenv/venv](https://devman.org/encyclopedia/pip/pip_virtualenv/) for better isolation.

# Quickstart

Launch script in terminal, fill parameters -f = file path; -s = courses list size limit <optional, 
default value 20>  

```bash
python3 coursera.py [-f test.xlsx] <-s 15>
Courses list wrote successfully into test.xlsx
```

# Project Goals

The code is written for educational purposes. Training course for web-developers - [DEVMAN.org](https://devman.org)

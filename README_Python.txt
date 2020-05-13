## Prerequisites

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. 

Running the program is easy, however there are several prerequisites to getting started. Please see down below. 

To begin, git clone the project onto your local machine. This can be done by copying the [github link]("https://github.com/rlandesman/GreenAmmoniaProject.git") (found on the website or the hyperlink) and typing the following into your terminal

```
git clone <link>
```

### Installing

There are several software packages that are neccessary for the successful deployment of this software.

First, Python3 is a must for running the script. Please [install Python3.x ](https://www.python.org/downloads/ "Python Download Page") for your specific operating system 

Next, you will need to install the appropriate support libraries. In your command line, please enter the directory into which you installed the project
Then enter the following commands

```
pip install PyYAML
pip install openpyxl
pip install tqdm
```

NOTE: All user-requested parameters can conveniently be found in a file titled config.yaml. For testing purposes, these values have been pre-determined, but **should** be  customized by changing the value fields (strings only) inside this file for non-demo production use.

## Program Execution
To execute the script, type into your command line the following instruction
I manually turn off warnings, you could keep it but they are useless and mess
with the progress bar :) 

```
python3 -Wd main.py
```

The output will be found in the (now modified) excel sheets the user provided

## More information

For readability and documentation purposes the Python code was developed accoridng to the steps laid out by Professor O'Connells instructions. 

## Built With

* [PyYAML](https://pyyaml.org/wiki/PyYAMLDocumentation) - User-Input Parser
* [openpyxl](https://openpyxl.readthedocs.io/en/stable/) - MS Excel Python library
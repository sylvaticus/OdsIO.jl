# OdsIO

ODS I/O for Julia Dictionaries or DataFrames using the python ezodf module.

[![Build Status](https://travis-ci.org/sylvaticus/OdsIO.jl.svg?branch=master)](https://travis-ci.org/sylvaticus/OdsIO.jl)

[![Coverage Status](https://coveralls.io/repos/sylvaticus/OdsIO.jl/badge.svg?branch=master&service=github)](https://coveralls.io/github/sylvaticus/OdsIO.jl?branch=master)

[![codecov.io](http://codecov.io/github/sylvaticus/OdsIO.jl/coverage.svg?branch=master)](http://codecov.io/github/sylvaticus/OdsIO.jl?branch=master)


### Installation
Untill this (experimental!) package is not registered in the official Julia package repository you can install it with:

`Pkg.clone("git://github.com/sylvaticus/OdsIO.jl.git")`


This package provides the following functions:

### ODS reading:
- ods2dics(filename;sheetsNames=[],sheetsPos=[],ranges=[])
- ods2dfs(filename;sheetsNames=[],sheetsPos=[],ranges=[])

sheetsNames (optional) are a list of sheet names from which to import data  
sheetsPos (optional) aree a list of sheet positions (starting from 1) from which to import data  
sheetsNames and sheetsPos can not be given together  
ranges (optional) is a list of pair of touples defining the ranges in each sheet to import ((tlr,trc),(brr,brc))  
The functions return a list of dictionaries or DataFrames indexed by position or names (this is the reason of the `s` in the function name)

The following functions are provided for convenience:
- ods2dic(filename; sheetName=Null,sheetPos=Null,range=())
- ods2df(filename;sheetName=Null,sheetPos=Null,range=())

Where sheetName, sheetPos and range are scalars and return directly a single dic/df. If none of sheetName or sheetPos are provided the function return the data of the first sheet.

### ODS writing (NOT YET IMPLEMENTED):
- dic2ods(dic, filename; topLefts=[])
- df2ods(df, filename; topLefts=[])

dic or df are respectively a list of dictionaries or DataFrames. Dic or df will be written to the ods sheets by names if the keys/headers are string, otherwise they will be written by position.
topLefts (optional) are a list of touples defining the top-left corner to where to write the data. Default to (1,1) on each sheet.  


The following functions are provided by convenience:
- dic2ods(dic, filename; topLeft=(1,1))
- df2ods(df, filename; topLeft=(1,1))

Where dic and df are single dic/df. 

### Requirements

This package requires:
- the [PyCall](https://github.com/JuliaPy/PyCall.jl) module to call Python
- a working local installation of Python with the python [ezodf](https://github.com/T0ha/ezodf) module already installed
- the [DataFrames](https://github.com/JuliaStats/DataFrames.jl) package if one want to work with DataFrames.

### Known limitations

As the data is saved in a dictionary, the order of the columns is not maintained.

It may be slow with very large data.
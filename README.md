# odsio
ODS I/O for Julia Dict or DF using the python ezodf module 

(EMPTY package. Doesn't do anything)

This package will provide the following functions:

### ODS reading:
- ods2dic(filename;sheetsNames=[],sheetsPos=[],ranges=[])
- ods2df(filename;sheetsNames=[],sheetsPos=[],ranges=[])

sheetsNames (optional) are a list of sheet names from which to import data  
sheetsPos (optional) aree a list of sheet positions (starting from 1) from which to import data  
sheetsNames and sheetsPos can not be given together  
ranges (optional) is a list of pair of touples defining the ranges in each sheet to import ((tlr,trc),(brr,brc))  
The functions return a list of dictionaries or DataFrames indexed by position or names

The following functions are provided for convenience:
- ods2dic(filename; sheetName=Null,sheetPos=Null,range=())
- ods2df(filename;sheetName=Null,sheetPos=Null,range=())

Where sheetName, sheetPos and range are scalars and return directly a single dic/df. One of sheetName or sheetPos must be provided or the array version of the funciton will be called instead.

### ODS writing:
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


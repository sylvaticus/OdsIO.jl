# OdsIO

ODS I/O for Julia Dictionaries or DataFrames using the python ezodf module.

[![Build Status](https://travis-ci.org/sylvaticus/OdsIO.jl.svg?branch=master)](https://travis-ci.org/sylvaticus/OdsIO.jl)

[![Coverage Status](https://coveralls.io/repos/sylvaticus/OdsIO.jl/badge.svg?branch=master&service=github)](https://coveralls.io/github/sylvaticus/OdsIO.jl?branch=master)

[![codecov.io](http://codecov.io/github/sylvaticus/OdsIO.jl/coverage.svg?branch=master)](http://codecov.io/github/sylvaticus/OdsIO.jl?branch=master)


## Installation
Untill this (experimental!) package is not yet registered in the official Julia package repository you can install it with:

`Pkg.clone("git://github.com/sylvaticus/OdsIO.jl.git")`


This package provides the following functions:

## ODS reading:

### ods_readall()

    ods_readall(filename; <keyword arguments>)

Return a dictionary of tables|dictionaries|dataframes indexed by position or name in the original OpenDocument Spreadsheet (.ods) file.

#### Arguments
* `sheetsNames=[]`: the list of sheet names from which to import data.
* `sheetsPos=[]`: the list of sheet positions (starting from 1) from which to import data.
* `ranges=[]`: a list of pair of touples defining the ranges in each sheet from which to import data, in the format ((tlr,trc),(brr,brc))
* `innerType="Matrix"`: the type of the inner container returned. Either "Matrix", "Dict" or "DataFrame"

#### Notes
* sheetsNames and sheetsPos can not be given together
* ranges is defined using integer positions for both rows and columns
* individual dictionaries or dataframes are keyed by the values of the cells in the first row specified in the range, or first row if `range` is not given
* innerType="Matrix", differently from innerType="Dict", preserves original column order, it is faster and require less memory
* using innerType="DataFrame" requires the package `DataFrames` and also preserves original column order

#### Examples
```julia
julia> outDic  = ods2dics("spreadsheet.ods";sheetsPos=[1,3],ranges=[((1,1),(3,3)),((2,2),(6,4))], innerType="Dict")
Dict{Any,Any} with 2 entries:
  3 => Dict{Any,Any}(Pair{Any,Any}("c",Any[33.0,43.0,53.0,63.0]),Pair{Any,Any}("b",Any[32.0,42.0,52.0,62.0]),Pair{Any,Any}("d",Any[34.0,44.0,54.…
  1 => Dict{Any,Any}(Pair{Any,Any}("c",Any[23.0,33.0]),Pair{Any,Any}("b",Any[22.0,32.0]),Pair{Any,Any}("a",Any[21.0,31.0]))
```


### ods_read()

    ods_read(filename; <keyword arguments>)

Return a  table|dictionary|dataframe from a sheet (or range within a sheet) in a OpenDocument Spreadsheet (.ods) file..

#### Arguments
* `sheetName=nothing`: the sheet name from which to import data.
* `sheetPos=nothing`: the position of the sheet (starting from 1) from which to import data.
* `ranges=[]`: a pair of touples defining the range in the sheet from which to import data, in the format ((tlr,trc),(brr,brc))
* `retType="Matrix"`: the type of container returned. Either "Matrix", "Dict" or "DataFrame"

#### Notes
* sheetName and sheetPos can not be given together
* if both sheetName and sheetPos are not specified data from the first sheet is returned
* ranges is defined using integer positions for both rows and columns
* the dictionary or dataframe is keyed by the values of the cells in the first row specified in the range, or first row if `range` is not given
* retType="Matrix", differently from innerType="Dict", preserves original column order, it is faster and require less memory
* using retType="DataFrame" requires the package `DataFrames` and also preserves original column order

#### Examples
```julia
julia> df = ods_read("spreadsheet.ods";sheetName="Sheet2",retType="DataFrame")
3×3 DataFrames.DataFrame
│ Row │ x1   │ x2   │ x3   │
├─────┼──────┼──────┼──────┤
│ 1   │ "a"  │ "b"  │ "c"  │
│ 2   │ 21.0 │ 22.0 │ 23.0 │
│ 3   │ 31.0 │ 32.0 │ 33.0 │
```


## ODS writing (NOT YET IMPLEMENTED. SCHEDULED FOR FEB 2017):
- dic2ods(dic, filename; topLefts=[])
- df2ods(df, filename; topLefts=[])

dic or df are respectively a list of dictionaries or DataFrames. Dic or df will be written to the ods sheets by names if the keys/headers are string, otherwise they will be written by position.
topLefts (optional) are a list of touples defining the top-left corner to where to write the data. Default to (1,1) on each sheet.  


The following functions are provided by convenience:
- dic2ods(dic, filename; topLeft=(1,1))
- df2ods(df, filename; topLeft=(1,1))

Where dic and df are single dic/df. 
 
## Testing

    odsio_test()

Provide a test to check that both the Julia 'OdsIO' and Python 'ezodf' modules are correctly installed.


## Requirements

This package requires:
- the [PyCall](https://github.com/JuliaPy/PyCall.jl) module to call Python
- a working local installation of Python with the python [ezodf](https://github.com/T0ha/ezodf) module already installed
- the [DataFrames](https://github.com/JuliaStats/DataFrames.jl) package if one want to work with DataFrames.

## Known limitations

* As the data is saved in a dictionary, the order of the columns is not maintained.
* It is relativelly slow with very large data.
* If the data has many columns, the conversion from Dictionary to DataFrame made in the ods2dfs and ods2df functions may not work. In that case call the ods2dics or ods2dic functions and perfom the conversion manually choosing the columns you need.

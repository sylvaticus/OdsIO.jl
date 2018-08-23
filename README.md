# OdsIO

Open Document Format Spreadsheet (ODS) I/O for Julia using the python ezodf module.

It allows to export (import) data from (to) Julia to (from) LibreOffice, OpenOffice and any other spreadsheet software that implements the OpenDocument specifications.

[![Build Status](https://travis-ci.org/sylvaticus/OdsIO.jl.svg?branch=master)](https://travis-ci.org/sylvaticus/OdsIO.jl)
[![codecov.io](http://codecov.io/github/sylvaticus/OdsIO.jl/coverage.svg?branch=master)](http://codecov.io/github/sylvaticus/OdsIO.jl?branch=master)
[![OdsIO](http://pkg.julialang.org/badges/OdsIO_0.6.svg)](http://pkg.julialang.org/?pkg=OdsIO&ver=0.6)
[![OdsIO](http://pkg.julialang.org/badges/OdsIO_1.0.svg)](http://pkg.julialang.org/?pkg=OdsIO&ver=1.0)

## Installation
`Pkg.add("OdsIO")`

This package provides the following functions:

## ODS reading:

### ods_readall()

    ods_readall(filename; <keyword arguments>)

Return a dictionary of tables|dictionaries|dataframes indexed by position or name in the original OpenDocument Spreadsheet (.ods) file.

#### Arguments
* `sheetsNames=[]`: the list of sheet names from which to import data.
* `sheetsPos=[]`: the list of sheet positions (starting from 1) from which to import data.
* `ranges=[]`: a list of pair of touples defining the ranges in each sheet from which to import data, in the format ((tlr,tlc),(brr,brc))
* `innerType="Matrix"`: the type of the inner container returned. Either "Matrix", "Dict" or "DataFrame"

#### Notes
* sheetsNames and sheetsPos can not be given together
* ranges is defined using integer positions for both rows and columns
* individual dictionaries or dataframes are keyed by the values of the cells in the first row specified in the range, or first row if `range` is not given
* innerType="Matrix", differently from innerType="Dict", preserves original column order, it is faster and require less memory
* using innerType="DataFrame" also preserves original column order and try to auto-convert column types (working for Int64, Float64, String, in that order)

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
* `ranges=[]`: a pair of touples defining the range in the sheet from which to import data, in the format ((tlr,tlc),(brr,brc))
* `retType="Matrix"`: the type of container returned. Either "Matrix", "Dict" or "DataFrame"

#### Notes
* sheetName and sheetPos can not be given together
* if both sheetName and sheetPos are not specified data from the first sheet is returned
* ranges is defined using integer positions for both rows and columns
* the dictionary or dataframe is keyed by the values of the cells in the first row specified in the range, or first row if `range` is not given
* retType="Matrix", differently from innerType="Dict", preserves original column order, it is faster and require less memory
* using innerType="DataFrame" also preserves original column order and try to auto-convert column types (working for Int64, Float64, String, in that order)

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


## ODS writing

### ods_write()

    ods_write(filename,data)

Write tabular data (2D Array, DataFrame or Dictionary) to OpenDocument spreadsheet format.

#### Arguments
* `filename`:    an existing ods file or the one to create.
* `data=Dict()`: a dictionary of locations in the files where to export the data => the actual data (see notes).

#### Notes:
* The locations where to save the data (the keys in the dictionary) are a tuple of tree elements:
The first one is the sheet name or sheet position, the other two are the index of row and column of the top
left corner where to export the data.
If using sheet positions, these must be within current file sheets boundaries. If you want to create new sheets,
use names.
* The actual data exported are either a Matrix (2D Array), a DataFrame or an OrderedDict. In case of DataFrame or
OrderedDict the headers ARE exported, so if you don't want them, first convert the DataFrame (or Dictionary)
to a Matrix. In case of OrderedDict, the inner data must all have the same length.
* Some spreadsheet software may not automatically recalculate cells that depends on exported cells (e.g. we are exporting
some data o cell `A1` and cells `A2` depends on `A2`, the content of cell `A2` may not be updated after the export).
In such case most spreadsheet software have a command to force recalculation of cells (e.g. in LibreOffice/OpenOffice
use `CTRL+Shift+F9`)

#### Examples
```julia
julia> ods_write("TestSpreadsheet.ods",Dict(("TestSheet",3,2)=>[[1,2,3,4,5] [6,7,8,9,10]]))
```

## Testing

Pkg.test("OdsIO")

Provide tests to check that both the Julia 'OdsIO' and Python 'ezodf' modules are correctly installed. It may return an error if the file system is not writeable.


## Requirements

This package requires:
- the [PyCall](https://github.com/JuliaPy/PyCall.jl) package to call Python
- a working local installation of Python with the python [ezodf](https://github.com/T0ha/ezodf) module already installed (if the `ezodf` module is not available and you have no access to the local python installation, you can use PyCall to try to install the `ezodf` using pip.. see [here](https://gist.github.com/Luthaf/368a23981c8ec095c3eb))
- the [DataFrames](https://github.com/JuliaStats/DataFrames.jl) package in order to return DataFrames.

## Known limitations

* In reading, as the data is saved in a dictionary, the order of the columns is not maintained.
* It is relatively slow with very large data.
* If the data has many columns, the conversion from Dictionary to DataFrame made in the ods2dfs and ods2df functions may not work. In that case call the ods2dics or ods2dic functions and perform the conversion manually choosing the columns you need.

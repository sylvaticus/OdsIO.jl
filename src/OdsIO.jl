# __precompile__()

module OdsIO

export ods2dics, ods2dic, ods2dfs, ods2df, odsio_test, odsio_autotest
using PyCall

dfPackIsInstalled = true
try
    using DataFrames
catch
    dfPackIsInstalled = false
end


"""
    ods2dics(filename; <keyword arguments>)

Return a dictionary of dictionaries indexed by position or name in the original OpenDocument Spreadsheet (.ods) file.

# Arguments
* `sheetsNames=[]`: the list of sheet names from which to import data.
* `sheetsPos=[]`: the list of sheet positions (starting from 1) from which to import data.
* `ranges=[]`: a list of pair of touples defining the ranges in each sheet from which to import data, in the format ((tlr,trc),(brr,brc))

# Notes
* sheetsNames and sheetsPos can not be given together
* ranges is defined using integer positions for both rows and columns
* individual dictionaries are keyed by the values of the cells in the first row specified in the range, or first row if `range` is not given

# Examples
```julia
julia> outDic  = ods2dics("spreadsheet.ods";sheetsPos=[1,3],ranges=[((1,1),(3,3)),((2,2),(6,4))])
Dict{Any,Any} with 2 entries:
  3 => Dict{Any,Any}(Pair{Any,Any}("c",Any[33.0,43.0,53.0,63.0]),Pair{Any,Any}("b",Any[32.0,42.0,52.0,62.0]),Pair{Any,Any}("d",Any[34.0,44.0,54.…
  1 => Dict{Any,Any}(Pair{Any,Any}("c",Any[23.0,33.0]),Pair{Any,Any}("b",Any[22.0,32.0]),Pair{Any,Any}("a",Any[21.0,31.0]))
```
"""
function ods2dics(filename;sheetsNames=[],sheetsPos=[],ranges=[])
    @pyimport ezodf
    toReturn = Dict()
    doc = ezodf.opendoc(filename)
    nsheets = length(doc[:sheets])
    toReturnKeyType = "name"
    if !isempty(sheetsNames) && !isempty(sheetsPos)
        error("Do not use sheetNames and sheetPos together")
    end
    if !isempty(sheetsPos)
        toReturnKeyType = "pos"
    end
    sheetsCounter=0

    for (is, sheet) in enumerate(doc[:sheets])
        if is in sheetsPos || sheet[:name] in sheetsNames || (isempty(sheetsNames) && isempty(sheetsPos))
            sheetsCounter += 1
            r_min = 1
            r_max = sheet[:nrows]()
            c_min = 1
            c_max = sheet[:ncols]()
            try
                if !isempty(ranges) && !isempty(ranges[sheetsCounter])
                    r_min     = ranges[sheetsCounter][1][1]
                    r_max     = ranges[sheetsCounter][2][1]
                    c_min     = ranges[sheetsCounter][1][2]
                    c_max     = ranges[sheetsCounter][2][2]
                end
            catch
                error("There is a problem with the range. Range should be defined as a list of pair of touples ((tlr,trc),(brr,brc)) for each sheet to import, using integer positions." )
            end
            df_dict   = Dict()
            col_index = Dict()

            for (i, row) in enumerate(sheet[:rows]())
                # row is a list of cells
                if i == r_min # header row
                    for (j,cell) in enumerate(row)
                        if(j>=c_min && j<=c_max)
                            df_dict[cell[:value]] = []
                            col_index[j]=cell[:value]
                        end
                        if j > c_max
                            break
                        end
                    end
                end # end header row
                if (i> r_min && i <= r_max) # data row
                    for (j, cell) in enumerate(row)
                        if(j>=c_min && j<=c_max)
                            # use header instead of column index
                            push!(df_dict[col_index[j]],cell[:value])
                        end
                        if j> c_max
                            break
                        end
                    end
                end # data row
                if i > r_max
                    break
                end
            end # end for each row loop
            if toReturnKeyType == "name"
                toReturn[sheet[:name]] = df_dict
            else
                toReturn[is] = df_dict
            end
        end # end check is a sheet to retain
    end # for each sheet
    return toReturn
end # end functionSS


"""
    ods2dic(filename; <keyword arguments>)

Return a dictionary from a sheet (or range within a sheet) in a OpenDocument Spreadsheet (.ods) file..

# Arguments
* `sheetName=nothing`: the sheet name from which to import data.
* `sheetPos=nothing`: the position of the sheet (starting from 1) from which to import data.
* `ranges=[]`: a pair of touples defining the range in the sheet from which to import data, in the format ((tlr,trc),(brr,brc))

# Notes
* sheetName and sheetPos can not be given together
* if both sheetName and sheetPos are not specified data from the first sheet is returned
* ranges is defined using integer positions for both rows and columns
* the dictionary is keyed by the values of the cells in the first row specified in the range, or first row if `range` is not given

# Examples
```julia
julia> outDic  = ods2dic("spreadsheet.ods";sheetPos=3,range=((2,2),(6,4)))
Dict{Any,Any} with 3 entries:
  "c" => Any[33.0,43.0,53.0,63.0]
  "b" => Any[32.0,42.0,52.0,62.0]
  "d" => Any[34.0,44.0,54.0,64.0]
```
"""
function ods2dic(filename;sheetName=nothing,sheetPos=nothing,range=nothing)
    sheetsNames_h = (sheetName == nothing ? []: [sheetName])
    sheetsPos_h = (sheetPos == nothing ? []: [sheetPos])
    ranges_h = (range == nothing ? []: [range])
    dictDict = ods2dics(filename;sheetsNames=sheetsNames_h,sheetsPos=sheetsPos_h,ranges=ranges_h)
    for (k,v) in dictDict
       return v # only one value should be present
    end
end



"""
    ods2dfs(filename; <keyword arguments>)

Return a dictionary of dataframes indexed by position or name in the orifinal OpenDocument Spreadsheet (.ods) file.

# Arguments
* `sheetsNames=[]`: the list of sheet names from which to import data.
* `sheetsPos=[]`: the list of sheet positions (starting from 1) from which to import data.
* `ranges=[]`: a list of pair of touples defining the ranges in each sheet from which to import data, in the format ((tlr,trc),(brr,brc))

# Notes
* sheetsNames and sheetsPos can not be given together
* ranges is defined using integer positions for both rows and columns
* dataframes are keyed by the values of the cells in the first row specified in the range, or first row if `range` is not given
* this function requires the package `DataFrames`

# Examples
```julia
julia> outDic  = ods2dfs("spreadsheet.ods";sheetsNames=["Sheet2","Sheet3"],ranges=[((1,1),(3,3)),((2,2),(6,4))])
Dict{Any,Any} with 2 entries:
  3 => 4×3 DataFrames.DataFrame…
  1 => 2×3 DataFrames.DataFrame…
```
"""
function ods2dfs(filename;sheetsNames=[],sheetsPos=[],ranges=[])
    if !dfPackIsInstalled
        error("To use the function ods2dfs you need to have the DataFrames module installed. Run 'Pkg.add(DataFrame)' to install the DataFrames package.")
    end
    dictDict = ods2dics(filename;sheetsNames=sheetsNames,sheetsPos=sheetsPos,ranges=ranges)
    dicToReturn = Dict()
    for (k,v) in dictDict
       dicToReturn[k] = DataFrame(v)
    end
    return dicToReturn
end


"""
    ods2df(filename; <keyword arguments>)

Return a `DataFrame` from a sheet (or range within a sheet) in a OpenDocument Spreadsheet (.ods) file..

# Arguments
* `sheetName=nothing`: the sheet name from which to import data.
* `sheetPos=nothing`: the position of the sheet (starting from 1) from which to import data.
* `ranges=[]`: a pair of touples defining the range in the sheet from which to import data, in the format ((tlr,trc),(brr,brc))

# Notes
* sheetName and sheetPos can not be given together
* if both sheetName and sheetPos are not specified data from the first sheet is returned
* ranges is defined using integer positions for both rows and columns
* the dataframe is keyed by the values of the cells in the first row specified in the range, or first row if `range` is not given
* this function requires the package `DataFrames`

# Examples
```julia
julia> outDic = ods2df("spreadsheet.ods";sheetName="Sheet2")
2×3 DataFrames.DataFrame
│ Row │ a    │ b    │ c    │
├─────┼──────┼──────┼──────┤
│ 1   │ 21.0 │ 22.0 │ 23.0 │
│ 2   │ 31.0 │ 32.0 │ 33.0 │

```
"""
function ods2df(filename;sheetName=nothing,sheetPos=nothing,range=nothing)
    if !dfPackIsInstalled
        error("To use the function ods2df you need to have the DataFrames module installed. Run 'Pkg.add(DataFrame)' to install the DataFrames package.")
    end
    return DataFrame(ods2dic(filename;sheetName=sheetName,sheetPos=sheetPos,range=range))
end


"""
    odsio_test()

Provide a test to check that both the Julia 'OdsIO' and Python 'ezodf' modules are correctly installed.

"""
function odsio_test()
  try
    @pyimport ezodf
  catch
    error("The OdsIO module is correctly installed, but your python installation is missing the 'ezodf' module.")
  end
  println("Congratulations, both the Julia 'OdsIO' and Python 'ezodf' modules are correctly installed, you can start using them !")
end


"""
    odsio_test()

Check that the module compiles and the PyCall dependency is respected (it doesn't however check for python ezodf presence)

"""
function odsio_autotest()
  return 1
end

end # module OdsIO

# __precompile__()

module OdsIO

export ods_readall, ods_read, odsio_test, odsio_autotest
using PyCall

dfPackIsInstalled = true
try
    using DataFrames
catch
    dfPackIsInstalled = false
end


"""
    ods_readall(filename; <keyword arguments>)

Return a dictionary of tables|dictionaries|dataframes indexed by position or name in the original OpenDocument Spreadsheet (.ods) file.

# Arguments
* `sheetsNames=[]`: the list of sheet names from which to import data.
* `sheetsPos=[]`: the list of sheet positions (starting from 1) from which to import data.
* `ranges=[]`: a list of pair of touples defining the ranges in each sheet from which to import data, in the format ((tlr,trc),(brr,brc))
* `innerType="Matrix"`: the type of the inner container returned. Either "Matrix", "Dict" or "DataFrame"

# Notes
* sheetsNames and sheetsPos can not be given together
* ranges is defined using integer positions for both rows and columns
* individual dictionaries or dataframes are keyed by the values of the cells in the first row specified in the range, or first row if `range` is not given
* innerType="Matrix", differently from innerType="Dict", preserves original column order, it is faster and require less memory
* using innerType="DataFrame" requires the package `DataFrames` and also preserves original column order

# Examples
```julia
julia> outDic  = ods_readall("spreadsheet.ods";sheetsPos=[1,3],ranges=[((1,1),(3,3)),((2,2),(6,4))], innerType="Dict")
Dict{Any,Any} with 2 entries:
  3 => Dict{Any,Any}(Pair{Any,Any}("c",Any[33.0,43.0,53.0,63.0]),Pair{Any,Any}("b",Any[32.0,42.0,52.0,62.0]),Pair{Any,Any}("d",Any[34.0,44.0,54.…
  1 => Dict{Any,Any}(Pair{Any,Any}("c",Any[23.0,33.0]),Pair{Any,Any}("b",Any[22.0,32.0]),Pair{Any,Any}("a",Any[21.0,31.0]))
```
"""
function ods_readall(filename::AbstractString;sheetsNames::AbstractVector=String[],sheetsPos::AbstractVector=Int64[],ranges::AbstractVector=[],innerType::AbstractString="Matrix")

    try
       @pyimport ezodf
    catch
      error("The OdsIO module is correctly installed, but your python installation is missing the 'ezodf' module.")
    end
    @pyimport ezodf
    toReturn = Dict() # The outer container is always a dictionary
    try
      global doc = ezodf.opendoc(filename)
    catch
      error("I can not open for reading file $filename at $(pwd())")
    end

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
                    r_max     = min(ranges[sheetsCounter][2][1],sheet[:nrows]())
                    c_min     = ranges[sheetsCounter][1][2]
                    c_max     = min(ranges[sheetsCounter][2][2],sheet[:ncols]())
                end
            catch
                error("There is a problem with the range. Range should be defined as a list of pair of touples ((tlr,trc),(brr,brc)) for each sheet to import, using integer positions." )
            end
            if innerType=="Dict"
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
            elseif (innerType=="Matrix" || innerType=="DataFrame")
                innerMatrix = Array{Any,2}(r_max-r_min+1,c_max-c_min+1)
                r=1
                for (i, row) in enumerate(sheet[:rows]())
                    if (i>=r_min && i <= r_max) # data row
                        c=1
                        for (j, cell) in enumerate(row)

                            if (j>=c_min && j<=c_max)
                                innerMatrix[[r],[c]]=cell[:value]
                                c = c+1
                            end
                        end
                        r = r+1
                    end
                end
                if innerType=="Matrix"
                    toReturnKeyType == "name"? toReturn[sheet[:name]] = innerMatrix : toReturn[is] = innerMatrix
                else
                    if !dfPackIsInstalled
                        error("To use the function ods2dfs you need to have the DataFrames module installed. Run 'Pkg.add(DataFrame)' to install the DataFrames package.")
                    end
                    toReturnKeyType == "name"? toReturn[sheet[:name]] =   DataFrame(Any[@view innerMatrix[2:end, i] for i in 1:size(innerMatrix, 2)], Symbol.(innerMatrix[1, :])) : toReturn[is] = DataFrame(Any[@view innerMatrix[2:end, i] for i in 1:size(innerMatrix, 2)], Symbol.(innerMatrix[1, :]))
                end # innerType is really a df
            else # end innerTpe is a Dict check
                error("Only 'Matrix', 'Dict' or 'DataFrame' are supported as innerType/retType.'")
            end # end innerTpe is a Dict or Matrix check
        end # end check is a sheet to retain
    end # for each sheet
    return toReturn
end # end functionSS


"""
    ods_read(filename; <keyword arguments>)

Return a  table|dictionary|dataframe from a sheet (or range within a sheet) in a OpenDocument Spreadsheet (.ods) file..

# Arguments
* `sheetName=nothing`: the sheet name from which to import data.
* `sheetPos=nothing`: the position of the sheet (starting from 1) from which to import data.
* `ranges=[]`: a pair of touples defining the range in the sheet from which to import data, in the format ((tlr,trc),(brr,brc))
* `retType="Matrix"`: the type of container returned. Either "Matrix", "Dict" or "DataFrame"

# Notes
* sheetName and sheetPos can not be given together
* if both sheetName and sheetPos are not specified data from the first sheet is returned
* ranges is defined using integer positions for both rows and columns
* the dictionary or dataframe is keyed by the values of the cells in the first row specified in the range, or first row if `range` is not given
* retType="Matrix", differently from innerType="Dict", preserves original column order, it is faster and require less memory
* using retType="DataFrame" requires the package `DataFrames` and also preserves original column order

# Examples
```julia
julia> df = ods_read("spreadsheet.ods";sheetName="Sheet2",retType="DataFrame")
3×3 DataFrames.DataFrame
│ Row │ x1   │ x2   │ x3   │
├─────┼──────┼──────┼──────┤
│ 1   │ "a"  │ "b"  │ "c"  │
│ 2   │ 21.0 │ 22.0 │ 23.0 │
│ 3   │ 31.0 │ 32.0 │ 33.0 │
```
"""
function ods_read(filename;sheetName=nothing,sheetPos=nothing,range=nothing, retType="Matrix")
    sheetsNames_h = (sheetName == nothing ? []: [sheetName])
    sheetsPos_h = (sheetPos == nothing ? []: [sheetPos])
    ranges_h = (range == nothing ? []: [range])
    dict = ods_readall(filename;sheetsNames=sheetsNames_h,sheetsPos=sheetsPos_h,ranges=ranges_h,innerType=retType)
    for (k,v) in dict
       return v # only one value should be present
    end
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

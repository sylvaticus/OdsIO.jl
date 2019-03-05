module OdsIO

export ods_readall, ods_read, ods_write, odsio_autotest, toDf!, toDf
using PyCall, DataFrames, DataStructures, Missings #, BinDeps


# This to allow precompilation
# Unlike the @pyimport macro, this does not define a Julia module and members cannot be accessed with s.name.
# @see https://github.com/JuliaPy/PyCall.jl/issues/328
const  ezodf = PyNULL()

function __init__()
   #@BinDeps.load_dependencies
    try
        pyimport("ezodf")
    catch
        error("The OdsIO module is correctly installed, but your python installation is missing the 'ezodf' module.")
    end
    copy!(ezodf, pyimport("ezodf"))
end

"""
   ods_write(filename,data)

Write tabular data (2D Array, DataFrame or Dictionary) to OpenDocument spreadsheet format.

# Arguments
* `filename`:    an existing ods file or the one to create.
* `data=Dict()`: a dictionary of locations in the files where to export the data => the actual data (see notes).

# Notes:
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

# Examples
```julia
julia> ods_write("TestSpreadsheet.ods",Dict(("TestSheet",3,2)=>[[1,2,3,4,5] [6,7,8,9,10]]))
```
"""
function ods_write(filename::AbstractString, data::Any)
    #try
    #    @pyimport ezodf
    #catch
    #    error("The OdsIO module is correctly installed, but your python installation is missing the 'ezodf' module.")
    #end
    #@pyimport ezodf
    nSheetsOrig = 0
    if isfile(filename)
        doc = ezodf.opendoc(filename)
        if doc.doctype == "ods"
            destDoc = doc
            nSheetsOrig = length(doc.sheets)
        else
            error("Trying to write to existing file $filename , but it is not an Opendocument spreadsheet.")
        end
    else
        # new document
        destDoc = ezodf.newdoc(doctype="ods", filename=filename)
    end

    # Checking data is a dictionary
    if ! isa(data,Dict)
        error("The data parameter must be a dictionary of location_where_to_export_the_data -> exported_data. Type `?ods_write` for more informations.")
    end
    # Looping over each item to be exported
    for (k,v) in data
        # Checking k is a tuple of three elements
        if isa(k, Tuple) && length(k) == 3 && (isa(k[1], String) || isa(k[1], Int64)) && isa(k[2], Int) && isa(k[3], Int)
            # pass
        else
            error("Each key of the dictionary must be a 3-elements tuple where the first one is the sheet name or position, and the second and third ones are respectivelly row and column index of the top-left cell where to paste the data (1-index based).")
        end
        # Converting DataFrames and OrderedDicts to matrices..
        if isa(v, DataFrame)
            v = vcat(reshape(names(v), 1, length(names(v))),convert(Matrix{Any}, v))
        elseif isa(v, OrderedDict)
            v = hcat([append!(Any[k],convert(Array{Any},v2))  for(k,v2) in v] ...)
        elseif isa(v, Array) && ndims(v) == 2
            # pass
        else
            error("Data is of unknow type. Only 2D Arrays (Matrices), dataframes and ordered dictionaries are supported.")
        end
        sRSize = k[2] + size(v)[1] -1; sCSize = k[3] + size(v)[2] -1;
        sheet = ezodf.Table()
        newsheet = false
        if isa(k[1],Int)
            if k[1] > nSheetsOrig
                error("You specified a sheet position that is bigger than the number of sheet in the destination ods file. Use sheet names to add a new sheet.")
            end
            sheet = get(destDoc.sheets,k[1]-1) # new Pycall 1.9/2.0 API 1-based index, in place of old but more readable `sheet = destDoc.sheets[k[1]]`
        else
            try
                sheet = destDoc.sheets.__getitem__(k[1])
            catch
                # this is a new sheet
                sheet = ezodf.Sheet(k[1], size=(sRSize, sCSize))
                destDoc.sheets.__iadd__(sheet)
            end
        end
        # adding empty rows/cols to fit with the new data
        if sheet.nrows()<sRSize
            sheet.append_rows(max(0,sRSize-sheet.nrows()))
        end
        if sheet.ncols()<sCSize
            sheet.append_columns(max(0,sCSize-sheet.ncols())) # adding empty rows to suit the new data
        end

        for r in range(1, length=size(v)[1])
            r2 = k[2]+r-1
            for c in range(1, length=size(v)[2])
                c2 = k[3] + c -1
                if ismissing(v[r,c]) || v[r,c]==nothing
                  emptyCell = ezodf.Cell()
                  dcell = get(sheet,(r2-1,c2-1)) # Pycall 1.9 update (moving from 0 based to 1 based)
                  dcell = emptyCell
                else
                   dcell = get(sheet,(r2-1,c2-1)) # Pycall 1.9 update (moving from 0 based to 1 based)
                   dcell.set_value(v[r,c])
                end
            end
        end
    end # end for each (k,v) in data
    destDoc.backup = false
    destDoc.save()
end

"""
    ods_readall(filename; <keyword arguments>)

Return a dictionary of tables|dictionaries|dataframes indexed by position or name in the original OpenDocument Spreadsheet (.ods) file.

# Arguments
* `sheetsNames=[]`: the list of sheet names from which to import data.
* `sheetsPos=[]`: the list of sheet positions (starting from 1) from which to import data.
* `ranges=[]`: a list of pair of touples defining the ranges in each sheet from which to import data, in the format ((tlr,tlc),(brr,brc))
* `innerType="Matrix"`: the type of the inner container returned. Either "Matrix", "Dict" or "DataFrame"

# Notes
* sheetsNames and sheetsPos can not be given together
* ranges is defined using integer positions for both rows and columns
* individual dictionaries or dataframes are keyed by the values of the cells in the first row specified in the range, or first row if `range` is not given
* innerType="Matrix", differently from innerType="Dict", preserves original column order, it is faster and require less memory
* using innerType="DataFrame" also preserves original column order and try to auto-convert column types (working for Int64, Float64, String, in that order)

# Examples
```julia
julia> outDic  = ods_readall("spreadsheet.ods";sheetsPos=[1,3],ranges=[((1,1),(3,3)),((2,2),(6,4))], innerType="Dict")
Dict{Any,Any} with 2 entries:
  3 => Dict{Any,Any}(Pair{Any,Any}("c",Any[33.0,43.0,53.0,63.0]),Pair{Any,Any}("b",Any[32.0,42.0,52.0,62.0]),Pair{Any,Any}("d",Any[34.0,44.0,54.…
  1 => Dict{Any,Any}(Pair{Any,Any}("c",Any[23.0,33.0]),Pair{Any,Any}("b",Any[22.0,32.0]),Pair{Any,Any}("a",Any[21.0,31.0]))
```
"""
function ods_readall(filename::AbstractString;sheetsNames::AbstractVector=String[],sheetsPos::AbstractVector=Int64[],ranges::AbstractVector=Tuple{Tuple{Int64,Int64},Tuple{Int64,Int64}}[],innerType::AbstractString="Matrix")

    #try
    #   @pyimport ezodf
    #catch
    #  error("The OdsIO module is correctly installed, but your python installation is missing the 'ezodf' module.")
    #end
    #@pyimport ezodf
    toReturn = Dict() # The outer container is always a dictionary
    try
      global doc = ezodf.opendoc(filename)
    catch
      error("I can not open for reading file $filename at $(pwd())")
    end

    nsheets = length(doc.sheets)
    toReturnKeyType = "name"
    if !isempty(sheetsNames) && !isempty(sheetsPos)
        error("Do not use sheetNames and sheetPos together")
    end
    if !isempty(sheetsPos)
        toReturnKeyType = "pos"
    end
    sheetsCounter=0

    for (is, sheet) in enumerate(doc.sheets)
        if is in sheetsPos || sheet.name in sheetsNames || (isempty(sheetsNames) && isempty(sheetsPos))
            sheetsCounter += 1
            r_min = 1
            r_max = sheet.nrows()
            c_min = 1
            c_max = sheet.ncols()
            try
                if !isempty(ranges) && !isempty(ranges[sheetsCounter])
                    r_min::Int64     = ranges[sheetsCounter][1][1]
                    r_max::Int64     = min(ranges[sheetsCounter][2][1],sheet.nrows())
                    c_min::Int64     = ranges[sheetsCounter][1][2]
                    c_max::Int64     = min(ranges[sheetsCounter][2][2],sheet.ncols())
                else
                    # ezodf module include also empty final rows/cols in nrows()/ncols()
                    # the following code adjust r_max and c_max as to exclude empty final rows/cols if
                    # these have not been manually specified (i.e., no corrections if manually specified)

                    # Checking empty final rows..
                    emptyFinalRows = 0
                    for i = r_max-1:-1:0
                        row = sheet.row(i)
                        allEmpty = true
                        for (j, cell) in enumerate(row)
                            if cell.value != nothing
                                allEmpty = false
                                break
                            end
                        end
                        if(!allEmpty)
                            break
                        else
                            emptyFinalRows += 1
                        end
                    end
                    r_max -= emptyFinalRows
                    # Checking empty final cols..
                    emptyFinalCols = 0
                    for i = c_max-1:-1:0
                        col = sheet.column(i)
                        allEmpty = true
                        for (j, cell) in enumerate(col)
                            if cell.value != nothing
                                allEmpty = false
                                break
                            end
                        end
                        if(!allEmpty)
                            break
                        else
                            emptyFinalCols += 1
                        end
                    end
                    c_max -= emptyFinalCols
                end
            catch
                error("There is a problem with the range. Range should be defined as a list of pair of touples ((tlr,tlc),(brr,brc)) for each sheet to import, using integer positions." )
            end
            if (innerType=="Matrix" || innerType=="Dict" || innerType=="DataFrame" )
                innerMatrix = Array{Any,2}(undef,r_max-r_min+1,c_max-c_min+1)
                r::Int64=1
                for (i::Int64, row) in enumerate(sheet.rows())
                    if (i>=r_min && i <= r_max) # data row
                        c::Int64=1
                        for (j::Int64, cell) in enumerate(row)
                            if (j>=c_min && j<=c_max)
                                # Try saving the value as integer if that's actually possible
                                if typeof(cell.value) <: Number
                                    if isinteger(cell.value)
                                        innerMatrix[[r],[c]] .= convert(Int64,cell.value)
                                    else
                                        innerMatrix[[r],[c]] .= cell.value
                                    end
                                else
                                    innerMatrix[[r],[c]] .= cell.value
                                end
                                c = c+1
                            end
                        end
                        r = r+1
                    end
                end
                if innerType=="Matrix"
                    toReturnKeyType == "name" ? toReturn[sheet.name] = innerMatrix : toReturn[is] = innerMatrix
                elseif innerType == "Dict"
                    toReturnKeyType == "name" ? toReturn[sheet.name] = Dict([(ch,innerMatrix[2:end,cix]) for (cix::Int64,ch) in enumerate(innerMatrix[1,:])]) : toReturn[is] = Dict([(ch,innerMatrix[2:end,cix]) for (cix,ch) in enumerate(innerMatrix[1,:])])
                elseif innerType == "DataFrame"
                    df = toDf!(innerMatrix)
                    toReturnKeyType == "name" ? toReturn[sheet.name] =   df : toReturn[is] = df
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
* `range=nothing`: a pair of touples defining the range in the sheet from which to import data, in the format ((tlr,tlc),(brr,brc))
* `retType="Matrix"`: the type of container returned. Either "Matrix", "Dict" or "DataFrame"

# Notes
* sheetName and sheetPos can not be given together
* if both sheetName and sheetPos are not specified data from the first sheet is returned
* ranges is defined using integer positions for both rows and columns
* the dictionary or dataframe is keyed by the values of the cells in the first row specified in the range, or first row if `range` is not given
* retType="Matrix", differently from innerType="Dict", preserves original column order, it is faster and require less memory
* using innerType="DataFrame" also preserves original column order and try to auto-convert column types (working for Int64, Float64, String, in that order)

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
function ods_read(filename::AbstractString; sheetName=nothing, sheetPos=nothing, range=nothing, retType::AbstractString="Matrix")
    sheetsNames_h = (sheetName == nothing ? [] : [sheetName])
    sheetsPos_h = (sheetPos == nothing ? [] : [sheetPos])
    ranges_h = (range == nothing ? [] : [range])
    dict = ods_readall(filename;sheetsNames=sheetsNames_h,sheetsPos=sheetsPos_h,ranges=ranges_h,innerType=retType)
    for (k,v) in dict
       return v # only one value should be present
    end
end

"""
    odsio_autotest()

Check that the module compiles and the PyCall dependency is respected (it doesn't however check for python ezodf presence)

"""
function odsio_autotest()
  return 1
end


"""
    toDf!(m)

Convert a mixed-type Matrix to DataFrame
"""
function toDf!(m)
    [m[i,j]==nothing ? m[i,j]=missing : m[i,j] for i in 2:size(m)[1], j in 1:size(m)[2]]
    return DataFrame([[m[2:end,i]...] for i in 1:size(m,2)], Symbol.(m[1,:]))
end

"""
    toDf(m)

Convert a mixed-type Matrix to DataFrame
"""
function toDf(m)
    m2 = [m[i,j]==nothing ? missing : m[i,j] for i in 1:size(m)[1], j in 1:size(m)[2]]
    return DataFrame([[m2[2:end,i]...] for i in 1:size(m2,2)], Symbol.(m2[1,:]))
end


end # module OdsIO

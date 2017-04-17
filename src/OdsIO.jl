# __precompile__()

module OdsIO

export ods_readall, ods_read, odsio_test, odsio_autotest
using PyCall
using DataFrames

"""
   ods_write(filename,data)

Write tabular data (2d Array, DataFrame or Dictionary) to OpenDocument spreadsheet format.

# Arguments
* `filename`:    an existing ods file or the one to create.
* `data=Dict()`: a dictionary of locations in the files where to export the data => the actual data (see notes).

# Notes:
* The locations where to save the data (the keys in the dictionary) are a tuple of tree elements:
  The first one is the sheet name or sheet position, the other two are the index of row and column of the top
  left corner where to export the data.
  If using sheet positions, these must be within current file boundaries. If you want to create new sheets,
  use names.
* The actual data exported are either a Matrix (2D Array), a DataFrame or a Dictionary. In case of DataFrame or
  Dictionary the headers ARE exported, so if you don't want them first convert the DataFrame (or Dictionary)
  to a Matrix.

"""

function ods_write(filename::AbstractString, data::Union{
    Dict{Tuple{String,Int64,Int64},Array{Any,2}},
    Dict{Tuple{Int64,Int64,Int64},Array{Any,2}},
    Dict{Tuple{String,Int64,Int64},DataFrames.DataFrame},
    Dict{Tuple{Int64,Int64,Int64},DataFrames.DataFrame},
    Dict{Tuple{String,Int64,Int64},Dict{Any}},
    Dict{Tuple{Int64,Int64,Int64},Dict{Any}},
    })
    try
        @pyimport ezodf
    catch
        error("The OdsIO module is correctly installed, but your python installation is missing the 'ezodf' module.")
    end
    @pyimport ezodf

    if isfile(filename)
        doc = ezodf.opendoc(filename)
        if doc[:doctype] == "ods"
            destDoc = doc
        else
            error("Trying to write to existing file $filename , but it is not an Opendocument spreadsheet.")
        end
    else
        destDoc = ezodf.newdoc(doctype="ods", filename=filename)
    end

    sheet = ezodf.Sheet("SHEET", size=(10, 10))
    push!(destDoc[:sheets],sheet)
    sheet["A1"].set_value("cell with text")
    sheet["B2"].set_value(3.141592)
    destDoc.save()

end

anarray =[[1,2,3] [4,5,6]]
adf = DataFrame(test = [1,2,3], pippo=[2,3,4])
test = Dict(("asheet",1,3) => adf, ("asheet2",1,3) => adf)
ods_write("newspreadsheet.ods",test)

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
* using innerType="DataFrame" also preserves original column order

# Examples
```julia
julia> outDic  = ods_readall("spreadsheet.ods";sheetsPos=[1,3],ranges=[((1,1),(3,3)),((2,2),(6,4))], innerType="Dict")
Dict{Any,Any} with 2 entries:
  3 => Dict{Any,Any}(Pair{Any,Any}("c",Any[33.0,43.0,53.0,63.0]),Pair{Any,Any}("b",Any[32.0,42.0,52.0,62.0]),Pair{Any,Any}("d",Any[34.0,44.0,54.…
  1 => Dict{Any,Any}(Pair{Any,Any}("c",Any[23.0,33.0]),Pair{Any,Any}("b",Any[22.0,32.0]),Pair{Any,Any}("a",Any[21.0,31.0]))
```
"""
function ods_readall(filename::AbstractString;sheetsNames::AbstractVector=String[],sheetsPos::AbstractVector=Int64[],ranges::AbstractVector=Tuple{Tuple{Int64,Int64},Tuple{Int64,Int64}}[],innerType::AbstractString="Matrix")

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
                    r_min::Int64     = ranges[sheetsCounter][1][1]
                    r_max::Int64     = min(ranges[sheetsCounter][2][1],sheet[:nrows]())
                    c_min::Int64     = ranges[sheetsCounter][1][2]
                    c_max::Int64     = min(ranges[sheetsCounter][2][2],sheet[:ncols]())
                end
            catch
                error("There is a problem with the range. Range should be defined as a list of pair of touples ((tlr,trc),(brr,brc)) for each sheet to import, using integer positions." )
            end
            if (innerType=="Matrix" || innerType=="Dict" || innerType=="DataFrame" )
                innerMatrix = Array{Any,2}(r_max-r_min+1,c_max-c_min+1)
                r::Int64=1
                for (i::Int64, row) in enumerate(sheet[:rows]())
                    if (i>=r_min && i <= r_max) # data row
                        c::Int64=1
                        for (j::Int64, cell) in enumerate(row)

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
                elseif innerType == "Dict"
                    toReturnKeyType == "name"? toReturn[sheet[:name]] = Dict([(ch,innerMatrix[2:end,cix]) for (cix::Int64,ch) in enumerate(innerMatrix[1,:])]) : toReturn[is] = Dict([(ch,innerMatrix[2:end,cix]) for (cix,ch) in enumerate(innerMatrix[1,:])])
                elseif innerType == "DataFrame"
                    toReturnKeyType == "name"? toReturn[sheet[:name]] =   DataFrame(Any[@view innerMatrix[2:end, i] for i::Int64 in 1:size(innerMatrix, 2)], Symbol.(innerMatrix[1, :])) : toReturn[is] = DataFrame(Any[@view innerMatrix[2:end, i] for i in 1:size(innerMatrix, 2)], Symbol.(innerMatrix[1, :]))
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
* using retType="DataFrame" also preserves original column order

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

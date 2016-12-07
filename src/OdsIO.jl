__precompile__()

module OdsIO
using PyCall

dfPackIsInstalled = true
try
    using DataFrames
catch
    dfPackIsInstalled = false
end


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


function ods2dic(filename;sheetName=nothing,sheetPos=nothing,range=nothing)
    sheetsNames_h = (sheetName == nothing ? []: [sheetName])
    sheetsPos_h = (sheetPos == nothing ? []: [sheetPos])
    ranges_h = (range == nothing ? []: [range])
    dictDict = ods2dics(filename;sheetsNames=sheetsNames_h,sheetsPos=sheetsPos_h,ranges=ranges_h)
    for (k,v) in dictDict
       return v # only one value should be present
    end
end

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

function ods2df(filename;sheetName=nothing,sheetPos=nothing,range=nothing)
    if !dfPackIsInstalled
        error("To use the function ods2df you need to have the DataFrames module installed. Run 'Pkg.add(DataFrame)' to install the DataFrames package.")
    end
    return DataFrame(ods2dic(filename;sheetName=sheetName,sheetPos=sheetPos,range=range))
end

function test()
  println("Congratulations, OdsIO module is correctly installed !")
end


end # module OdsIO

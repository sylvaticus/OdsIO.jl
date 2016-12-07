workspace()
using PyCall
using DataFrames
@pyimport ezodf
doc = ezodf.opendoc("spreadsheet.ods")
nsheets = length(doc[:sheets])
println("Spreadsheet contains $nsheets sheet(s).")
for sheet in doc[:sheets]
    println("---------")
    println("   Sheet name : $(sheet[:name])")
    println("Size of Sheet : (rows=$(sheet[:nrows]()), cols=$(sheet[:ncols]()))")
end

# convert the first sheet to a pandas.DataFrame
sheet = doc[:sheets][1]
df_dict = Dict()
col_index = Dict()

test= sheet[:rows]()
for (i, row) in enumerate(test)
  println(i)
  print([cell[:value] for cell in row])
end
for (i, row) in enumerate(sheet[:rows]())
  # row is a list of cells
  # assume the header is on the first row
  if i == 1
      # columns as lists in a dictionary
      [df_dict[cell[:value]] = [] for cell in row]
      # create index for the column headers
      [col_index[j]=cell[:value]  for (j, cell) in enumerate(row)]
      continue
  end
  for (j, cell) in enumerate(row)
      # use header instead of column index
      append!(df_dict[col_index[j]],cell[:value])
  end
end
# and convert to a DataFrame
df = DataFrame(df_dict)




"""
# convert the first sheet to a pandas.DataFrame
sheet = doc.sheets[0]
df_dict = {}
for i, row in enumerate(sheet.rows()):
    # row is a list of cells
    # assume the header is on the first row
    if i == 0:
        # columns as lists in a dictionary
        df_dict = {cell.value:[] for cell in row}
        # create index for the column headers
        col_index = {j:cell.value for j, cell in enumerate(row)}
        continue
    for j, cell in enumerate(row):
        # use header instead of column index
        df_dict[col_index[j]].append(cell.value)
# and convert to a DataFrame
df = pd.DataFrame(df_dict)
"""

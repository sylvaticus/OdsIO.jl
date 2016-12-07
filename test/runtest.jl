workspace()

include("../src/OdsIO.jl")

OdsIO.test()

outDic1  = OdsIO.ods2dics("spreadsheet.ods";sheetsPos=[1,3],ranges=[((1,1),(3,3)),((2,2),(6,4))])
outDic2  = OdsIO.ods2dics("spreadsheet.ods";sheetsNames=["Sheet1","Sheet3"],ranges=[(),((2,2),(6,4))])
outDic3  = OdsIO.ods2dics("spreadsheet.ods";ranges=[(),(),((2,2),(6,4))])
outDic4  = OdsIO.ods2dics("spreadsheet.ods")
outDic5  = OdsIO.ods2dics("spreadsheet.ods";sheetsNames=["Sheet3"],ranges=[((2,1),(9,7))])
outDic6  =  OdsIO.ods2dic("spreadsheet.ods";sheetName="Sheet2")
outDir7  =  OdsIO.ods2dic("spreadsheet.ods";sheetPos=3,range=((2,2),(6,4)))
outDir8  =  OdsIO.ods2dic("spreadsheet.ods")
outDic9  = OdsIO.ods2dfs("spreadsheet.ods";sheetsPos=[1,3],ranges=[((1,1),(3,3)),((2,2),(6,4))])
outDic11 =  OdsIO.ods2df("spreadsheet.ods";sheetName="Sheet2")
println("Done!")

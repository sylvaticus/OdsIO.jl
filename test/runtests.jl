using OdsIO
using Base.Test




#odsio_test()

#outDic1  = ods2dics("spreadsheet.ods";sheetsPos=[1,3],ranges=[((1,1),(3,3)),((2,2),(6,4))])
#outDic2  = ods2dics("spreadsheet.ods";sheetsNames=["Sheet1","Sheet3"],ranges=[(),((2,2),(6,4))])
#outDic3  = ods2dics("spreadsheet.ods";ranges=[(),(),((2,2),(6,4))])
#outDic4  = ods2dics("spreadsheet.ods")
#outDic5  = ods2dics("spreadsheet.ods";sheetsNames=["Sheet3"],ranges=[((2,1),(9,7))])
#outDic6  = ods2dic("spreadsheet.ods";sheetName="Sheet2")
#outDir7  = ods2dic("spreadsheet.ods";sheetPos=3,range=((2,2),(6,4)))
#outDir8  = ods2dic("spreadsheet.ods")
#outDic9  = ods2dfs("spreadsheet.ods";sheetsPos=[1,3],ranges=[((1,1),(3,3)),((2,2),(6,4))])
#outDic11 = ods2df("spreadsheet.ods";sheetName="Sheet2")
#println("Done!")


# write your own tests here
@test 1 == odsio_autotest()

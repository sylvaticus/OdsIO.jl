using OdsIO, DataFrames
using Base.Test

# TEST 1: Just that everything is installed
@test 1 == odsio_autotest()

# Test 2: Testing that what is written with ods_write then is read back with ods_read
testData = [[1,2,3,4,5] [6,7,8,9,10]]
filename = "testspreadheet"
ods_write(filename,Dict(("TestSheet",3,2)=>testData))
testDataOut = ods_read(filename;sheetName="TestSheet",range=((3,2),(7,3)))
rm(filename)
@test convert(Array{Any,2}, testData) == testDataOut

# Test 3: Write/read test of dataframes with missing values
dfIn = DataFrame(A = [1,2,3,missing],
               B = [1.1,2.2,missing,3.3],
               C = ["a",missing,"c","d"],
               D = [1,2,3,4])
filename = "testspreadsheet"
ods_write(filename,Dict(("TestSheetDf",1,1)=>dfIn))
dfOut = ods_read(filename; sheetName="TestSheetDf", retType="DataFrame")
rm(filename)
#@test dfIn == dfOut


# Test 3: Fake test, this should not pass
# @test 1 == 2

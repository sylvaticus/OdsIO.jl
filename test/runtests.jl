using OdsIO, DataFrames
using Test

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

# @test dfIn == dfOut #would not pass because dataframe report the equal test as "missing" not as "true" because of the propagation of missing values
# I need to check each single column time and value
# ismissing(dfIn == dfOut) report true even when columns are differern types

# Checking size
@test (size(dfIn) == size(dfOut))
# Checking col name and types
typesCheck = [typeof(dfIn[i]) == typeof(dfOut[i]) for i in 1:size(dfIn)[2] ]
@test minimum(typesCheck)
@test names(dfIn) == names(dfOut)
# Checking actual values
global areequals = true
for i in 1:size(dfIn)[1]
    for j in 1:size(dfIn)[2]
        if (ismissing(dfIn[i,j]) && ismissing(dfOut[i,j]))
        elseif (ismissing(dfIn[i,j]) && ! ismissing(dfOut[i,j]))
            global areequals = false
        elseif (! ismissing(dfIn[i,j]) && ismissing(dfOut[i,j]))
            global areequals = false
        elseif (dfIn[i,j] != dfOut[i,j])
            global areequals = false
        end
    end
end
@test areequals

# Test 3: Fake test, this should not pass
# @test 1 == 2

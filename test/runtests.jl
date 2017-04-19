using OdsIO
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

Dim objuft

Set objuft=CreateObject("QuickTest.Application")
objuft.visible=True
objuft.launch
objuft.open("C:\Users\sfjbs\Desktop\OP_M5_F\Driver\OP_M5_driver")
objuft.Test.Run
objuft.Test.Close
objuft.quit
set objuft=nothing
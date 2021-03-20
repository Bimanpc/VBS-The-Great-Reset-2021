Set oFSO = CreateObject("Scripting.FileSystemObject")
sFile1 = "MyComputerInfo.csv"
Set oFile1 = oFSO.CreateTextFile(sFile1, 1)
strQuery = "SELECT Family,Manufacturer,NumberOfCores FROM Win32_Processor"
Set colResults = GetObject("winmgmts://./root/cimv2").ExecQuery( strQuery )

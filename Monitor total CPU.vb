Set objWMIService = GetObject("winmgmts:\\localhost\root\CIMV2") 
Set CPUInfo = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor",,48) 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile("C:\Logs.log", 8, True, 0)
For Each Item in CPUInfo 
    Wscript.Echo "Win32_PerfFormattedData_PerfOS_Processor instance"
    objLogFile.WriteLine "Win32_PerfFormattedData_PerfOS_Processor instance"
    Wscript.Echo "PercentProcessorTime: " & Item.PercentProcessorTime
    objLogFile.WriteLine "PercentProcessorTime: " & Item.PercentProcessorTime

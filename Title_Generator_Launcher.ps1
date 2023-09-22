$argumentlist = "/c powershell.exe -file `"$PSScriptRoot\mainScript.ps1`" -param1 `"paramstring`""
Start-Process cmd.exe -WindowStyle Hidden -ArgumentList $argumentlist
Set objShell = CreateObject("WScript.Shell")
Set objExec = objShell.Exec("wmic path softwarelicensingservice get OA3xOriginalProductKey")
strOutput = objExec.StdOut.ReadAll
strProductKey = Trim(Replace(strOutput, "OA3xOriginalProductKey", ""))
MsgBox "Windows Product Key: " & strProductKey, vbInformation, "GetKey"

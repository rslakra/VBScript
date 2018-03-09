HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Internet Explorer\Desktop\SafeMode\Components"
strValue = "0"
ValueName = "DeskHtmlVersion"
objReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
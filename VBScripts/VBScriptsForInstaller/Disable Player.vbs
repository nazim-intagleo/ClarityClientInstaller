

const HKEY_CURRENT_USER = &H80000001
strComputer = "."

Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
strComputer & "\root\default:StdRegProv")

'Define byte's array
Dim bArray

strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"

'Changing value to disable startup application
bArray= Array(03,00,00,00,00,00,00,00,00,00,00,00)

objReg.SetBinaryValue HKEY_CURRENT_USER, strKeyPath, "DMPlayer", bArray

Set objReg = Nothing


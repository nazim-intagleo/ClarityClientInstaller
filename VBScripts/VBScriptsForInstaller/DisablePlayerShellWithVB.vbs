




Set oShellApp = CreateObject("Shell.Application")
Set oWshShell = CreateObject( "WScript.Shell" )

Const Q = """"


' Before executing explorer.exe we MUST set the "shell" (REG_SZ) value to be explorer.exe, or else when Windows Explorer loads,
' then it will read this regval and thing we don't need a desktop, since a different shell is being used.
'
' 'shell' value replacement idea from here: http://social.msdn.microsoft.com/Forums/en-US/windowsgeneraldevelopmentissues/thread/49c05fe8-0ec3-4e1b-9d11-8d893cdea11c/
' 
Call oWshShell.RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\shell", "explorer.exe", "REG_SZ")

Call oShellApp.ShellExecute("explorer.exe", " uac", "", "runas", 1)


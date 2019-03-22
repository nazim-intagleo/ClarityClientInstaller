 


Set fso = CreateObject("Scripting.FileSystemObject")
sourceDir = "C:\Program Files\Allure\Clarity\"
If NOT (fso.FolderExists(sourceDir)) Then
	sourceDir = "C:\Program Files (x86)\Allure\Clarity\"
END If
destDir = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\"

sourceFile = fso.GetAbsolutePathName(sourceDir & "ScriptForStartUp.vbs")
destFolder = fso.GetAbsolutePathName(destDir)

Set objShell = CreateObject("Shell.Application")
objShell.NameSpace(destFolder).copyHere sourceFile

Set fso = Nothing
Set objShell = Nothing

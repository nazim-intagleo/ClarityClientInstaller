

Set fso = CreateObject("Scripting.FileSystemObject")
sourceDir = "C:\Program Files\Allure\Clarity\"
If NOT (fso.FolderExists(sourceDir)) Then
	sourceDir = "C:\Program Files (x86)\Allure\Clarity\"
END If
destDir = "C:\ClarityClientInstaller\"

fso.CopyFile sourceDir & "StartPlayer.vbs", destDir, True

set fso = nothing

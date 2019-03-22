


Set fso = CreateObject("Scripting.FileSystemObject")

ZipFile="C:\Program Files\Allure\Clarity\ClarityClientWebApp.1.5.37.zip"
If NOT (fso.FolderExists(ZipFile)) Then
	zipFile = "C:\Program Files (x86)\Allure\Clarity\ClarityClientWebApp.1.5.37.zip"
END If

ExtractTo="C:\ClarityClientInstaller"
If NOT (fso.FolderExists(ExtractTo)) Then
	fso.CreateFolder("C:\ClarityClientInstaller")
END If

sourceFile = fso.GetAbsolutePathName(ZipFile)
destFolder = fso.GetAbsolutePathName(ExtractTo)

Set objShell = CreateObject("Shell.Application")
Set FilesInZip=objShell.NameSpace(sourceFile).Items()
objShell.NameSpace(destFolder).copyHere FilesInZip, 16
 
Set fso = Nothing
Set objShell = Nothing
Set FilesInZip = Nothing


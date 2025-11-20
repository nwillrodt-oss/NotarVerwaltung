Dim objFSO ' FileSystemObject

Const AppName = "Notarverwaltung"
Const ExeName="NotarVerwaltung.exe"
Const DevFolder ="D:\_OLG\Entwicklung\VB6\NotarVerwaltung\"
Const SetupPath ="Setup\Notarverwaltung\Express\Custom\DiskImages\DISK1"

Const ReleaseFolder ="_Releases\"
Const ThisReleasFolderPrefix ="Release "
Const SetupFolder ="Setup\"
Const SourcenFolder = "Sourcen\"


'Version Ermitteln
Set objFSO = CreateObject("Scripting.FileSystemObject")
szVersion = objFSO.GetFileVersion("NotarVerwaltung.exe")							' Version ermitteln

ThisRelaseFolder = DevFolder & ReleaseFolder & ThisReleasFolderPrefix & szVersion & "\"
If objFSO.FolderExists(DevFolder & ReleaseFolder) Then
'	msgbox DevFolder & ReleaseFolder & " exist"
	If objFSO.FolderExists(ThisRelaseFolder) Then
		If msgbox("Der ordner für die Version " & szVersion & " existiert bereits. Mochten sie Abbrechen?", vbOKCancel ,"Weitermachen?") =2 Then
			bCancel =true
		end if
	else
		objFSO.CreateFolder(ThisRelaseFolder)										' Release Ordner anlegen
	end if
end if

if NOT bCancel Then
	If objFSO.FolderExists(ThisRelaseFolder) Then
		bCreateReleaseFolderOK = True 
		'If Not FolderExists(hisRelaseFolder & SetupFolder) Then
			objFSO.CreateFolder(ThisRelaseFolder & SetupFolder)						' ReleaseSetup Folder anlegen
		'end if
		'If Not FolderExists(hisRelaseFolder & SourcenFolder) Then
			objFSO.CreateFolder(ThisRelaseFolder & SourcenFolder)					' Release Sourcen Folder anlegen
		'end if
	end if

	If objFSO.FolderExists(ThisRelaseFolder & SetupFolder) Then
		bCreateSetupFolderOK = True 												' Prüfen
	end if

	If objFSO.FolderExists(ThisRelaseFolder & SourcenFolder) Then
		bCreateSourceFolderOK = True 												' Prüfen
	end if

	If  bCreateSetupFolderOK and objFSO.FolderExists(DevFolder & SetupPath) Then	' Setup Kopieren
		objFSO.CopyFolder DevFolder & SetupPath, ThisRelaseFolder & SetupFolder
	end if
	
	If bCreateSourceFolderOK And objFSO.FolderExists(DevFolder) Then
		' Fast alle ordner Kopieren
		Set oFolder = objFSO.GetFolder(DevFolder)
		For each subFolder in oFolder.SubFolders
			If Not UCASE(subFolder.Name) = "_RELEASES" AND Not ucase(subFolder.name) = "ABLAGE" Then
				objFSO.CopyFolder subFolder.Path, ThisRelaseFolder & SourcenFolder
			end if
		next
		
		For each oFile in oFolder.files
			objFSO.copyFile oFile.path, ThisRelaseFolder & SourcenFolder, true
		next
	end if
end if
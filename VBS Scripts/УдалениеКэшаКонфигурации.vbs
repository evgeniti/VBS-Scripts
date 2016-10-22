On Error Resume Next

Set fso      = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.Shell")

Const Path_APPDATA = "%APPDATA%\1C\1Cv82"
Const Path_USERPROFILE = "%USERPROFILE%\Local Settings\Application Data\1C\1Cv82"
Const Path_LOCALAPPDATA = "%LOCALAPPDATA%\1C\1Cv82"

Const Path_APPDATA_83 = "%APPDATA%\1C\1Cv8"
Const Path_USERPROFILE_83 = "%USERPROFILE%\Local Settings\Application Data\1C\1Cv8"
Const Path_LOCALAPPDATA_83 = "%LOCALAPPDATA%\1C\1Cv8"

' Delete 82 version
DeleteSubfolders(objShell.expandenviromentstrings(Path_APPDATA))
DeleteSubfolders(objShell.expandenviromentstrings(Path_USERPROFILE))
DeleteSubfolders(objShell.expandenviromentstrings(Path_LOCALAPPDATA))
If not (fso.folderexists("%APPDATA%\1C\1Cv82\tmplts")) Then
    fso.createfolder "%APPDATA%\1C\1Cv82\tmplts"
end If

' Delete 83 version
DeleteSubfolders(objShell.expandenviromentstrings(Path_APPDATA_83))
DeleteSubfolders(objShell.expandenviromentstrings(Path_USERPROFILE_83))
DeleteSubfolders(objShell.expandenviromentstrings(Path_LOCALAPPDATA_83))
If not (fso.folderexists("%APPDATA%\1C\1Cv8\tmplts")) Then
    fso.createfolder "%APPDATA%\1C\1Cv8\tmplts"
end If

Wscript.Echo "Done"

' functions --------------

Sub DeleteSubfolders(InDir)
    For Each tFolder In fso.GetFolder(InDir).SubFolders
        If Len(tFolder.Name) > 20 Then fso.DeleteFolder(tFolder.Path)
    Next
End Sub
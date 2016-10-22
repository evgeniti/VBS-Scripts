' file download over ADO Connection. bulkadmin role required

Sub ExtractFileOverSQL(FilePath,FileToSave,ServerName,UserName,Password) 
	Const adSaveCreateOverWrite=2
	Const adTypeBinary = 1
	Dim ADOConnection, ADOCommand, RS, stm
	Set ADOConnection=CreateObject("ADODB.Connection")
	If (UserName = "") and (Password = "") Then
		ADOConnection.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;Integrated Security=SSPI;Initial Catalog=AdminDB;Data Source="&ServerName
	Else
		ADOConnection.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID="&UserName&";Password="&Password&";Initial Catalog=AdminDB;Data Source="&ServerName
	End If
	ADOConnection.CursorLocation = 3
	ADOConnection.ConnectionTimeout = 120
	ADOConnection.CommandTimeout = 3600
	On Error Resume Next
	ADOConnection.Open()
	If Err.Number <> 0 Then
		WScript.StdErr.WriteLine CStr(Now)&vbTab&"Code: "&CStr(Err.Number)&vbTab&"Description: "&Err.Description
		Exit sub
	End If
	On Error Goto 0
	
	ADOConnection.Execute("set nocount on set dateformat ymd")
	Set RS = ADOConnection.Execute("SELECT bulkColumn FROM OPENROWSET(BULK N'"& FilePath &"', SINGLE_BLOB) rs")
	If Not RS.EOF Then
		Set stm = CreateObject("ADODB.Stream")
		With stm
			.Open
			.Type = adTypeBinary
			.Write RS.Fields("bulkColumn").Value
			.SaveToFile FileToSave, adSaveCreateOverWrite
		End With
		stm.Close()
		Set stm = Nothing
	End If
	RS.Close()
	Set RS = Nothing

	ADOConnection.Close()
	
	Set ADOConnection = Nothing
	
End sub 

Call ExtractFileOverSQL(".bak",".bak","server","user","prolll")
